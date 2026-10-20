import type {
	SheetAbortSignal,
	SheetEvent,
	SheetHost,
	SheetLockedFile,
	SheetOpenResult,
	SheetSessionApi,
	SheetWorkbookInfo
} from "../../core/sheet/protocol"
import type {
	FromSheetWorker,
	SheetWorkerMethod,
	SheetWorkerNotice,
	SheetWorkerOpenOptions,
	SheetWorkerUnlocked,
	ToSheetWorker
} from "./messages"

// The page's side of a sheet worker: a SheetHost whose sessions run in the worker. The file's bytes are
// transferred to it at open; answers come back in call order, events between them.

// A Worker, or anything that talks like one.
export interface SheetWorkerLike {
	postMessage(message: unknown, transfer?: ArrayBuffer[]): void
	addEventListener(type: "message" | "error", listener: (event: { readonly data?: unknown; readonly message?: string }) => void): void
	terminate?(): void
}

interface Pending {
	readonly resolve: (value: unknown) => void
	readonly reject: (error: Error) => void
}

interface Open {
	readonly onEvent: (event: SheetEvent) => void
	readonly calls: Map<number, Pending>
	opened: ((result: Extract<FromSheetWorker, { t: "opened" }>["result"]) => void) | null
	readonly unlocks: Map<number, (result: SheetWorkerUnlocked) => void>
	closed: boolean
}

const FAILED = "The sheet worker stopped"

export function sheetWorkerHost(worker: SheetWorkerLike): SheetHost {
	const opens = new Map<number, Open>()
	let next = 1
	let nextCall = 1
	let failed: string | null = null

	const send = (message: ToSheetWorker, transfer: ArrayBuffer[] = []): void => {
		worker.postMessage(message, transfer)
	}
	const fail = (message: string): void => {
		if (failed !== null) return

		failed = message

		for (const open of opens.values()) {
			for (const pending of open.calls.values()) pending.reject(new Error(message))

			open.calls.clear()
			open.opened?.({ kind: "failed", reason: "failed", message })
			open.opened = null

			for (const unlock of open.unlocks.values()) unlock({ ok: false, reason: "aborted" })

			open.unlocks.clear()

			if (!open.closed) open.onEvent({ type: "failed", message })
		}
	}

	// What waits on an open the page let go of: an unlock answers "aborted", an open "aborted".
	const settle = (open: Open): void => {
		for (const unlock of open.unlocks.values()) unlock({ ok: false, reason: "aborted" })

		open.unlocks.clear()
		open.opened?.({ kind: "aborted" })
		open.opened = null
	}

	worker.addEventListener("message", event => {
		const message = event.data as FromSheetWorker | undefined

		if (message === undefined || typeof message !== "object") return

		const open = opens.get(message.s)

		if (open === undefined) return

		switch (message.t) {
			case "opened":
				open.opened?.(message.result)
				open.opened = null
				break
			case "unlocked":
				open.unlocks.get(message.id)?.(message.result)
				open.unlocks.delete(message.id)
				break
			case "event":
				if (!open.closed) open.onEvent(message.event)
				break
			case "result": {
				const pending = open.calls.get(message.id)

				open.calls.delete(message.id)

				if (message.error !== undefined) pending?.reject(new Error(message.error))
				else pending?.resolve(message.value)
				break
			}
		}
	})
	worker.addEventListener("error", event => fail(event.message ?? FAILED))

	const session = (s: number): SheetSessionApi => {
		const call = <T>(method: SheetWorkerMethod, args: readonly unknown[], signal?: SheetAbortSignal): Promise<T> => {
			const open = opens.get(s)

			if (failed !== null) return Promise.reject(new Error(failed))
			if (open === undefined || open.closed) return Promise.reject(new Error("The session is closed"))

			const id = nextCall++

			return new Promise<T>((resolve, reject) => {
				open.calls.set(id, { resolve: value => resolve(value as T), reject })
				send({ t: "call", s, id, method, args })

				if (signal !== undefined) {
					if (signal.aborted) send({ t: "abort", s, id })
					else signal.addEventListener?.("abort", () => send({ t: "abort", s, id }))
				}
			})
		}
		const notice = (kind: SheetWorkerNotice, args: readonly unknown[] = []): void => {
			if (failed === null && opens.get(s)?.closed === false) send({ t: "notice", s, notice: kind, args })
		}

		return {
			tiles: (sheet, coordinates) => call("tiles", [sheet, coordinates]),
			release: keys => notice("release", [keys]),
			layout: sheet => call("layout", [sheet]),
			command: command => call("command", [command]),
			undo: () => call("undo", []),
			redo: () => call("redo", []),
			query: (name, request, options) => call("query", [name, request], options?.signal),
			intent: () => call("intent", []),
			viewOnly: () => notice("viewOnly"),
			configure: settings => notice("configure", [settings]),
			save: options =>
				call<Awaited<ReturnType<SheetSessionApi["save"]>>>("save", [], options?.signal).catch((error: unknown) => ({
					ok: false as const,
					reason: "closed" as const,
					message: error instanceof Error ? error.message : String(error)
				})),
			export: (format, options) => call("export", [format, options]),
			markSaved: version => call("markSaved", [version]),
			close: () => {
				const open = opens.get(s)

				if (open === undefined || open.closed) return

				notice("close")
				open.closed = true

				for (const pending of open.calls.values()) pending.reject(new Error("The session is closed"))

				opens.delete(s)
			}
		}
	}

	const lockedFile = (s: number, open: Open, scheme: "standard" | "agile", spinCount: number): SheetLockedFile => ({
		scheme,
		spinCount,
		unlock(password, options = {}) {
			if (failed !== null || open.closed) return Promise.resolve({ ok: false, reason: "aborted" })

			const id = nextCall++

			return new Promise<SheetWorkerUnlocked>(resolve => {
				open.unlocks.set(id, resolve)
				send({ t: "unlock", s, id, password })

				const signal = options.signal

				if (signal !== undefined) {
					if (signal.aborted) send({ t: "abortUnlock", s, id })
					else signal.addEventListener?.("abort", () => send({ t: "abortUnlock", s, id }))
				}
			}).then(result => (result.ok ? { ok: true, session: session(s), info: result.info as SheetWorkbookInfo } : result))
		},
		dispose() {
			if (open.closed) return

			open.closed = true
			opens.delete(s)
			settle(open)

			if (failed === null) send({ t: "dispose", s })
		}
	})

	return {
		open(source, options, onEvent) {
			if (!(source instanceof Uint8Array))
				return Promise.resolve({ kind: "failed", reason: "unsupported", message: "A worker host opens bytes, not a source" })
			if (failed !== null) return Promise.resolve({ kind: "failed", reason: "failed", message: failed })

			const s = next++
			const open: Open = { onEvent, calls: new Map(), opened: null, unlocks: new Map(), closed: false }
			const { signal, slicing: _slicing, ...rest } = options
			const sent: SheetWorkerOpenOptions = rest

			opens.set(s, open)

			return new Promise<SheetOpenResult>(resolve => {
				open.opened = result => {
					switch (result.kind) {
						case "opened":
							resolve({ kind: "opened", session: session(s), info: result.info })
							break
						case "needsPassword":
							resolve({ kind: "needsPassword", locked: lockedFile(s, open, result.scheme, result.spinCount) })
							break
						default:
							opens.delete(s)
							resolve(result)
					}
				}

				const buffer = source.buffer

				send(
					{ t: "open", s, bytes: source, options: sent },
					buffer instanceof ArrayBuffer && buffer.byteLength === source.byteLength ? [buffer] : []
				)

				if (signal !== undefined) {
					if (signal.aborted) send({ t: "abortOpen", s })
					else signal.addEventListener?.("abort", () => send({ t: "abortOpen", s }))
				}
			})
		},
		dispose() {
			for (const [s, open] of opens) {
				if (!open.closed) send({ t: "dispose", s })

				open.closed = true

				for (const pending of open.calls.values()) pending.reject(new Error("The host was disposed"))

				open.calls.clear()
				settle(open)
			}

			opens.clear()
			worker.terminate?.()
		}
	}
}
