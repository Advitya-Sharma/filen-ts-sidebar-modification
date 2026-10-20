import { bytesSource, type ByteSource } from "../../core/opc/io"
import type {
	SheetAbortSignal,
	SheetEvent,
	SheetLockedFile,
	SheetOpenOptions,
	SheetOpenResult,
	SheetSessionApi
} from "../../core/sheet/protocol"
import { transferables, type FromSheetWorker, type SheetWorkerMethod, type ToSheetWorker } from "./messages"

// The worker's side: opens, unlocks and runs sessions for the page, any number at a time. The file's bytes, the
// key of a password-protected one and its decrypted package stay here; the page gets tiles, layouts, verdicts and
// (sealed, for a protected workbook) saved bytes.

// The worker's end of its channel to the page (adapt `self`).
export interface SheetWorkerScope {
	post(message: unknown, transfer: readonly ArrayBuffer[]): void
	listen(onMessage: (data: unknown) => void): void
}

// Opens a file where it is held: workbooks and text files (host.ts workbookOpener), or text files alone (text.ts),
// so a worker for text loads neither the workbook reader nor the formula engine. `track` hears every session
// opened, a locked file's unlock included.
export type SheetOpener = (
	source: ByteSource,
	options: SheetOpenOptions,
	onEvent: (event: SheetEvent) => void,
	track: (session: SheetSessionApi) => void
) => Promise<SheetOpenResult>

interface Open {
	session: SheetSessionApi | null
	locked: SheetLockedFile | null
	readonly opening: { aborted: boolean }
	readonly aborts: Map<number, { aborted: boolean }>
	closed: boolean
}

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === "object" && value !== null
}

function messageOf(error: unknown): string {
	return error instanceof Error ? error.message : String(error)
}

export function serveSheetWorker(scope: SheetWorkerScope, opener: SheetOpener): void {
	const opens = new Map<number, Open>()
	const post = (message: FromSheetWorker, transfer: readonly ArrayBuffer[] = []): void => {
		scope.post(message, transfer)
	}
	const event =
		(s: number) =>
		(sheetEvent: SheetEvent): void => {
			if (opens.get(s)?.closed === false) post({ t: "event", s, event: sheetEvent }, transferables(sheetEvent))
		}
	const signalOf = (open: Open, id: number): SheetAbortSignal => {
		const own = { aborted: false }

		open.aborts.set(id, own)

		return {
			get aborted() {
				return own.aborted || open.closed
			}
		}
	}
	const call = async (open: Open, id: number, method: SheetWorkerMethod, args: readonly unknown[]): Promise<unknown> => {
		const session = open.session

		if (session === null) throw new Error("The workbook is not open")

		const signal = signalOf(open, id)

		try {
			switch (method) {
				case "tiles":
					return await session.tiles(args[0] as number, args[1] as Uint32Array)
				case "layout":
					return await session.layout(args[0] as number)
				case "command":
					return await session.command(args[0] as Parameters<SheetSessionApi["command"]>[0])
				case "undo":
					return await session.undo()
				case "redo":
					return await session.redo()
				case "query":
					return await session.query(args[0] as Parameters<SheetSessionApi["query"]>[0], args[1] as never, { signal })
				case "intent":
					return await session.intent()
				case "save":
					return await session.save({ signal })
				case "export":
					return await session.export(args[0] as "xlsx" | "csv", args[1] as { readonly sheet?: number } | undefined)
				case "markSaved":
					return await session.markSaved(args[0] as number)
			}
		} finally {
			open.aborts.delete(id)
		}
	}
	const close = (s: number): void => {
		const open = opens.get(s)

		if (open === undefined) return

		opens.delete(s)
		open.closed = true
		open.opening.aborted = true
		open.session?.close()
		open.locked?.dispose()
	}

	scope.listen(data => {
		if (!isRecord(data) || typeof data["s"] !== "number") return

		const message = data as ToSheetWorker
		const s = message.s

		switch (message.t) {
			case "open": {
				if (opens.has(s)) return

				const entry: Open = { session: null, locked: null, opening: { aborted: false }, aborts: new Map(), closed: false }

				opens.set(s, entry)
				void opener(bytesSource(message.bytes), { ...message.options, signal: entry.opening }, event(s), session => {
					if (entry.closed) session.close()
					else entry.session = session
				}).then(
					result => {
						if (entry.closed) {
							if (result.kind === "needsPassword") result.locked.dispose()

							return
						}

						switch (result.kind) {
							case "opened":
								post({ t: "opened", s, result: { kind: "opened", info: result.info } }, transferables(result.info))
								break
							case "needsPassword":
								entry.locked = result.locked
								post({
									t: "opened",
									s,
									result: { kind: "needsPassword", scheme: result.locked.scheme, spinCount: result.locked.spinCount }
								})
								break
							default:
								opens.delete(s)
								post({ t: "opened", s, result })
						}
					},
					(error: unknown) => {
						opens.delete(s)
						post({ t: "opened", s, result: { kind: "failed", reason: "failed", message: messageOf(error) } })
					}
				)
				break
			}
			case "abortOpen": {
				const open = opens.get(s)

				if (open !== undefined) open.opening.aborted = true
				break
			}
			case "unlock": {
				const open = opens.get(s)
				const locked = open?.locked

				if (open === undefined || locked === null || locked === undefined) {
					post({ t: "unlocked", s, id: message.id, result: { ok: false, reason: "aborted" } })

					return
				}

				const signal = signalOf(open, message.id)

				void locked.unlock(message.password, { signal }).then(
					result => {
						open.aborts.delete(message.id)

						if (!result.ok) {
							post({ t: "unlocked", s, id: message.id, result })

							return
						}

						if (open.closed) {
							result.session.close()

							return
						}

						open.session = result.session
						open.locked = null
						post({ t: "unlocked", s, id: message.id, result: { ok: true, info: result.info } }, transferables(result.info))
					},
					() => {
						open.aborts.delete(message.id)
						post({ t: "unlocked", s, id: message.id, result: { ok: false, reason: "damaged" } })
					}
				)
				break
			}
			case "abortUnlock":
			case "abort": {
				const signal = opens.get(s)?.aborts.get(message.id)

				if (signal !== undefined) signal.aborted = true
				break
			}
			case "dispose":
				close(s)
				break
			case "call": {
				const open = opens.get(s)

				if (open === undefined) {
					post({ t: "result", s, id: message.id, error: "The workbook is closed" })

					return
				}

				void call(open, message.id, message.method, message.args).then(
					value => post({ t: "result", s, id: message.id, value }, transferables(value)),
					(error: unknown) => post({ t: "result", s, id: message.id, error: messageOf(error) })
				)
				break
			}
			case "notice": {
				const session = opens.get(s)?.session

				switch (message.notice) {
					case "release":
						session?.release(message.args[0] as Float64Array)
						break
					case "viewOnly":
						session?.viewOnly()
						break
					case "configure":
						session?.configure(message.args[0] as Parameters<SheetSessionApi["configure"]>[0])
						break
					case "close":
						close(s)
						break
				}
				break
			}
		}
	})
}
