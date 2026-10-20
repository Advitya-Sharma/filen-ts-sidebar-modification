import { bytesSource, type ByteSource } from "../../core/opc/io"
import type { Slicing } from "../../core/slicing"
import type { SheetHost, SheetSessionApi } from "../../core/sheet/protocol"
import type { SheetFeature } from "../../core/sheet/registry"
import { isTextKind, openTextSession, type TextSheetDocument } from "../../core/sheet/text"
import { saveAsCellsFeature } from "../../core/sheet/save/cells"
import { serveSheetWorker, type SheetOpener, type SheetWorkerScope } from "./worker"

// Text files (CSV, TSV, .tab) on their own: a host and a worker that load neither the workbook reader and writer
// nor the formula engine, for an app that opens a text file in a worker of its own.

async function wholeSource(source: ByteSource): Promise<Uint8Array> {
	return source.size === 0 ? new Uint8Array(0) : source.read(0, source.size)
}

// Opens a text file; any other kind is not one. Its cells can be handed out for Save as .xlsx (saveAs.cells).
export function textOpener(features: readonly SheetFeature<TextSheetDocument>[] = []): SheetOpener {
	const all = [saveAsCellsFeature, ...features]

	return async (source, options, onEvent, track) => {
		if (!isTextKind(options.kind)) return { kind: "failed", reason: "unsupported", message: "Not a text file" }

		const opened = await openTextSession(await wholeSource(source), options, onEvent, all)

		if (opened.kind === "opened") track(opened.session)

		return opened
	}
}

// Text files in a worker: the worker's whole entry is serveTextWorker(scope).
export function serveTextWorker(scope: SheetWorkerScope, features: readonly SheetFeature<TextSheetDocument>[] = []): void {
	serveSheetWorker(scope, textOpener(features))
}

// The sessions an in-thread host opened, closed with it.
export function trackedSessions(): { readonly track: (session: SheetSessionApi) => void; dispose(): void; readonly disposed: boolean } {
	const sessions = new Set<SheetSessionApi>()
	let disposed = false

	return {
		get disposed() {
			return disposed
		},
		track(session) {
			if (disposed) {
				session.close()

				return
			}

			const close = session.close.bind(session)

			sessions.add(session)
			session.close = () => {
				sessions.delete(session)
				close()
			}
		},
		dispose() {
			disposed = true

			for (const session of [...sessions]) session.close()
		}
	}
}

// Text files on the calling thread.
export function inThreadTextHost(features: readonly SheetFeature<TextSheetDocument>[] = [], slicing?: Slicing): SheetHost {
	const sessions = trackedSessions()
	const open = textOpener(features)

	return {
		async open(source, options, onEvent) {
			if (sessions.disposed) return { kind: "failed", reason: "failed", message: "The host was disposed" }

			try {
				return await open(
					source instanceof Uint8Array ? bytesSource(source) : source,
					slicing === undefined || options.slicing !== undefined ? options : { ...options, slicing },
					onEvent,
					sessions.track
				)
			} catch (error) {
				return { kind: "failed", reason: "failed", message: error instanceof Error ? error.message : String(error) }
			}
		},
		dispose() {
			sessions.dispose()
		}
	}
}

export { serveSheetWorker, type SheetOpener, type SheetWorkerScope } from "./worker"
