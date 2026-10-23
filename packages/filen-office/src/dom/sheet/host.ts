import { bytesSource, type ByteSource, type InflateHook } from "../../core/opc/index"
import type { Slicing } from "../../core/slicing"
import { openXlsxSession, type XlsxSessionOpen } from "../../core/sheet/commands"
import { openOdsSession } from "../../core/ods/session"
import type {
	SheetEvent,
	SheetHost,
	SheetLockedFile,
	SheetOpenOptions,
	SheetOpenResult,
	SheetSessionApi,
	SheetSessionOptions
} from "../../core/sheet/protocol"
import type { SheetFeature } from "../../core/sheet/registry"
import type { XlsxDocument } from "../../core/xlsx/edit/document"
import { isTextKind, type TextSheetDocument } from "../../core/sheet/text"
import { saveAsXlsxFeature } from "../../core/sheet/save/xlsx"
import { textOpener, trackedSessions } from "./text"
import type { SheetOpener } from "./worker"
import { openProtectedWorkbook, sheetEncryption } from "./protection"
import type { DecryptedPackage, LockedPackage } from "../encryption/sealedPackage"

// The in-thread host: sessions run on the calling thread (mobile under Hermes, a WebView, tests), sliced so long
// work yields. No Worker, MessageChannel or OffscreenCanvas is needed. The worker server (worker.ts) opens files
// through the same `openWorkbook`.

export interface SheetHostOptions {
	// Features added to the xlsx session's own, and to a text file's.
	readonly features?: readonly SheetFeature<XlsxDocument>[]
	readonly textFeatures?: readonly SheetFeature<TextSheetDocument>[]
	// The raw-deflate decompressor an OpenDocument spreadsheet's content streams through (a native
	// DecompressionStream in a browser); fflate's without one.
	readonly inflate?: InflateHook
}

const COALESCE = 256 * 1024

// The parts of the WHATWG streams platformInflate uses (this module builds without the DOM library).
interface PlatformReader {
	read(): Promise<{ readonly done: boolean; readonly value?: Uint8Array }>
	cancel(): Promise<void>
}

interface PlatformStream {
	pipeThrough(transform: unknown): PlatformStream
	getReader(): PlatformReader
}

interface PlatformController {
	enqueue(chunk: Uint8Array): void
	close(): void
}

type PlatformStreamConstructor = new (
	source: { pull(controller: PlatformController): Promise<void>; cancel(): Promise<void> },
	strategy: { highWaterMark: number }
) => PlatformStream

// Raw deflate through the platform's DecompressionStream (browsers, workers, Node), undefined where there is none
// (Hermes): an OpenDocument spreadsheet's content then inflates through fflate. Input is pulled only as the output is
// read, so an entry outgrowing its declared size stops the inflate.
export function platformInflate(): InflateHook | undefined {
	const Decompression: unknown = Reflect.get(globalThis, "DecompressionStream")
	const Readable: unknown = Reflect.get(globalThis, "ReadableStream")

	if (typeof Decompression !== "function" || typeof Readable !== "function") return undefined

	return input => {
		const iterator = input[Symbol.asyncIterator]()
		const source = new (Readable as PlatformStreamConstructor)(
			{
				async pull(controller) {
					const next = await iterator.next()

					if (next.done === true) controller.close()
					else controller.enqueue(next.value)
				},
				async cancel() {
					await iterator.return?.()
				}
			},
			{ highWaterMark: 0 }
		)
		const output = source.pipeThrough(new (Decompression as new (format: string) => unknown)("deflate-raw"))

		// Small outputs (Firefox's are 16 KiB) are gathered into one of COALESCE bytes: each hand-over costs the
		// consumer a turn of its loop. The buffer is reused once the consumer asks for the next chunk (InflateHook).
		return (async function* () {
			const reader = output.getReader()
			const gathered = new Uint8Array(COALESCE)
			let length = 0

			try {
				for (;;) {
					const { done, value } = await reader.read()

					if (done || value === undefined) break

					if (value.length >= COALESCE) {
						if (length > 0) {
							yield gathered.subarray(0, length)
							length = 0
						}

						yield value

						continue
					}

					if (length + value.length > COALESCE) {
						yield gathered.subarray(0, length)
						length = 0
					}

					gathered.set(value, length)
					length += value.length
				}

				if (length > 0) yield gathered.subarray(0, length)
			} finally {
				await reader.cancel().catch(() => undefined)
			}
		})()
	}
}

// Slices of 8 ms between macrotask turns, when the caller gave none.
export function defaultSlicing(): Slicing {
	const timer: unknown = Reflect.get(globalThis, "setTimeout")

	return {
		sliceMs: 8,
		now: () => Date.now(),
		pause: () =>
			new Promise<void>(resolve => {
				if (typeof timer === "function") (timer as (callback: () => void, ms: number) => void)(resolve, 0)
				else resolve()
			})
	}
}

function opened(result: XlsxSessionOpen, track: (session: SheetSessionApi) => void): SheetOpenResult {
	if (result.kind !== "opened") return result

	track(result.session)

	return result
}

// Opens a decrypted package; its key goes when the session closes (or at once when it does not open).
async function openDecrypted(
	document: DecryptedPackage,
	options: SheetOpenOptions,
	onEvent: (event: SheetEvent) => void,
	features: readonly SheetFeature<XlsxDocument>[]
): Promise<XlsxSessionOpen> {
	const encryption = sheetEncryption(document)
	const sessionOptions: SheetSessionOptions = { ...options, encryption }
	let result: XlsxSessionOpen

	try {
		result = await openXlsxSession(bytesSource(document.bytes), sessionOptions, onEvent, features)
	} catch (error) {
		encryption.release()

		throw error
	}

	if (result.kind !== "opened") encryption.release()

	return result
}

function lockedFile(
	locked: LockedPackage,
	options: SheetOpenOptions,
	onEvent: (event: SheetEvent) => void,
	features: readonly SheetFeature<XlsxDocument>[],
	track: (session: SheetSessionApi) => void
): SheetLockedFile {
	let disposed = false
	let busy = false

	return {
		scheme: locked.scheme,
		spinCount: locked.spinCount,
		async unlock(password, unlockOptions = {}) {
			if (disposed || busy) return { ok: false, reason: "aborted" }

			busy = true

			try {
				const signal = unlockOptions.signal
				const result = await locked.unlock(password, {
					...(signal !== undefined && { signal }),
					...(options.slicing !== undefined && { slicing: options.slicing })
				})

				if (!result.ok) return result

				if (disposed) {
					result.document.release()

					return { ok: false, reason: "aborted" }
				}

				const session = await openDecrypted(result.document, options, onEvent, features)

				if (session.kind === "opened" && disposed) {
					session.session.close()

					return { ok: false, reason: "aborted" }
				}

				if (session.kind === "opened") {
					track(session.session)

					return { ok: true, session: session.session, info: session.info }
				}

				return { ok: false, reason: session.kind === "aborted" ? "aborted" : "damaged" }
			} finally {
				busy = false
			}
		},
		dispose() {
			disposed = true
		}
	}
}

// Opens a workbook on the thread that holds it: an encrypted package is unlocked (Excel's default password first);
// a text file opens as text (text.ts), and here, beside the workbook writer, also converts to .xlsx (saveAs.xlsx).
export async function openWorkbook(
	source: ByteSource,
	options: SheetOpenOptions,
	onEvent: (event: SheetEvent) => void,
	features: SheetHostOptions,
	track: (session: SheetSessionApi) => void
): Promise<SheetOpenResult> {
	if (isTextKind(options.kind)) return textOpener([saveAsXlsxFeature, ...(features.textFeatures ?? [])])(source, options, onEvent, track)
	if (options.signal?.aborted === true) return { kind: "aborted" }

	const workbook = features.features ?? []

	// An OpenDocument spreadsheet: its own package (an encrypted one is refused, not unlocked).
	if (options.kind === "ods" || options.kind === "ots") {
		const result = await openOdsSession(source, options, onEvent, workbook, features.inflate ?? platformInflate())

		return result.kind === "refused" ? result : opened(result, track)
	}
	const protectedOpen = await openProtectedWorkbook(source, options.signal, options.slicing)

	switch (protectedOpen.kind) {
		case "refused":
			return { kind: "refused", reason: protectedOpen.reason }
		case "aborted":
			return { kind: "aborted" }
		case "needsPassword":
			return { kind: "needsPassword", locked: lockedFile(protectedOpen.locked, options, onEvent, workbook, track) }
		case "unlocked":
			return opened(await openDecrypted(protectedOpen.document, options, onEvent, workbook), track)
		case "plain":
			return opened(await openXlsxSession(source, options, onEvent, workbook), track)
	}
}

// Every kind a sheet worker opens: serveSheetWorker(scope, workbookOpener()).
export function workbookOpener(options: SheetHostOptions = {}): SheetOpener {
	return (source, openOptions, onEvent, track) => openWorkbook(source, openOptions, onEvent, options, track)
}

export function inThreadSheetHost(hostOptions: SheetHostOptions = {}): SheetHost {
	const sessions = trackedSessions()

	return {
		async open(source, options, onEvent) {
			if (sessions.disposed) return { kind: "failed", reason: "failed", message: "The host was disposed" }

			const sliced: SheetOpenOptions = options.slicing === undefined ? { ...options, slicing: defaultSlicing() } : options

			try {
				return await openWorkbook(
					source instanceof Uint8Array ? bytesSource(source) : source,
					sliced,
					onEvent,
					hostOptions,
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
