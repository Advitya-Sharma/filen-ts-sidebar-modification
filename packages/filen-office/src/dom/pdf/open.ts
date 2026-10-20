import { PdfWorkerUnavailableError, passwordReason, type PdfPasswordReason } from "./errors"
import type { PdfDocument, PdfDocumentParams, PdfJs, PdfRangeTransport } from "./types"

// Opens a document with the host's pdf.js module, worker and assets. The security-relevant parameters are set
// here, in one place, never by a caller:
// - enableXfa off: XFA is a second, scriptable form engine.
// - no docBaseUrl: a document's catalog can override it, and relative URLs inside the document resolve
//   against it.
// - no password option: passwords answer onPassword, so they never sit in the parameters.
// - verbosity ERRORS: pdf.js warnings quote document content.
// Scripting is the annotation layer's own opt-in and is never enabled; pdf.js 6 has no eval path.

export type PdfSource =
	// The whole file. Copied: pdf.js transfers the buffer it is given to its worker, detaching it.
	| { readonly kind: "bytes"; readonly bytes: Uint8Array }
	// Read on demand (a file larger than one should hold twice, or bytes behind a bridge).
	| {
			readonly kind: "ranges"
			readonly length: number
			readonly read: (offset: number, length: number) => Promise<Uint8Array>
			// The most one read asks for; pdf.js merges adjacent chunks into one request of any size.
			readonly maxRead?: number
			readonly chunkSize?: number
	  }

export type PdfWorker =
	// A dedicated Worker per document, from pdf.js's worker script (a same-origin URL).
	| { readonly kind: "script"; readonly src: string }
	// pdf.js's main-thread worker, for a page that cannot start one (an opaque origin): the host imports
	// pdf.worker.mjs statically, which installs globalThis.pdfjsWorker.
	| { readonly kind: "inThread" }

export type PdfAssetKind = "cMapUrl" | "standardFontDataUrl" | "wasmUrl"

export type PdfAssets =
	// pdf.js's defaults: system fonts for the standard 14, no cMaps or wasm decoders.
	| { readonly kind: "none" }
	// Directories pdf.js fetches from, each ending in "/".
	| {
			readonly kind: "urls"
			readonly cMapUrl?: string
			readonly standardFontDataUrl?: string
			readonly wasmUrl?: string
			readonly iccUrl?: string
	  }
	// Bytes handed over directly, nothing fetched (a page whose origin cannot fetch). The bundled standard fonts
	// are used instead of the platform's, so documents lay out alike everywhere.
	| { readonly kind: "inline"; readonly read: (kind: PdfAssetKind, filename: string) => Promise<Uint8Array> }

export interface PdfOpenOptions {
	readonly worker: PdfWorker
	readonly assets?: PdfAssets
	// Largest image decoded, in pixels; larger ones are skipped. Unset: no limit.
	readonly maxImageSize?: number
	// Largest canvas pdf.js allocates itself, in bytes. Unset: pdf.js's default.
	readonly canvasMaxAreaInBytes?: number
	// pdf.js asks for a password: answer through PdfOpening.password (pdf.js waits meanwhile).
	readonly onPassword: (reason: PdfPasswordReason) => void
	// A ranged read failed after the document opened (pdf.js never asks for a failed range again).
	readonly onReadError?: () => void
}

export interface PdfOpening<S = unknown> {
	readonly promise: Promise<PdfDocument<S>>
	// Answers the pending password request on the same task, which already holds the file. False when none is
	// pending.
	password(password: string): boolean
	destroy(): Promise<void>
}

// Thrown into PdfOpening.promise when a ranged read fails before the document opened.
export class PdfReadError extends Error {
	override readonly name = "PdfReadError"

	constructor() {
		super("range read failed")
	}
}

const VERBOSITY_ERRORS = 0
const DEFAULT_CHUNK = 256 * 1024

function hasInThreadWorker(): boolean {
	const worker: unknown = Reflect.get(globalThis, "pdfjsWorker")

	return typeof worker === "object" && worker !== null && typeof Reflect.get(worker, "WorkerMessageHandler") === "function"
}

export function pdfDocumentParams(
	source: PdfSource,
	assets: PdfAssets,
	options: Pick<PdfOpenOptions, "maxImageSize" | "canvasMaxAreaInBytes">
): PdfDocumentParams {
	const params: { -readonly [K in keyof PdfDocumentParams]: PdfDocumentParams[K] } = { enableXfa: false, verbosity: VERBOSITY_ERRORS }

	if (source.kind === "bytes") params.data = source.bytes.slice()
	else {
		// Both, or neither takes effect: pdf.js would otherwise read ahead through the whole file.
		params.disableStream = true
		params.disableAutoFetch = true
		params.rangeChunkSize = source.chunkSize ?? DEFAULT_CHUNK
	}

	if (assets.kind === "urls") {
		if (assets.cMapUrl !== undefined) {
			params.cMapUrl = assets.cMapUrl
			params.cMapPacked = true
		}

		if (assets.standardFontDataUrl !== undefined) params.standardFontDataUrl = assets.standardFontDataUrl
		if (assets.wasmUrl !== undefined) params.wasmUrl = assets.wasmUrl
		if (assets.iccUrl !== undefined) params.iccUrl = assets.iccUrl
	} else if (assets.kind === "inline") {
		const read = assets.read

		// pdf.js constructs it with the asset URLs (unused: nothing is fetched) and calls fetch().
		params.BinaryDataFactory = class {
			fetch({ kind, filename }: { readonly kind: PdfAssetKind; readonly filename: string }): Promise<Uint8Array> {
				return read(kind, filename)
			}
		}
		// Pinned, never inferred: pdf.js infers it from whether its asset URLs look fetchable, which differs
		// between a development server and a release build.
		params.useWorkerFetch = false
		params.useSystemFonts = false
	}

	if (options.maxImageSize !== undefined) params.maxImageSize = options.maxImageSize
	if (options.canvasMaxAreaInBytes !== undefined) params.canvasMaxAreaInBytes = options.canvasMaxAreaInBytes

	return params
}

export function openPdf<S>(pdfjs: PdfJs<S>, source: PdfSource, options: PdfOpenOptions): PdfOpening<S> {
	if (options.worker.kind === "inThread") {
		if (!hasInThreadWorker()) {
			const failed = Promise.reject(new PdfWorkerUnavailableError())

			return { promise: failed, password: () => false, destroy: () => Promise.resolve() }
		}
	} else if (pdfjs.GlobalWorkerOptions.workerSrc !== options.worker.src) {
		pdfjs.GlobalWorkerOptions.workerSrc = options.worker.src
	}

	let destroyed = false
	let opened = false
	let failRead: ((error: Error) => void) | null = null
	const readFailed = new Promise<never>((_, reject) => {
		failRead = reject
	})

	let range: PdfRangeTransport | undefined

	if (source.kind === "ranges") {
		const transport = new pdfjs.PDFDataRangeTransport(source.length, null)
		const maxRead = Math.max(1, source.maxRead ?? Number.MAX_SAFE_INTEGER)

		transport.requestDataRange = (begin: number, end: number): void => {
			const fetchRange = async (): Promise<void> => {
				const total = end - begin

				if (total <= 0) return

				const merged = new Uint8Array(total)

				for (let offset = begin; offset < end;) {
					const chunk = await source.read(offset, Math.min(maxRead, end - offset))

					if (chunk.length === 0 || chunk.length > end - offset) throw new PdfReadError()

					merged.set(chunk, offset - begin)
					offset += chunk.length
				}

				if (!destroyed) transport.onDataRange(begin, merged)
			}

			fetchRange().catch(() => {
				if (destroyed) return

				if (opened) options.onReadError?.()
				else failRead?.(new PdfReadError())
			})
		}

		range = transport
	}

	const params = pdfDocumentParams(source, options.assets ?? { kind: "none" }, options)
	const call: PdfDocumentParams = range === undefined ? params : { ...params, range }
	const task = pdfjs.getDocument(call)
	let pending: ((password: string) => void) | null = null

	task.onPassword = (update: (password: string) => void, reason: unknown): void => {
		pending = update
		options.onPassword(passwordReason(reason))
	}

	const promise = Promise.race([task.promise, readFailed]).then(document => {
		opened = true

		return document
	})

	// A failure before anyone awaits (a destroyed task) must not surface as unhandled.
	readFailed.catch(() => undefined)

	return {
		promise,
		password(password: string): boolean {
			const update = pending

			if (update === null || destroyed) return false

			pending = null
			update(password)

			return true
		},
		destroy(): Promise<void> {
			destroyed = true
			pending = null

			return task.destroy()
		}
	}
}
