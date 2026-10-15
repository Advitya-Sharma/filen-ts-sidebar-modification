// Pure sizing, encoding and routing helpers for the client-side thumbnail generators (video, pdf, svg) — the
// categories the Rust SDK cannot decode and this app therefore still draws to a canvas itself. Zero
// dependencies of its own, so it can be imported from anywhere without risking a cycle.

// Aspect-preserving max-dimension fit, never upscaling — a source already at or under maxDim in both
// dimensions returns its own size unchanged (scale clamped to 1).
export function fitWithin(width: number, height: number, maxDim: number): { width: number; height: number } {
	const scale = Math.min(1, maxDim / Math.max(width, height))

	return {
		width: Math.round(width * scale),
		height: Math.round(height * scale)
	}
}

// Where a drive video thumbnail looks for its frame: one second in (half the video when it is shorter),
// and once more a few seconds in only when that frame is black. Kept near the start on purpose: a
// container without an index (Matroska without Cues, a fragmented MP4) can only reach a later position
// by reading every byte before it.
export function driveVideoThumbSeeks(duration: number): { first: number; retry: number | null } {
	const first = Math.min(1, duration / 2)
	const retry = Math.min(duration / 2, 5)

	return { first, retry: retry > first ? retry : null }
}

// What canPlayType is asked for a content type. QuickTime and M4V are MP4 underneath: Chrome and Firefox
// play them by sniffing but answer their own types with nothing.
const CAN_PLAY_QUERY: ReadonlyMap<string, string> = new Map([
	["video/quicktime", "video/mp4"],
	["video/x-m4v", "video/mp4"]
])

export function canPlayQuery(contentType: string): string {
	return CAN_PLAY_QUERY.get(contentType) ?? contentType
}

// The content type a drive video thumbnail's hidden <video> is served with, or null when that path is
// not worth a read: no inline-allowed type, a container the browser says it cannot play (MKV on Chrome
// and Safari), or a file whose element already failed to decode. Those go straight to the keyframe read.
// Decided before any read, so a file the element cannot play costs none on it.
export function driveVideoElementType(
	contentType: string | null,
	canPlay: (query: string) => boolean,
	undecodable: boolean
): string | null {
	if (contentType === null || undecodable || !canPlay(canPlayQuery(contentType))) {
		return null
	}

	return contentType
}

// HTMLMediaElement error codes that are verdicts about bytes that are always readable (a local file):
// MEDIA_ERR_DECODE and MEDIA_ERR_SRC_NOT_SUPPORTED, unlike an aborted or failed fetch. Over a stream that
// can fail, MEDIA_ERR_SRC_NOT_SUPPORTED is ambiguous (classifyMediaError).
export function isFormatMediaError(code: number | undefined): boolean {
	return code === 3 || code === 4
}

// A canvas silently downgrades an unsupported requested encode type to image/png instead of throwing
// (WHATWG canvas spec) — the produced blob's own .type is the only reliable signal a webp attempt
// actually produced webp, so this keeps it when honored and otherwise awaits the caller's jpeg
// fallback attempt.
export async function selectWebpOrFallback(webpAttempt: Blob, jpegFallback: () => Promise<Blob>): Promise<Blob> {
	return webpAttempt.type === "image/webp" ? webpAttempt : await jpegFallback()
}

const JPEG_FALLBACK_QUALITY = 0.85

// OffscreenCanvas.convertToBlob and HTMLCanvasElement.toBlob are different APIs (promise-returning
// vs. callback-based) — convertToBlob's presence, which HTMLCanvasElement never has, is what
// TypeScript narrows the union on below, so no runtime `instanceof OffscreenCanvas` (unavailable
// outside a browser/worker realm, e.g. under a plain Node unit test) is ever needed here. A type the canvas
// cannot encode comes back as PNG (see selectWebpOrFallback).
export async function attemptEncode(canvas: OffscreenCanvas | HTMLCanvasElement, type: string, quality?: number): Promise<Blob> {
	if ("convertToBlob" in canvas) {
		// exactOptionalPropertyTypes forbids an explicit `quality: undefined` property, so the options
		// object is built conditionally rather than always including the (possibly absent) key.
		return await (quality === undefined ? canvas.convertToBlob({ type }) : canvas.convertToBlob({ type, quality }))
	}

	return await new Promise<Blob>((resolve, reject) => {
		canvas.toBlob(
			blob => {
				if (blob === null) {
					reject(new Error("canvas toBlob produced no result"))

					return
				}

				resolve(blob)
			},
			type,
			quality
		)
	})
}

// webp-first, jpeg-0.85 fallback — the on-disk format policy the canvas-drawing generators share. The
// SDK arm needs none of this: it encodes webp itself, inside wasm. A drawing with transparency (an
// svg) falls back to png instead, since jpeg has no alpha and would paint it black.
export async function encodeCanvasThumb(canvas: OffscreenCanvas | HTMLCanvasElement, options: { alpha?: boolean } = {}): Promise<Blob> {
	const webp = await attemptEncode(canvas, "image/webp")

	return await selectWebpOrFallback(webp, () =>
		options.alpha === true ? attemptEncode(canvas, "image/png") : attemptEncode(canvas, "image/jpeg", JPEG_FALLBACK_QUALITY)
	)
}
