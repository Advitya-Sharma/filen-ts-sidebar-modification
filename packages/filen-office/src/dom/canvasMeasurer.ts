import type { FontExtents, TextMeasurer } from "../core/text/index"

// TextMeasurer over a 2D canvas context: an OffscreenCanvas in the layout worker on web, or one in the
// WebView on mobile. Measure in one realm only: Firefox's <canvas> and OffscreenCanvas widths differ
// slightly. The shapes are declared here so the package needs no DOM typings.

interface TextMetricsLike {
	readonly width: number
	readonly fontBoundingBoxAscent?: number
	readonly fontBoundingBoxDescent?: number
}

export interface CanvasContextLike {
	font: string
	// Missing in WebKit, whose canvas always kerns.
	fontKerning?: string
	measureText(text: string): TextMetricsLike
}

export interface OffscreenCanvasScope {
	readonly OffscreenCanvas: new (width: number, height: number) => { getContext(type: "2d"): unknown }
}

// Font extents are read at this size and scaled: small sizes come back rounded to whole pixels.
const EXTENTS_SIZE = 1000
const SIZE = /(\d+(?:\.\d+)?)px/

export function canvasMeasurer(context: CanvasContextLike): TextMeasurer {
	const kerningControl = "fontKerning" in context && typeof context.fontKerning === "string"
	let font = ""
	let kerning = ""
	const extents = new Map<string, FontExtents>()

	const use = (css: string, kern: boolean): void => {
		// Setting the font is the expensive part of measuring; skip it when unchanged.
		if (css !== font) {
			context.font = css
			font = css
		}

		const wanted = kern ? "normal" : "none"

		if (kerningControl && wanted !== kerning) {
			context.fontKerning = wanted
			kerning = wanted
		}
	}

	return {
		kerningControl,
		measure(css, kern, texts, count, widths) {
			use(css, kern)

			for (let i = 0; i < count; i++) widths[i] = context.measureText(texts[i] ?? "").width
		},
		fontExtents(css) {
			let cached = extents.get(css)

			if (cached !== undefined) return cached

			const match = SIZE.exec(css)
			const size = match === null ? 0 : Number(match[1])

			use(match === null ? css : css.replace(SIZE, `${EXTENTS_SIZE}px`), true)

			const metrics = context.measureText("Hg")
			const scale = size > 0 ? size / EXTENTS_SIZE : 1

			cached = {
				ascent: (metrics.fontBoundingBoxAscent ?? EXTENTS_SIZE * 0.8) * scale,
				descent: (metrics.fontBoundingBoxDescent ?? EXTENTS_SIZE * 0.2) * scale
			}
			extents.set(css, cached)

			return cached
		}
	}
}

// Whether `scope` can measure with an OffscreenCanvas; without one, measure on the main thread with a
// <canvas> context passed to canvasMeasurer().
export function canMeasureOffscreen(scope: unknown): scope is OffscreenCanvasScope {
	return typeof scope === "object" && scope !== null && "OffscreenCanvas" in scope && typeof scope.OffscreenCanvas === "function"
}

// A measurer on a 1×1 OffscreenCanvas of `scope` (a worker's `self` or a window).
export function offscreenMeasurer(scope: unknown): TextMeasurer {
	if (!canMeasureOffscreen(scope)) throw new Error("OffscreenCanvas is not available: measure with a <canvas> context")

	const context = new scope.OffscreenCanvas(1, 1).getContext("2d")

	if (!isContext(context)) throw new Error("OffscreenCanvas 2D contexts are not available: measure with a <canvas> context")

	return canvasMeasurer(context)
}

function isContext(value: unknown): value is CanvasContextLike {
	return typeof value === "object" && value !== null && "measureText" in value && typeof value.measureText === "function"
}
