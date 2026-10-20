// The PDF viewer's controls as serialisable commands and state, like the DOCX viewer's: an app renders its own
// toolbar from the state and sends commands back (directly on web, over the bridge from a native toolbar on
// mobile). Plus the page geometry the viewer lays pages out with, framework-free.

// Pages render at this scale at 100 % (pdf.js scale 1 is 72 dpi, small next to a 96 dpi screen's text).
export const PDF_BASE_SCALE = 1.5
export const PDF_MIN_SCALE = 0.5
export const PDF_MAX_SCALE = 4
export const PDF_ZOOM_STEP = 0.25
export const PDF_ZOOM_WHEEL_SENSITIVITY = 0.0015

// A page renders once it is this close to the view, and its canvas is released once it is further than the
// evict margin. The gap is hysteresis: one shared threshold would render and evict a page sitting at it again
// and again.
export const PDF_PAGE_RENDER_MARGIN_PX = 400
export const PDF_PAGE_EVICT_MARGIN_PX = 1200

export type PdfViewerCommand =
	| { readonly type: "page.previous" }
	| { readonly type: "page.next" }
	// 1-based.
	| { readonly type: "page.goto"; readonly page: number }
	| { readonly type: "zoom.in" }
	| { readonly type: "zoom.out" }
	// pdf.js scale (PDF_BASE_SCALE is 100 %).
	| { readonly type: "zoom.set"; readonly scale: number }
	// Adds `delta` to the scale on screen (wheel and pinch gestures, which arrive faster than renders).
	| { readonly type: "zoom.by"; readonly delta: number }
	| { readonly type: "find.open" }
	| { readonly type: "find.close" }
	| { readonly type: "find.query"; readonly query: string; readonly caseSensitive?: boolean }
	| { readonly type: "find.next" }
	| { readonly type: "find.previous" }

export type PdfViewerCommandType = PdfViewerCommand["type"]

export interface PdfViewerState {
	// 1-based page most in view.
	readonly page: number
	readonly pageCount: number
	readonly scale: number
	// Percent of PDF_BASE_SCALE.
	readonly zoom: number
	readonly find: {
		readonly open: boolean
		readonly query: string
		readonly caseSensitive: boolean
		// False while the query is being searched.
		readonly done: boolean
		readonly count: number
		// There are more matches than `count` (PDF_FIND_LIMIT).
		readonly capped: boolean
		// Index of the current match, or -1.
		readonly active: number
	}
}

export function isPdfCommandEnabled(state: PdfViewerState, type: PdfViewerCommandType): boolean {
	switch (type) {
		case "page.previous":
			return state.page > 1
		case "page.next":
			return state.page < state.pageCount
		case "zoom.in":
			return state.scale < PDF_MAX_SCALE
		case "zoom.out":
			return state.scale > PDF_MIN_SCALE
		case "find.next":
		case "find.previous":
			return state.find.open && state.find.count > 0
		default:
			return true
	}
}

export function clampPdfScale(scale: number): number {
	return Math.min(PDF_MAX_SCALE, Math.max(PDF_MIN_SCALE, scale))
}

// The +/- steps land on the PDF_ZOOM_STEP grid wherever a wheel zoom left the scale.
export function pdfStepZoomScale(current: number, direction: 1 | -1): number {
	const steps = (current - PDF_BASE_SCALE) / PDF_ZOOM_STEP
	const onGrid = Math.abs(steps - Math.round(steps)) < 1e-9
	const next = onGrid ? Math.round(steps) + direction : direction > 0 ? Math.ceil(steps) : Math.floor(steps)

	return clampPdfScale(PDF_BASE_SCALE + next * PDF_ZOOM_STEP)
}

export function pdfWheelZoomScale(current: number, deltaY: number): number {
	return clampPdfScale(current - deltaY * PDF_ZOOM_WHEEL_SENSITIVITY)
}

export function pdfZoomPercent(scale: number): number {
	return Math.round((scale / PDF_BASE_SCALE) * 100)
}

export type PdfPageAction = "render" | "evict" | "idle"

// What a page's canvas should do. Outside the evict margin it is released whether or not it finished: a render
// cancelled midway had already sized its canvas.
export function pdfPageAction(withinExtendedView: boolean, rendered: boolean): PdfPageAction {
	if (!withinExtendedView) return "evict"

	return rendered ? "idle" : "render"
}

// The page with the highest visible ratio, or null when none is visible (keep the current one). A tie keeps the
// first in iteration order.
export function mostVisiblePage(ratios: ReadonlyMap<number, number>): number | null {
	let bestPage: number | null = null
	let bestRatio = 0

	for (const [page, ratio] of ratios) {
		if (ratio <= 0) continue

		if (bestPage === null || ratio > bestRatio) {
			bestPage = page
			bestRatio = ratio
		}
	}

	return bestPage
}

export interface PdfCanvasDims {
	readonly bufferWidth: number
	readonly bufferHeight: number
	readonly cssWidth: number
	readonly cssHeight: number
}

// The canvas's backing store scales with the device pixel ratio for a sharp bitmap; its CSS box stays the
// viewport's size. Whole pixels: a fractional buffer leaves a transparent 1px gutter in some browsers.
export function pdfCanvasDims(viewportWidth: number, viewportHeight: number, devicePixelRatio: number): PdfCanvasDims {
	const ratio = devicePixelRatio > 0 ? devicePixelRatio : 1

	return {
		bufferWidth: Math.floor(viewportWidth * ratio),
		bufferHeight: Math.floor(viewportHeight * ratio),
		cssWidth: Math.floor(viewportWidth),
		cssHeight: Math.floor(viewportHeight)
	}
}

// The ratio to render at: the device pixel ratio, lowered so the canvas stays within `maxBytes` (4 bytes a
// pixel). A large page zoomed in on a high-density screen would otherwise allocate hundreds of megabytes.
export function pdfRenderRatio(viewportWidth: number, viewportHeight: number, devicePixelRatio: number, maxBytes: number): number {
	const ratio = devicePixelRatio > 0 ? devicePixelRatio : 1
	const bytes = viewportWidth * viewportHeight * ratio * ratio * 4

	if (!(bytes > maxBytes) || viewportWidth <= 0 || viewportHeight <= 0) return ratio

	return Math.sqrt(maxBytes / (viewportWidth * viewportHeight * 4))
}

// The render transform for pdfCanvasDims' buffer scale.
export function pdfRenderTransform(devicePixelRatio: number): [number, number, number, number, number, number] {
	const ratio = devicePixelRatio > 0 ? devicePixelRatio : 1

	return [ratio, 0, 0, ratio, 0, 0]
}

export interface PdfBox {
	readonly left: number
	readonly top: number
	readonly width: number
	readonly height: number
}

// convertToViewportPoint returns the corners in no set order (the y axis flips).
export function pdfBox(x1: number, y1: number, x2: number, y2: number): PdfBox {
	return { left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) }
}
