// Scaled scrolling: browsers cap an element's size (Firefox near 17.9 million pixels, Chrome and Safari near 33.5
// million), and a sheet's 1,048,576 rows at the default height already pass 20 million. The scrolled content is
// therefore at most MAX_SCROLL_PX. A scroll position maps one to one onto the virtual one for SCROLL_MARGIN_PX at
// each end, and linearly in between: both ends meet, every row stays reachable, one pixel of scrollbar moves at most
// a few pixels of sheet, and a finger's pan near either end (which moves the sheet one to one) reaches it exactly.

export const MAX_SCROLL_PX = 8_000_000
export const SCROLL_MARGIN_PX = 1_000_000

// The content size the browser scrolls for a virtual extent.
export function domExtent(virtual: number): number {
	return Math.min(Math.max(0, virtual), MAX_SCROLL_PX)
}

// The one-to-one stretch at each end: a quarter of the browser's range at most, so the middle keeps half of it.
function margin(domRange: number): number {
	return Math.min(SCROLL_MARGIN_PX, domRange / 4)
}

// The virtual scroll offset of a browser scroll offset: `view` is the scrolled pane's own size.
export function toVirtual(dom: number, virtual: number, view: number): number {
	const domRange = domExtent(virtual) - view
	const virtualRange = virtual - view

	if (domRange <= 0 || virtualRange <= 0) return 0

	const at = Math.max(0, Math.min(dom, domRange))

	if (domRange === virtualRange) return at

	const edge = margin(domRange)

	if (at <= edge) return at
	if (at >= domRange - edge) return virtualRange - (domRange - at)

	return Math.round(edge + ((at - edge) / (domRange - 2 * edge)) * (virtualRange - 2 * edge))
}

export function toDom(offset: number, virtual: number, view: number): number {
	const domRange = domExtent(virtual) - view
	const virtualRange = virtual - view

	if (domRange <= 0 || virtualRange <= 0) return 0

	const at = Math.max(0, Math.min(offset, virtualRange))

	if (domRange === virtualRange) return at

	const edge = margin(domRange)

	if (at <= edge) return at
	if (at >= virtualRange - edge) return domRange - (virtualRange - at)

	return edge + ((at - edge) / (virtualRange - 2 * edge)) * (domRange - 2 * edge)
}
