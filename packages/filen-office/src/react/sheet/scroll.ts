// Scaled scrolling: browsers cap an element's size (Firefox near 17.9 million pixels, Chrome and Safari near 33.5
// million), and a sheet's 1,048,576 rows at the default height already pass 20 million. The scrolled content is
// therefore at most MAX_SCROLL_PX, and a scroll position maps linearly onto the virtual one: both ends meet, every
// row stays reachable, and one pixel of scrollbar moves at most a few pixels of sheet.

export const MAX_SCROLL_PX = 8_000_000

// The content size the browser scrolls for a virtual extent.
export function domExtent(virtual: number): number {
	return Math.min(Math.max(0, virtual), MAX_SCROLL_PX)
}

// The virtual scroll offset of a browser scroll offset: `view` is the scrolled pane's own size.
export function toVirtual(dom: number, virtual: number, view: number): number {
	const domRange = domExtent(virtual) - view
	const virtualRange = virtual - view

	if (domRange <= 0 || virtualRange <= 0) return 0
	if (domRange === virtualRange) return Math.max(0, Math.min(dom, virtualRange))

	return Math.max(0, Math.min(virtualRange, Math.round((dom / domRange) * virtualRange)))
}

export function toDom(offset: number, virtual: number, view: number): number {
	const domRange = domExtent(virtual) - view
	const virtualRange = virtual - view

	if (domRange <= 0 || virtualRange <= 0) return 0
	if (domRange === virtualRange) return Math.max(0, Math.min(offset, domRange))

	return Math.max(0, Math.min(domRange, (offset / virtualRange) * domRange))
}
