import type { DrawingAnchor, DrawingMarker } from "../../../core/xlsx/drawings/anchor"
import type { GridAxis } from "../axis"
import type { AxisView, GridView } from "../view"

// Drawing objects on the grid: an anchor's box in the sheet's own pixels at the zoom (from A1's top left, through the
// axes the grid draws with), where that box is on screen, and the anchor a moved or resized box gets back. Offsets
// past a cell's size end at its far edge, as Excel draws them; indices past the axes' extent take the default size.

export const EMU_PER_PX = 9525

export interface SheetBox {
	readonly x: number
	readonly y: number
	readonly width: number
	readonly height: number
}

// Where `index` starts on the axis, px (past the axis's extent at its default size).
export function axisStart(axis: GridAxis, index: number): number {
	return index <= axis.count ? axis.offset(index) : axis.total + (index - axis.count) * axis.defaultSize
}

function axisSize(axis: GridAxis, index: number): number {
	return index < axis.count ? axis.size(index) : axis.defaultSize
}

function edge(axis: GridAxis, index: number, offset: number, scale: number): number {
	return axisStart(axis, index) + Math.min((offset / EMU_PER_PX) * scale, axisSize(axis, index))
}

export interface GridAxesLike {
	readonly rows: GridAxis
	readonly cols: GridAxis
	readonly scale: number
}

// The box an anchor puts its object in, px at the zoom from A1's top left.
export function anchorBox(anchor: DrawingAnchor, axes: GridAxesLike): SheetBox {
	const { rows, cols, scale } = axes

	switch (anchor.mode) {
		case "absolute":
			return {
				x: (anchor.x / EMU_PER_PX) * scale,
				y: (anchor.y / EMU_PER_PX) * scale,
				width: (anchor.cx / EMU_PER_PX) * scale,
				height: (anchor.cy / EMU_PER_PX) * scale
			}
		case "oneCell": {
			const x = edge(cols, anchor.from.col, anchor.from.colOff, scale)
			const y = edge(rows, anchor.from.row, anchor.from.rowOff, scale)

			return { x, y, width: (anchor.cx / EMU_PER_PX) * scale, height: (anchor.cy / EMU_PER_PX) * scale }
		}
		case "twoCell": {
			const x = edge(cols, anchor.from.col, anchor.from.colOff, scale)
			const y = edge(rows, anchor.from.row, anchor.from.rowOff, scale)

			return {
				x,
				y,
				width: Math.max(0, edge(cols, anchor.to.col, anchor.to.colOff, scale) - x),
				height: Math.max(0, edge(rows, anchor.to.row, anchor.to.rowOff, scale) - y)
			}
		}
	}
}

// The index whose span holds `at` and the offset into it (EMU at 100 %), past the axis's extent too.
function cellAt(axis: GridAxis, at: number, scale: number, limit: number): { index: number; offset: number } {
	const position = Math.max(0, at)
	let index: number

	if (position >= axis.total) index = axis.count + Math.floor((position - axis.total) / axis.defaultSize)
	else index = axis.indexAt(position)

	index = Math.min(index, limit - 1)

	const offset = Math.min(position - axisStart(axis, index), axisSize(axis, index))

	return { index, offset: Math.max(0, Math.round((offset / scale) * EMU_PER_PX)) }
}

function marker(axes: GridAxesLike, x: number, y: number, limits: { rows: number; cols: number }): DrawingMarker {
	const col = cellAt(axes.cols, x, axes.scale, limits.cols)
	const row = cellAt(axes.rows, y, axes.scale, limits.rows)

	return { row: row.index, col: col.index, rowOff: row.offset, colOff: col.offset }
}

// The anchor of the same kind as `anchor` (and its editAs) that puts the object in `box`.
export function anchorOf(anchor: DrawingAnchor, box: SheetBox, axes: GridAxesLike, limits: { rows: number; cols: number }): DrawingAnchor {
	const emu = (px: number): number => Math.max(0, Math.round((px / axes.scale) * EMU_PER_PX))
	const x = Math.max(0, box.x)
	const y = Math.max(0, box.y)

	switch (anchor.mode) {
		case "absolute":
			return { mode: "absolute", x: emu(x), y: emu(y), cx: emu(box.width), cy: emu(box.height) }
		case "oneCell":
			return { mode: "oneCell", from: marker(axes, x, y, limits), cx: emu(box.width), cy: emu(box.height) }
		case "twoCell":
			return {
				mode: "twoCell",
				editAs: anchor.editAs,
				from: marker(axes, x, y, limits),
				to: marker(axes, x + Math.max(0, box.width), y + Math.max(0, box.height), limits)
			}
	}
}

// Where a sheet position along an axis is on screen: in the frozen part when it lies there, else in the scrolled part.
export function screenAt(view: AxisView, at: number, frozen: boolean): number {
	return frozen
		? view.header + at - view.axis.offset(view.frozenFirst)
		: view.header + view.frozenSize + at - view.bodyStart - view.scroll
}

// Whether a box starting at `index` lies in the frozen part of the axis.
export function inFrozen(view: AxisView, index: number): boolean {
	return view.frozen > 0 && index < view.frozenEnd
}

export interface ScreenPlacement {
	// The box on screen, from the grid's top left (laid out: mirror it for a right-to-left sheet).
	readonly box: SheetBox
	// The part of the screen it may draw in (its pane, below the headers).
	readonly clip: SheetBox
}

// An object's box on screen, and the pane it is drawn in (by the cell its top-left corner is in).
export function placeOnScreen(view: GridView, box: SheetBox): ScreenPlacement {
	const rowsFrozen = inFrozen(view.rows, view.rows.axis.indexAt(box.y)) && box.y < view.rows.axis.offset(view.rows.frozenEnd)
	const colsFrozen = inFrozen(view.cols, view.cols.axis.indexAt(box.x)) && box.x < view.cols.axis.offset(view.cols.frozenEnd)
	const top = rowsFrozen ? view.rows.header : view.rows.header + view.rows.frozenSize
	const left = colsFrozen ? view.cols.header : view.cols.header + view.cols.frozenSize

	return {
		box: { x: screenAt(view.cols, box.x, colsFrozen), y: screenAt(view.rows, box.y, rowsFrozen), width: box.width, height: box.height },
		clip: { x: left, y: top, width: Math.max(0, view.width - left), height: Math.max(0, view.height - top) }
	}
}

// Whether a box placed on screen (placeOnScreen's) meets the view widened by `margin` times its size on each side.
export function nearView(view: { readonly width: number; readonly height: number }, box: SheetBox, margin: number): boolean {
	const dx = view.width * margin
	const dy = view.height * margin

	return box.x < view.width + dx && box.x + box.width > -dx && box.y < view.height + dy && box.y + box.height > -dy
}

// A selection handle: a corner or an edge, by compass point.
export type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w"

const MIN_SIZE = 4

// The box after a drag of (dx, dy) screen px from `start`, by `handle` (null: the whole object).
export function draggedBox(start: SheetBox, handle: Handle | null, dx: number, dy: number, keepRatio: boolean): SheetBox {
	if (handle === null) return { ...start, x: start.x + dx, y: start.y + dy }

	let { x, y, width, height } = start
	const west = handle.includes("w")
	const east = handle.includes("e")
	const north = handle.startsWith("n")
	const south = handle.startsWith("s")

	if (east) width = Math.max(MIN_SIZE, start.width + dx)
	if (west) width = Math.max(MIN_SIZE, start.width - dx)
	if (south) height = Math.max(MIN_SIZE, start.height + dy)
	if (north) height = Math.max(MIN_SIZE, start.height - dy)

	if (keepRatio && (east || west) && (north || south) && start.width > 0 && start.height > 0) {
		const ratio = start.width / start.height

		if (width / height > ratio) height = width / ratio
		else width = height * ratio
	}

	if (west) x = start.x + start.width - width
	if (north) y = start.y + start.height - height

	return { x, y, width, height }
}
