import { columnName } from "../../../core/sheet/refs/cellRef"
import type { SheetRange } from "../../../core/sheet/protocol"
import type { GridSelection } from "../selection"
import type { SheetGridTheme } from "../theme"
import { screenOf, type AxisView, type GridView } from "../view"
import type { GridCanvas } from "./canvas"
import { segments } from "./cells"

// The chrome layer over the cells: column and row headers (the selection's tinted), the frozen panes' dividers,
// the selection's tint and outline, and a resize's guide. Cheap to draw again on every selection move, so the
// cells layer is left alone then.

export interface ChromeScene {
	readonly view: GridView
	readonly theme: SheetGridTheme
	readonly pixelRatio: number
	readonly selection: GridSelection
	readonly merges: readonly SheetRange[]
	// A resize being dragged: the edge's position on screen.
	readonly guide: { readonly axis: "rows" | "cols"; readonly at: number } | null
}

interface Rect {
	readonly x: number
	readonly y: number
	readonly width: number
	readonly height: number
}

// A range's rectangle within one axis's segment: [start, end) on screen, or null when it is not in it.
function spanIn(view: AxisView, first: number, last: number, segmentFirst: number, segmentLast: number): [number, number] | null {
	const from = Math.max(first, segmentFirst)
	const to = Math.min(last, segmentLast)

	if (from > to) return null

	return [screenOf(view, from), screenOf(view, to) + view.axis.size(to)]
}

// A range's on-screen rectangles, one per pane it shows in.
export function rangeRects(view: GridView, range: SheetRange): Rect[] {
	const rects: Rect[] = []

	for (const rows of segments(view.rows)) {
		const ys = spanIn(view.rows, range.startRow, range.endRow, rows.first, rows.last)

		if (ys === null) continue

		for (const cols of segments(view.cols)) {
			const xs = spanIn(view.cols, range.startCol, range.endCol, cols.first, cols.last)

			if (xs === null) continue

			const x = Math.max(xs[0], cols.start)
			const y = Math.max(ys[0], rows.start)

			rects.push({ x, y, width: Math.min(xs[1], cols.end) - x, height: Math.min(ys[1], rows.end) - y })
		}
	}

	return rects
}

function covers(ranges: readonly SheetRange[], axis: "rows" | "cols", index: number): boolean {
	return ranges.some(range =>
		axis === "rows" ? index >= range.startRow && index <= range.endRow : index >= range.startCol && index <= range.endCol
	)
}

function headers(canvas: GridCanvas, scene: ChromeScene, axis: "rows" | "cols"): void {
	const { view, theme, selection } = scene
	const own = axis === "cols" ? view.cols : view.rows
	const other = axis === "cols" ? view.rows : view.cols

	if (other.header === 0) return

	const band = other.header

	canvas.fillStyle = theme.headerBackground

	if (axis === "cols") canvas.fillRect(own.header, 0, view.width - own.header, band)
	else canvas.fillRect(0, own.header, band, view.height - own.header)

	for (const part of segments(own)) {
		canvas.save()
		canvas.beginPath()

		if (axis === "cols") canvas.rect(part.start, 0, part.end - part.start, band)
		else canvas.rect(0, part.start, band, part.end - part.start)

		canvas.clip()

		for (let index = part.first; index <= part.last; index++) {
			const size = own.axis.size(index)

			if (size === 0) continue

			const start = screenOf(own, index)
			const active = covers(selection.ranges, axis, index)

			if (active) {
				canvas.fillStyle = theme.headerActiveBackground

				if (axis === "cols") canvas.fillRect(start, 0, size, band)
				else canvas.fillRect(0, start, band, size)
			}

			canvas.fillStyle = theme.headerLine

			if (axis === "cols") canvas.fillRect(start + size - 1, 0, 1, band)
			else canvas.fillRect(0, start + size - 1, band, 1)

			if (size >= 6) {
				canvas.fillStyle = active ? theme.headerActiveText : theme.headerText

				if (axis === "cols") canvas.fillText(columnName(index), start + size / 2, band / 2)
				else canvas.fillText(String(index + 1), band / 2, start + size / 2)
			}
		}

		canvas.restore()
	}

	canvas.fillStyle = theme.headerLine

	if (axis === "cols") canvas.fillRect(0, band - 1, view.width, 1)
	else canvas.fillRect(band - 1, 0, 1, view.height)
}

export function paintChrome(canvas: GridCanvas, scene: ChromeScene): void {
	const { view, theme, selection, pixelRatio } = scene

	canvas.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0)
	canvas.clearRect(0, 0, view.width, view.height)

	// The selection: a tint over every range (the active cell left clear), then outlines: 2 px round the last range,
	// 1 px round the others and round the active cell when more than it is selected. Each has a 1 px halo inside it,
	// drawn first, so a line never loses to another outline's halo.
	canvas.fillStyle = theme.selectionFill

	for (const range of selection.ranges)
		for (const rect of rangeRects(view, range)) canvas.fillRect(rect.x, rect.y, rect.width, rect.height)

	const activeRange = scene.merges.find(
		merge =>
			selection.active.row >= merge.startRow &&
			selection.active.row <= merge.endRow &&
			selection.active.col >= merge.startCol &&
			selection.active.col <= merge.endCol
	) ?? { startRow: selection.active.row, startCol: selection.active.col, endRow: selection.active.row, endCol: selection.active.col }
	const activeRects = rangeRects(view, activeRange)

	for (const rect of activeRects) canvas.clearRect(rect.x, rect.y, rect.width, rect.height)

	const last = selection.ranges.length - 1
	const outlines: { rect: Rect; width: number }[] = []

	selection.ranges.forEach((range, k) => {
		for (const rect of rangeRects(view, range)) outlines.push({ rect, width: k === last ? 2 : 1 })
	})

	if (last > 0 || !sameRange(selection.ranges[0], activeRange)) for (const rect of activeRects) outlines.push({ rect, width: 1 })

	canvas.fillStyle = theme.selectionHalo

	for (const { rect, width } of outlines) ring(canvas, rect, width, 1)

	canvas.fillStyle = theme.selectionLine

	for (const { rect, width } of outlines) ring(canvas, rect, 0, width)

	// Frozen panes' dividers.
	canvas.fillStyle = theme.frozenLine

	if (view.rows.frozen > 0) canvas.fillRect(0, view.rows.header + view.rows.frozenSize - 1, view.width, 1)
	if (view.cols.frozen > 0) canvas.fillRect(view.cols.header + view.cols.frozenSize - 1, 0, 1, view.height)

	canvas.font = theme.headerFont
	canvas.textAlign = "center"
	canvas.textBaseline = "middle"
	headers(canvas, scene, "cols")
	headers(canvas, scene, "rows")

	// The corner.
	if (view.cols.header > 0 && view.rows.header > 0) {
		canvas.fillStyle = theme.headerBackground
		canvas.fillRect(0, 0, view.cols.header, view.rows.header)
		canvas.fillStyle = theme.headerLine
		canvas.fillRect(view.cols.header - 1, 0, 1, view.rows.header)
		canvas.fillRect(0, view.rows.header - 1, view.cols.header, 1)
	}

	if (scene.guide !== null) {
		canvas.fillStyle = theme.resizeLine

		if (scene.guide.axis === "cols") canvas.fillRect(scene.guide.at - 1, 0, 2, view.height)
		else canvas.fillRect(0, scene.guide.at - 1, view.width, 2)
	}
}

// A ring `width` wide, `inset` inside the rectangle's edges (its left and top gridlines included), so it never
// covers the neighbouring cells' text.
function ring(canvas: GridCanvas, rect: Rect, inset: number, width: number): void {
	const x = rect.x - 1 + inset
	const y = rect.y - 1 + inset
	const w = rect.width + 1 - 2 * inset
	const h = rect.height + 1 - 2 * inset

	if (w <= 0 || h <= 0) return

	canvas.fillRect(x, y, w, width)
	canvas.fillRect(x, y + h - width, w, width)
	canvas.fillRect(x, y, width, h)
	canvas.fillRect(x + w - width, y, width, h)
}

function sameRange(a: SheetRange | undefined, b: SheetRange): boolean {
	return a !== undefined && a.startRow === b.startRow && a.startCol === b.startCol && a.endRow === b.endRow && a.endCol === b.endCol
}
