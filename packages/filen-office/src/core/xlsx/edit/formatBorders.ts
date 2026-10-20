import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetId, type SheetRange } from "../../sheet/protocol"
import type { SheetChangeSink, SheetHandled } from "../../sheet/registry"
import type { Slicing } from "../../slicing"
import type { SheetBorderEdges, SheetBorderLineInput } from "../../sheet/style/borders"
import { XfApply, type XlsxBorder, type XlsxBorderEdge, type XlsxXf } from "../model/types"
import type { XlsxDocument } from "./document"
import { appendXf, formatPlan, xlsxColor, type FormatPlanEntry } from "./format"
import { wholeAxis } from "./formatAxes"

// format.borders: each cell of a range takes the lines of the edges it is on (the outline on the range's outer
// cells, the inside lines on both cells that share them), as appended border records and cell formats. A cleared
// outline edge clears the neighbouring cell's facing line too (SheetBorderPatch).

// What one cell's border record gets: per side a line (set), null (cleared) or undefined (kept).
interface CellBorderChange {
	readonly top?: SheetBorderLineInput | null
	readonly bottom?: SheetBorderLineInput | null
	readonly left?: SheetBorderLineInput | null
	readonly right?: SheetBorderLineInput | null
	readonly diagonalUp?: SheetBorderLineInput | null
	readonly diagonalDown?: SheetBorderLineInput | null
}

type Side = "top" | "bottom" | "left" | "right"

// An automatic line's colour as Excel writes it: the system foreground (indexed 64, §18.8.27).
const AUTOMATIC = { indexed: 64 } as const

function edgeOf(line: SheetBorderLineInput): XlsxBorderEdge {
	return { style: line.style, color: line.color === undefined || line.color === null ? AUTOMATIC : xlsxColor(line.color) }
}

function sameEdge(a: XlsxBorderEdge | undefined, b: XlsxBorderEdge | undefined): boolean {
	if (a?.style === undefined || b?.style === undefined) return a?.style === b?.style

	const x = a.color
	const y = b.color

	return (
		a.style === b.style &&
		x?.argb === y?.argb &&
		x?.theme === y?.theme &&
		x?.tint === y?.tint &&
		x?.indexed === y?.indexed &&
		x?.auto === y?.auto
	)
}

function withEdge(border: XlsxBorder, side: Side, line: SheetBorderLineInput | null): XlsxBorder {
	const next: { -readonly [K in keyof XlsxBorder]: XlsxBorder[K] } = { ...border }

	if (line === null) delete next[side]
	else next[side] = edgeOf(line)

	return next
}

// The diagonal record holds one line for both directions (§18.8.4 diagonalUp, diagonalDown): setting one
// direction draws it in that line, clearing the last direction drops it.
function withDiagonal(border: XlsxBorder, direction: "diagonalUp" | "diagonalDown", line: SheetBorderLineInput | null): XlsxBorder {
	const next: { -readonly [K in keyof XlsxBorder]: XlsxBorder[K] } = { ...border }

	if (line === null) {
		delete next[direction]

		if (next.diagonalUp !== true && next.diagonalDown !== true) delete next.diagonal
	} else {
		next[direction] = true
		next.diagonal = edgeOf(line)
	}

	return next
}

export function changedBorder(border: XlsxBorder, change: CellBorderChange): XlsxBorder {
	let next = border

	for (const side of ["top", "bottom", "left", "right"] as const) {
		const line = change[side]

		if (line !== undefined) next = withEdge(next, side, line)
	}

	if (change.diagonalUp !== undefined) next = withDiagonal(next, "diagonalUp", change.diagonalUp)
	if (change.diagonalDown !== undefined) next = withDiagonal(next, "diagonalDown", change.diagonalDown)

	return next
}

function sameBorder(a: XlsxBorder, b: XlsxBorder): boolean {
	return (
		sameEdge(a.top, b.top) &&
		sameEdge(a.bottom, b.bottom) &&
		sameEdge(a.left, b.left) &&
		sameEdge(a.right, b.right) &&
		sameEdge(a.diagonal, b.diagonal) &&
		(a.diagonalUp === true) === (b.diagonalUp === true) &&
		(a.diagonalDown === true) === (b.diagonalDown === true)
	)
}

const NO_BORDER: XlsxBorder = {}

// Cell formats with a border changed, each distinct (format, change) pair once per edit.
class BorderPatcher {
	private readonly done = new Map<string, number>()

	private readonly document: XlsxDocument

	constructor(document: XlsxDocument) {
		this.document = document
	}

	xf(index: number, change: CellBorderChange, key: number): number {
		const memo = `${String(key)}:${String(index)}`
		let next = this.done.get(memo)

		if (next === undefined) {
			next = this.patch(index, change)
			this.done.set(memo, next)
		}

		return next
	}

	private patch(index: number, change: CellBorderChange): number {
		const styles = this.document.workbook.styles
		const xf: XlsxXf = styles.cellXfs[index] ??
			styles.cellXfs[0] ?? { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, flags: 0 }
		const border = styles.borders[xf.borderId] ?? NO_BORDER
		const next = changedBorder(border, change)

		if (sameBorder(border, next)) return index

		return appendXf(this.document, { ...xf, borderId: styles.appendBorder(next), flags: xf.flags | XfApply.BORDER })
	}
}

// The change of a cell at `row`, `col` of `range`, and a small key naming it (its place in the range).
function cellChange(edges: SheetBorderEdges, range: SheetRange, row: number, col: number): { change: CellBorderChange; key: number } {
	const first = row === range.startRow
	const last = row === range.endRow
	const leftmost = col === range.startCol
	const rightmost = col === range.endCol
	const key = (first ? 1 : 0) | (last ? 2 : 0) | (leftmost ? 4 : 0) | (rightmost ? 8 : 0)
	const change: { -readonly [K in keyof CellBorderChange]: CellBorderChange[K] } = {}
	const set = (side: keyof CellBorderChange, line: SheetBorderLineInput | null | undefined): void => {
		if (line !== undefined) change[side] = line
	}

	set("top", first ? edges.top : edges.insideHorizontal)
	set("bottom", last ? edges.bottom : edges.insideHorizontal)
	set("left", leftmost ? edges.left : edges.insideVertical)
	set("right", rightmost ? edges.right : edges.insideVertical)
	set("diagonalUp", edges.diagonalUp)
	set("diagonalDown", edges.diagonalDown)

	return { change, key }
}

// The cells beside a range whose facing line a cleared outline edge clears: above, below, left and right of it.
function neighbours(edges: SheetBorderEdges, range: SheetRange): { range: SheetRange; side: Side; key: number }[] {
	const out: { range: SheetRange; side: Side; key: number }[] = []

	if (edges.top === null && range.startRow > 0)
		out.push({ range: { ...range, startRow: range.startRow - 1, endRow: range.startRow - 1 }, side: "bottom", key: 16 })
	if (edges.bottom === null && range.endRow < SHEET_MAX_ROWS - 1)
		out.push({ range: { ...range, startRow: range.endRow + 1, endRow: range.endRow + 1 }, side: "top", key: 17 })
	if (edges.left === null && range.startCol > 0)
		out.push({ range: { ...range, startCol: range.startCol - 1, endCol: range.startCol - 1 }, side: "right", key: 18 })
	if (edges.right === null && range.endCol < SHEET_MAX_COLS - 1)
		out.push({ range: { ...range, startCol: range.endCol + 1, endCol: range.endCol + 1 }, side: "left", key: 19 })

	return out
}

const hasEdges = (edges: SheetBorderEdges): boolean => Object.values(edges).some(line => line !== undefined)

export function applyBorders(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	edges: SheetBorderEdges,
	changes: SheetChangeSink,
	slicing?: Slicing
): Promise<SheetHandled> {
	if (!hasEdges(edges)) return Promise.resolve(null)

	const patcher = new BorderPatcher(document)
	const entries: FormatPlanEntry[] = []

	for (const range of ranges) {
		entries.push({
			range,
			patch: (xf, row, col) => {
				const { change, key } = cellChange(edges, range, row, col)

				return patcher.xf(xf, change, key)
			}
		})

		// A neighbour of whole columns (rows) is whole too: its axis format loses the facing line as its cells do.
		for (const neighbour of neighbours(edges, range)) {
			const change: CellBorderChange = { [neighbour.side]: null }

			entries.push({
				range: neighbour.range,
				patch: xf => patcher.xf(xf, change, neighbour.key),
				existing: wholeAxis(neighbour.range) === null
			})
		}
	}

	return formatPlan(document, sheet, entries, changes, slicing)
}
