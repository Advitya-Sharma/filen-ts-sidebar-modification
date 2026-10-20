import type { XlsxColRange, XlsxSheetView } from "../model/types"
import { axisLimit, shiftAxisSpan, shiftIndex, shiftRange, type AxisShift } from "./axis"
import { shiftRanges } from "./sqref"

// What the model holds for the sheet's own layout, moved: merged ranges, column ranges and the first sheet view.

// Merges as quads; a merge left with one cell is no merge (ECMA-376 Part 1 §18.3.1.55). The same array when
// nothing moved.
export function shiftMerges(merges: Uint32Array, shift: AxisShift): Uint32Array {
	const out: number[] = []
	let changed = false

	for (let k = 0; k + 3 < merges.length; k += 4) {
		const range = { startRow: merges[k] ?? 0, startCol: merges[k + 1] ?? 0, endRow: merges[k + 2] ?? 0, endCol: merges[k + 3] ?? 0 }
		const next = shiftRange(range, shift)

		if (next !== range) changed = true
		if (next === null || (next.startRow === next.endRow && next.startCol === next.endCol)) {
			changed = true

			continue
		}

		out.push(next.startRow, next.startCol, next.endRow, next.endCol)
	}

	return changed ? Uint32Array.from(out) : merges
}

// Column ranges (widths, styles): an insert inside a range widens it, as the inserted columns take the format of
// the columns around them. The same array when nothing moved.
export function shiftCols(cols: readonly XlsxColRange[], shift: AxisShift): readonly XlsxColRange[] {
	if (shift.axis !== "cols") return cols

	let changed = false
	const out: XlsxColRange[] = []

	for (const col of cols) {
		const span = shiftAxisSpan(col.first, col.last, shift)

		if (span === null) {
			changed = true

			continue
		}

		if (span[0] !== col.first || span[1] !== col.last) {
			changed = true
			out.push({ ...col, first: span[0], last: span[1] })
		} else {
			out.push(col)
		}
	}

	return changed ? out : cols
}

// A frozen split (the count of frozen rows or columns) after the edit.
export function shiftSplit(split: number, shift: AxisShift): number {
	if (shift.type === "insert") return shift.at < split ? Math.min(split + shift.count, axisLimit(shift.axis) - 1) : split

	return split - Math.max(0, Math.min(split, shift.at + shift.count) - Math.min(split, shift.at))
}

// A view index (active cell, top-left cell) after the edit: a deleted one goes to the first index after the run.
function viewIndex(index: number, shift: AxisShift): number {
	const next = shiftIndex(index, shift)

	if (next >= 0) return next

	return Math.min(shift.at, axisLimit(shift.axis) - 1)
}

export function shiftView(view: XlsxSheetView, shift: AxisShift): XlsxSheetView {
	const rows = shift.axis === "rows"
	const activeRow = rows ? viewIndex(view.activeRow, shift) : view.activeRow
	const activeCol = rows ? view.activeCol : viewIndex(view.activeCol, shift)
	let selection = shiftRanges(view.selection, shift)

	if (selection.length === 0 && view.selection.length > 0)
		selection = [{ startRow: activeRow, startCol: activeCol, endRow: activeRow, endCol: activeCol }]

	let frozen = view.frozen

	if (frozen !== null) {
		const next = {
			rows: rows ? shiftSplit(frozen.rows, shift) : frozen.rows,
			cols: rows ? frozen.cols : shiftSplit(frozen.cols, shift),
			topLeftRow: rows ? viewIndex(frozen.topLeftRow, shift) : frozen.topLeftRow,
			topLeftCol: rows ? frozen.topLeftCol : viewIndex(frozen.topLeftCol, shift)
		}

		// The bottom-right pane starts at or after the split.
		next.topLeftRow = Math.max(next.topLeftRow, next.rows)
		next.topLeftCol = Math.max(next.topLeftCol, next.cols)

		frozen =
			next.rows === 0 && next.cols === 0
				? null
				: next.rows === frozen.rows &&
					  next.cols === frozen.cols &&
					  next.topLeftRow === frozen.topLeftRow &&
					  next.topLeftCol === frozen.topLeftCol
					? frozen
					: next
	}

	if (activeRow === view.activeRow && activeCol === view.activeCol && selection === view.selection && frozen === view.frozen) return view

	return { ...view, activeRow, activeCol, selection, frozen }
}
