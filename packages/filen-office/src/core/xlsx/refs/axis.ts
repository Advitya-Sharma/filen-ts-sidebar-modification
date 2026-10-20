import type { AxisEdit } from "../../sheet/model/edits"
import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetRange } from "../../sheet/protocol"
import { shiftSpan } from "../../sheet/refs/formulaRefs"

// Rows or columns inserted or deleted on one sheet, as positions see them. Excel's rules (support.microsoft.com,
// "Insert or delete rows and columns"; ECMA-376 Part 1 §18.17.2.4 for references): an insert at an index moves
// everything at and after it; a range that the insert falls strictly inside grows, one starting at the index
// moves whole. A delete removes what it covers; a range it cuts into shrinks; one it covers whole is gone.

export type Axis = "rows" | "cols"

export interface AxisShift extends AxisEdit {
	readonly axis: Axis
}

export function axisLimit(axis: Axis): number {
	return axis === "rows" ? SHEET_MAX_ROWS : SHEET_MAX_COLS
}

// The new index of `index` on the shifted axis, or -1 when it was deleted or pushed off the sheet.
export function shiftIndex(index: number, shift: AxisShift): number {
	if (index < shift.at) return index

	if (shift.type === "insert") {
		const moved = index + shift.count

		return moved < axisLimit(shift.axis) ? moved : -1
	}

	return index < shift.at + shift.count ? -1 : index - shift.count
}

// The index before the shift of a cell now at `index`, or -1 for an inserted one.
export function preimage(index: number, shift: AxisShift): number {
	if (index < shift.at) return index
	if (shift.type === "delete") return index + shift.count

	return index < shift.at + shift.count ? -1 : index - shift.count
}

// An inclusive span on the shifted axis: null when the delete took all of it (or the insert pushed it off). A span
// reaching the sheet's last row (column) still does after either edit, as the rows a delete brings in at the end
// fall inside it.
export function shiftAxisSpan(start: number, end: number, shift: AxisShift): [number, number] | null {
	const limit = axisLimit(shift.axis)
	const span = shiftSpan(start, end, shift, limit, true)

	if (span !== null && end === limit - 1) span[1] = limit - 1

	return span
}

export function shiftRange(range: SheetRange, shift: AxisShift): SheetRange | null {
	const rows = shift.axis === "rows"
	const span = shiftAxisSpan(rows ? range.startRow : range.startCol, rows ? range.endRow : range.endCol, shift)

	if (span === null) return null
	if (rows) return span[0] === range.startRow && span[1] === range.endRow ? range : { ...range, startRow: span[0], endRow: span[1] }

	return span[0] === range.startCol && span[1] === range.endCol ? range : { ...range, startCol: span[0], endCol: span[1] }
}

// A cell (a range of one) moved: null when it was deleted.
export function shiftCell(row: number, col: number, shift: AxisShift): { row: number; col: number } | null {
	if (shift.axis === "rows") {
		const next = shiftIndex(row, shift)

		return next < 0 ? null : { row: next, col }
	}

	const next = shiftIndex(col, shift)

	return next < 0 ? null : { row, col: next }
}

// Whether a delete cuts a range without taking all of it, or an insert falls strictly inside it: what an array
// formula or a pivot table cannot take.
export function splits(start: number, end: number, shift: AxisShift): boolean {
	if (shift.type === "insert") return shift.at > start && shift.at <= end

	const last = shift.at + shift.count - 1

	return shift.at <= end && last >= start && (shift.at > start || last < end)
}

// Whether the edit moves or removes anything of a span ending at `end`.
export function reaches(end: number, shift: AxisShift): boolean {
	return shift.at <= end
}

export function sameRange(a: SheetRange, b: SheetRange): boolean {
	return a.startRow === b.startRow && a.startCol === b.startCol && a.endRow === b.endRow && a.endCol === b.endCol
}
