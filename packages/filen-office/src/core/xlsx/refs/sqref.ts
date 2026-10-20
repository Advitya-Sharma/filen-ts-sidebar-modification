import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetRange } from "../../sheet/protocol"
import { cellName, columnIndex } from "../../sheet/refs/cellRef"
import { parseRange } from "../read/sheetData"
import { shiftRange, type AxisShift } from "./axis"

// ST_Sqref (ECMA-376 Part 1 §18.18.76): references separated by spaces. A1 cells and areas, with whole columns
// (A:C) and rows (1:3) read as the areas they name.

const WHOLE_COLUMNS = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/
const WHOLE_ROWS = /^\$?([0-9]{1,7}):\$?([0-9]{1,7})$/

export function parseRef(text: string): SheetRange | null {
	const range = parseRange(text)

	if (range !== null) return range

	const columns = WHOLE_COLUMNS.exec(text)

	if (columns !== null) {
		const a = columnIndex(columns[1] ?? "")
		const b = columnIndex(columns[2] ?? "")

		if (Math.max(a, b) >= SHEET_MAX_COLS) return null

		return { startRow: 0, startCol: Math.min(a, b), endRow: SHEET_MAX_ROWS - 1, endCol: Math.max(a, b) }
	}

	const rows = WHOLE_ROWS.exec(text)

	if (rows !== null) {
		const a = Number(rows[1]) - 1
		const b = Number(rows[2]) - 1

		if (Math.min(a, b) < 0 || Math.max(a, b) >= SHEET_MAX_ROWS) return null

		return { startRow: Math.min(a, b), startCol: 0, endRow: Math.max(a, b), endCol: SHEET_MAX_COLS - 1 }
	}

	return null
}

// null when any part does not read.
export function parseSqref(text: string): SheetRange[] | null {
	const out: SheetRange[] = []

	for (const part of text.trim().split(/\s+/)) {
		if (part === "") continue

		const range = parseRef(part)

		if (range === null) return null

		out.push(range)
	}

	return out
}

export function formatRef(range: SheetRange): string {
	const start = cellName(range.startRow, range.startCol)

	return range.startRow === range.endRow && range.startCol === range.endCol ? start : `${start}:${cellName(range.endRow, range.endCol)}`
}

export function formatSqref(ranges: readonly SheetRange[]): string {
	return ranges.map(formatRef).join(" ")
}

// The ranges moved; deleted ones dropped. The same array when nothing moved.
export function shiftRanges(ranges: readonly SheetRange[], shift: AxisShift): readonly SheetRange[] {
	let changed = false
	const out: SheetRange[] = []

	for (const range of ranges) {
		const next = shiftRange(range, shift)

		if (next !== range) changed = true
		if (next !== null) out.push(next)
	}

	return changed ? out : ranges
}
