import { SHEET_MAX_COLS, SHEET_MAX_ROWS, columnIndex, type SheetRange } from "@filen/office/core/sheet"

const CELL = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/

function cell(text: string): { row: number; col: number } | null {
	const match = CELL.exec(text)

	if (match === null) {
		return null
	}

	const col = columnIndex(match[1] ?? "")
	const row = Number(match[2]) - 1

	return row >= 0 && row < SHEET_MAX_ROWS && col >= 0 && col < SHEET_MAX_COLS ? { row, col } : null
}

// What the name box was given: a cell ("B4", "$B$4") or a range ("A1:C3", either corner first), in this sheet.
// null for anything else.
export function parseNameBox(input: string): SheetRange | null {
	const parts = input.trim().split(":")

	if (parts.length > 2) {
		return null
	}

	const start = cell(parts[0] ?? "")
	const end = parts.length === 2 ? cell(parts[1] ?? "") : start

	if (start === null || end === null) {
		return null
	}

	return {
		startRow: Math.min(start.row, end.row),
		startCol: Math.min(start.col, end.col),
		endRow: Math.max(start.row, end.row),
		endCol: Math.max(start.col, end.col)
	}
}
