import type { SizeAxis } from "../sizes/sizes"

// A run of rows or columns inserted or deleted.
export interface AxisEdit {
	type: "insert" | "delete"
	axis: SizeAxis
	at: number
	count: number
}

// The largest sheet opened or grown by an edit. A workbook sheet holds its values as a dense rectangle,
// so this bounds its rows times its columns, whatever is filled.
export const MAX_SHEET_CELLS = 5_000_000
