import { rangeName } from "../refs/cellRef"
import type { SheetRange } from "../protocol"

// AutoSum (Alt+=, Excel's "Use AutoSum to sum numbers" on support.microsoft.com): a SUM of the numbers next to the
// selection. One cell: the numbers above it, else those to its left, as the entry to finish (Excel opens the formula
// for the user to check); none, an empty SUM() to fill in. A block: its columns' sums in its empty last row (or, with
// none, in the row below it), or its rows' sums in its empty last column; with both empty, both and a grand total.

declare module "../protocol" {
	interface SheetQueryMap {
		autoSum: {
			request: { readonly sheet: SheetId; readonly range: SheetRange; readonly active: SheetCell }
			result: SheetAutoSum
		}
	}
}

export type SheetAutoSum =
	// The entry to open in the active cell (with the caret before its last parenthesis).
	| { readonly kind: "edit"; readonly entry: string }
	// Sums written at once.
	| { readonly kind: "fill"; readonly cells: readonly { readonly row: number; readonly col: number; readonly entry: string }[] }

// One sheet's cells as AutoSum reads them.
export interface AutoSumCells {
	// The cell holds a number (a value, or a formula's numeric result).
	number(row: number, col: number): boolean
	// The cell holds anything at all.
	filled(row: number, col: number): boolean
}

// The run of numbers that ends next to (row, col) going up or left, as a range; null when the cell there holds none.
// Leading blanks between the cell and the numbers are passed over, as Excel's AutoSum reaches past them.
function run(cells: AutoSumCells, row: number, col: number, up: boolean): SheetRange | null {
	let at = up ? row - 1 : col - 1

	while (at >= 0 && !(up ? cells.filled(at, col) : cells.filled(row, at))) at--

	if (at < 0 || !(up ? cells.number(at, col) : cells.number(row, at))) return null

	const end = at

	while (at - 1 >= 0 && (up ? cells.number(at - 1, col) : cells.number(row, at - 1))) at--

	return up ? { startRow: at, startCol: col, endRow: end, endCol: col } : { startRow: row, startCol: at, endRow: row, endCol: end }
}

function sum(range: SheetRange): string {
	return `=SUM(${rangeName(range)})`
}

export function autoSum(cells: AutoSumCells, range: SheetRange, active: { readonly row: number; readonly col: number }): SheetAutoSum {
	const single = range.startRow === range.endRow && range.startCol === range.endCol

	if (single) {
		const target = run(cells, active.row, active.col, true) ?? run(cells, active.row, active.col, false)

		return { kind: "edit", entry: target === null ? "=SUM()" : sum(target) }
	}

	const { startRow, startCol, endRow, endCol } = range
	const rowEmpty = (row: number): boolean => {
		for (let col = startCol; col <= endCol; col++) if (cells.filled(row, col)) return false

		return true
	}
	const colEmpty = (col: number): boolean => {
		for (let row = startRow; row <= endRow; row++) if (cells.filled(row, col)) return false

		return true
	}
	const out: { row: number; col: number; entry: string }[] = []

	// A one-row selection sums along the row, into its empty last cell or the cell after it.
	if (startRow === endRow) {
		const last = colEmpty(endCol) ? endCol : endCol + 1
		const from = last === endCol ? endCol - 1 : endCol

		out.push({ row: startRow, col: last, entry: sum({ startRow, startCol, endRow, endCol: from }) })

		return { kind: "fill", cells: out }
	}

	// Both the last row and the last column empty: the rows' sums in the last column, the columns' in the last row, and
	// the corner the sum of the last row (LibreOffice's AutoSum, the tie-break where Microsoft's page is silent).
	if (endRow > startRow && endCol > startCol && rowEmpty(endRow) && colEmpty(endCol)) {
		for (let row = startRow; row < endRow; row++)
			out.push({ row, col: endCol, entry: sum({ startRow: row, startCol, endRow: row, endCol: endCol - 1 }) })

		for (let col = startCol; col < endCol; col++)
			out.push({ row: endRow, col, entry: sum({ startRow, startCol: col, endRow: endRow - 1, endCol: col }) })

		out.push({ row: endRow, col: endCol, entry: sum({ startRow: endRow, startCol, endRow, endCol: endCol - 1 }) })

		return { kind: "fill", cells: out }
	}

	if (endRow > startRow && rowEmpty(endRow) === false && endCol > startCol && colEmpty(endCol)) {
		for (let row = startRow; row <= endRow; row++)
			out.push({ row, col: endCol, entry: sum({ startRow: row, startCol, endRow: row, endCol: endCol - 1 }) })

		return { kind: "fill", cells: out }
	}

	const into = rowEmpty(endRow) ? endRow : endRow + 1
	const last = into === endRow ? endRow - 1 : endRow

	for (let col = startCol; col <= endCol; col++)
		out.push({ row: into, col, entry: sum({ startRow, startCol: col, endRow: last, endCol: col }) })

	return { kind: "fill", cells: out }
}
