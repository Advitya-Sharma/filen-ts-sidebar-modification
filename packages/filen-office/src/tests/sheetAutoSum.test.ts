import { describe, expect, it } from "vitest"
import { autoSum, type AutoSumCells } from "../core/sheet/functions/autoSum"

// A sheet from rows of cells: numbers, text, or null for empty.
function sheet(rows: (number | string | null)[][]): AutoSumCells {
	const at = (row: number, col: number): number | string | null => rows[row]?.[col] ?? null

	return { number: (row, col) => typeof at(row, col) === "number", filled: (row, col) => at(row, col) !== null }
}

const cell = (row: number, col: number) => ({ startRow: row, startCol: col, endRow: row, endCol: col })

describe("AutoSum", () => {
	it("sums the numbers above a cell, else those to its left, as an entry to finish", () => {
		const cells = sheet([
			["Item", "Jan"],
			["a", 1],
			["b", 2],
			[null, null],
			[3, 4, null]
		])

		expect(autoSum(cells, cell(3, 1), { row: 3, col: 1 })).toEqual({ kind: "edit", entry: "=SUM(B2:B3)" })
		expect(autoSum(cells, cell(4, 2), { row: 4, col: 2 })).toEqual({ kind: "edit", entry: "=SUM(A5:B5)" })
		expect(autoSum(cells, cell(0, 3), { row: 0, col: 3 })).toEqual({ kind: "edit", entry: "=SUM()" })
	})

	it("passes over blanks between the cell and the numbers above, and stops at text", () => {
		expect(autoSum(sheet([["x"], [1], [2], [null], [null]]), cell(4, 0), { row: 4, col: 0 })).toEqual({
			kind: "edit",
			entry: "=SUM(A2:A3)"
		})
	})

	it("puts a block's column sums in its empty last row, or the row below it", () => {
		const cells = sheet([
			[1, 2],
			[3, 4],
			[null, null]
		])

		expect(autoSum(cells, { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, { row: 0, col: 0 })).toEqual({
			kind: "fill",
			cells: [
				{ row: 2, col: 0, entry: "=SUM(A1:A2)" },
				{ row: 2, col: 1, entry: "=SUM(B1:B2)" }
			]
		})
		expect(autoSum(cells, { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, { row: 0, col: 0 })).toEqual({
			kind: "fill",
			cells: [
				{ row: 2, col: 0, entry: "=SUM(A1:A2)" },
				{ row: 2, col: 1, entry: "=SUM(B1:B2)" }
			]
		})
	})

	it("puts a block's row sums in its empty last column, and a row's sum after it", () => {
		const cells = sheet([
			[1, 2, null],
			[3, 4, null]
		])

		expect(autoSum(cells, { startRow: 0, startCol: 0, endRow: 1, endCol: 2 }, { row: 0, col: 0 })).toEqual({
			kind: "fill",
			cells: [
				{ row: 0, col: 2, entry: "=SUM(A1:B1)" },
				{ row: 1, col: 2, entry: "=SUM(A2:B2)" }
			]
		})
		expect(autoSum(cells, { startRow: 0, startCol: 0, endRow: 0, endCol: 1 }, { row: 0, col: 0 })).toEqual({
			kind: "fill",
			cells: [{ row: 0, col: 2, entry: "=SUM(A1:B1)" }]
		})
	})

	// LibreOffice 7.5 (UNO .uno:AutoSum over A1:C3): C1 =SUM(A1:B1), C2 =SUM(A2:B2), A3 =SUM(A1:A2), B3 =SUM(B1:B2),
	// C3 =SUM(A3:B3).
	it("fills row sums, column sums and a grand total when the block's last row and last column are both empty", () => {
		const cells = sheet([
			[1, 2, null],
			[3, 4, null],
			[null, null, null]
		])

		expect(autoSum(cells, { startRow: 0, startCol: 0, endRow: 2, endCol: 2 }, { row: 0, col: 0 })).toEqual({
			kind: "fill",
			cells: [
				{ row: 0, col: 2, entry: "=SUM(A1:B1)" },
				{ row: 1, col: 2, entry: "=SUM(A2:B2)" },
				{ row: 2, col: 0, entry: "=SUM(A1:A2)" },
				{ row: 2, col: 1, entry: "=SUM(B1:B2)" },
				{ row: 2, col: 2, entry: "=SUM(A3:B3)" }
			]
		})
	})
})
