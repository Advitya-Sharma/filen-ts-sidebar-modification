import { describe, expect, it } from "vitest"
import { CellKind, RowFlag } from "../core/xlsx/model/types"
import { BLOCK_ROWS, SheetStore, StoreBuilder } from "../core/xlsx/read/index"
import { FormulaTable } from "../core/xlsx/model/formulas"
import { SheetStrings } from "../core/xlsx/model/strings"
import { rowState } from "../core/xlsx/edit/history"

// Rows and cells set many at a time (a format on whole rows or a long column, a resize of many): the same as one
// setRow or setCell each, with each block rebuilt once.

function storeOf(rows: number, firstRow: number): SheetStore {
	const builder = new StoreBuilder()

	for (let row = firstRow; row < firstRow + rows; row++) {
		builder.beginRow()
		builder.cell(0, CellKind.NUMBER, row, 0, -1)
		builder.endRow(row, 0, 0, 0, Number.NaN, -1)
	}

	return new SheetStore(builder.finish(), new SheetStrings(), new FormulaTable())
}

describe("setting many rows' properties", () => {
	it("gives every row what one setRow per row gives it, across blocks and before the first", () => {
		const one = storeOf(5000, 100)
		const many = storeOf(5000, 100)
		// Rows before the sheet's first, inside blocks, between them and past the last, out of order.
		const rows = [...Array.from({ length: 1500 }, (_, k) => 7 + k * 4), ...Array.from({ length: 50 }, (_, k) => 9000 + k), 3, 40].map(
			(row, k) => ({ row, height: 12 + (k % 5), style: k % 3, flags: RowFlag.CUSTOM_FORMAT | RowFlag.CUSTOM_HEIGHT })
		)

		for (const { row, height, style, flags } of rows) one.setRow(row, height, style, flags)

		many.setRows(rows)

		for (const row of [3, 7, 40, 101, 103, 4003, 5099, 6000, 9000, 9049, 9050]) expect(rowState(many, row)).toEqual(rowState(one, row))

		expect(many.blocks.every(block => block.rowCount <= BLOCK_ROWS)).toBe(true)
		// The cells stay where they were.
		expect(many.value(many.locate(4321, 0))).toBe(4321)
	})
})

describe("setting many cells", () => {
	it("writes what one setCell per cell writes: in place, added, removed, in new rows and blocks", () => {
		const one = storeOf(3000, 100)
		const many = storeOf(3000, 100)
		const writes = [
			// In place, a style only.
			...Array.from({ length: 300 }, (_, k) => ({
				row: 100 + k * 7,
				col: 0,
				kind: CellKind.NUMBER,
				value: 100 + k * 7,
				style: 2,
				formula: -1,
				origin: -1
			})),
			// Added beside existing cells, and in rows the sheet does not have, before, between and after its rows.
			...Array.from({ length: 2000 }, (_, k) => ({
				row: 100 + k,
				col: 3,
				kind: CellKind.EMPTY,
				value: 0,
				style: 1,
				formula: -1,
				origin: -1
			})),
			...[2, 50, 4000, 70_000].map(row => ({ row, col: 5, kind: CellKind.EMPTY, value: 0, style: 1, formula: -1, origin: -1 })),
			// Removed: content gone and the default format.
			...Array.from({ length: 50 }, (_, k) => ({
				row: 2500 + k,
				col: 0,
				kind: CellKind.EMPTY,
				value: 0,
				style: 0,
				formula: -1,
				origin: -1
			}))
		]

		for (const { row, col, kind, value, style, formula, origin } of writes) one.setCell(row, col, kind, value, style, formula, origin)

		many.setCells(writes)

		expect(many.cellCount).toBe(one.cellCount)

		for (const row of [2, 50, 99, 100, 107, 2099, 2100, 2549, 2550, 3099, 4000, 70_000]) {
			for (const col of [0, 3, 5]) {
				const a = one.locate(row, col)
				const b = many.locate(row, col)

				expect(b < 0).toBe(a < 0)

				if (a >= 0) expect([many.style(b), many.value(b), many.kind(b)]).toEqual([one.style(a), one.value(a), one.kind(a)])
			}
		}

		expect(many.blocks.every(block => block.rowCount <= BLOCK_ROWS)).toBe(true)
	})
})
