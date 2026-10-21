import { describe, expect, it } from "vitest"
import { CellKind, RowFlag, blockOrigin } from "../core/xlsx/model/types"
import { BLOCK_CELLS, BLOCK_ROWS, SheetStore, StoreBuilder } from "../core/xlsx/read/index"
import { FormulaTable } from "../core/xlsx/model/formulas"
import { SheetStrings } from "../core/xlsx/model/strings"
import type { CellWrite } from "../core/xlsx/model/store"

// Many cells written at once (a paste, a fill, an undo): the store holds what one setCell per cell gives it, every
// row and cell with its flags, spans and source identity, whatever rows the writes add, empty or fall before the
// sheet's first; the blocks keep their limits.

// A store of `rows` rows from `firstRow`, three cells each, every fifth row with a formula and every seventh hidden.
function storeOf(rows: number, firstRow: number, step = 1): SheetStore {
	const builder = new StoreBuilder()

	for (let k = 0; k < rows; k++) {
		const row = firstRow + k * step

		builder.beginRow()

		for (const col of [0, 2, 5])
			builder.cell(col, CellKind.NUMBER, row * 10 + col, col === 2 ? 1 : 0, row % 5 === 0 && col === 5 ? 0 : -1)

		builder.endRow(row, row * 3, row * 3 + 40, row % 7 === 0 ? RowFlag.HIDDEN : 0, row % 11 === 0 ? 18 : Number.NaN, -1)
	}

	return new SheetStore(builder.finish(), new SheetStrings(), new FormulaTable())
}

// Everything the store holds, row by row, whatever its blocks.
function dump(store: SheetStore): unknown[] {
	const out: unknown[] = []

	for (const block of store.blocks) {
		for (let i = 0; i < block.rowCount; i++) {
			const cells: number[][] = []

			for (let k = block.rowStarts[i] ?? 0; k < (block.rowStarts[i + 1] ?? 0); k++)
				cells.push([
					block.cols[k] ?? 0,
					block.kinds[k] ?? 0,
					block.values[k] ?? 0,
					block.styles[k] ?? 0,
					block.formulas?.[k] ?? -1,
					blockOrigin(block, i, k)
				])

			out.push([
				block.firstRow + (block.rows[i] ?? 0),
				block.rowFlags[i],
				block.rowHeights[i],
				block.rowStyles[i],
				block.rowSpans[i * 2],
				block.rowSpans[i * 2 + 1],
				cells
			])
		}
	}

	return [store.cellCount, store.rowExtent, store.colExtent, out]
}

function limits(store: SheetStore): boolean {
	return store.blocks.every(
		block =>
			block.rowCount > 0 &&
			block.rows[0] === 0 &&
			block.rowCount <= BLOCK_ROWS &&
			(block.cellCount <= BLOCK_CELLS || block.rowCount === 1)
	)
}

// A deterministic generator (mulberry32).
function random(seed: number): () => number {
	let state = seed

	return () => {
		state = (state + 0x6d2b79f5) | 0

		let t = Math.imul(state ^ (state >>> 15), 1 | state)

		t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t

		return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296
	}
}

function write(row: number, col: number, kind: CellKind, value: number, style = 0, formula = -1, origin = -1): CellWrite {
	return { row, col, kind, value, style, formula, origin }
}

function sequential(store: SheetStore, writes: readonly CellWrite[]): void {
	for (const cell of writes) store.setCell(cell.row, cell.col, cell.kind, cell.value, cell.style, cell.formula, cell.origin)
}

describe("cells written many at once", () => {
	it(
		"equals one setCell per cell over random writes: added, changed and removed cells, new and emptied rows",
		{ timeout: 120_000 },
		() => {
			for (let seed = 1; seed <= 12; seed++) {
				const rand = random(seed)
				const one = storeOf(1500, 50, 2)
				const many = storeOf(1500, 50, 2)
				const writes: CellWrite[] = []
				const named = new Set<string>()

				for (let n = 0; n < 2500; n++) {
					const row = Math.floor(rand() * 4200)
					const col = Math.floor(rand() * 8)
					const key = `${row},${col}`

					if (named.has(key)) continue

					named.add(key)

					const pick = rand()

					writes.push(
						pick < 0.25
							? write(row, col, CellKind.EMPTY, 0)
							: pick < 0.35
								? write(row, col, CellKind.NUMBER, n, 0, 0)
								: pick < 0.45
									? write(row, col, CellKind.NUMBER, n, 2, -1, 1 + (n % 3))
									: write(row, col, CellKind.NUMBER, n, n % 4)
					)
				}

				// Out of order, as an undo's log gives them.
				if (seed % 2 === 0) writes.reverse()

				sequential(one, writes)
				many.setCells(writes)

				expect(dump(many), `seed ${String(seed)}`).toEqual(dump(one))
				expect(limits(many), `seed ${String(seed)}`).toBe(true)
			}
		}
	)

	it("writes a long column into empty rows as full blocks, and empties them again", { timeout: 120_000 }, () => {
		const one = storeOf(10, 0)
		const many = storeOf(10, 0)
		const column = Array.from({ length: 4000 }, (_, k) => write(k + 1, 11, CellKind.NUMBER, k % 1000))

		sequential(one, column)
		many.setCells(column)
		expect(dump(many)).toEqual(dump(one))
		expect(limits(many)).toBe(true)
		expect(many.blocks.length).toBe(Math.ceil(4001 / BLOCK_ROWS))

		const cleared = column.map(cell => write(cell.row, cell.col, CellKind.EMPTY, 0))

		sequential(one, cleared)
		many.setCells(cleared)
		expect(dump(many)).toEqual(dump(one))
		expect(many.blocks.length).toBe(1)
	})

	it("writes 100,000 cells of a column into blocks within their limits", () => {
		const store = storeOf(10, 0)

		store.setCells(Array.from({ length: 100_000 }, (_, k) => write(k + 1, 11, CellKind.NUMBER, k % 1000)))
		expect(limits(store)).toBe(true)
		expect([store.cellCount, store.rowExtent, store.value(store.locate(77_777, 11))]).toEqual([100_030, 100_001, 77_776 % 1000])
	})

	it("writes rows before the sheet's first, and a row of more cells than a block holds", { timeout: 120_000 }, () => {
		const one = storeOf(20, 5000)
		const many = storeOf(20, 5000)
		const writes = [
			...Array.from({ length: 30 }, (_, k) => write(k * 3, 1, CellKind.NUMBER, k)),
			...Array.from({ length: BLOCK_CELLS + 10 }, (_, k) => write(5005, 6 + k, CellKind.NUMBER, k))
		]

		sequential(one, writes)
		many.setCells(writes)
		expect(dump(many)).toEqual(dump(one))
		expect(limits(many)).toBe(true)
	})
})
