import { afterEach, describe, expect, it, vi } from "vitest"
import { SheetStore } from "../core/xlsx/model/store"
import { cells, firstTile, tileEntry } from "./sheetSessionFixtures"
import { openKeys, sheetBody } from "./sheetKeysFixtures"

// An edit writing many cells hands the store all of them at once, each block rebuilt once, where one write after
// another copies a block for every row added; undo and redo write them back the same way.

afterEach(() => {
	vi.restoreAllMocks()
})

function spied(): { one: () => number; many: () => number } {
	const one = vi.spyOn(SheetStore.prototype, "setCell")
	const many = vi.spyOn(SheetStore.prototype, "setCells")

	return { one: () => one.mock.calls.length, many: () => many.mock.calls.length }
}

const ROWS = 40

describe("batched cell writes", () => {
	it("fills a column with one store write, and undoes and redoes it with one each", async () => {
		const { session } = await openKeys({ sheets: [{ name: "S", body: sheetBody([[1], [2]]) }] })
		const store = spied()
		const target = { startRow: 0, startCol: 0, endRow: ROWS - 1, endCol: 0 }
		const result = await session.command({ type: "fill.auto", sheet: 0, source: { ...target, endRow: 1 }, target, mode: "auto" })

		expect(result.ok).toBe(true)
		expect([store.one(), store.many()]).toEqual([0, 1])
		expect(tileEntry(await firstTile(session, 0), ROWS - 1, 0)?.text).toBe(String(ROWS))

		await session.undo()
		expect([store.one(), store.many()]).toEqual([0, 2])
		expect(tileEntry(await firstTile(session, 0), ROWS - 1, 0)).toBeNull()

		await session.redo()
		expect([store.one(), store.many()]).toEqual([0, 3])
		expect(tileEntry(await firstTile(session, 0), ROWS - 1, 0)?.text).toBe(String(ROWS))
	})

	it("sets and clears many cells with one store write each, a lone cell through the single-cell path", async () => {
		const { session } = await openKeys({ sheets: [{ name: "S", body: sheetBody([[1]]) }] })
		const store = spied()
		const many: [number, number, string][] = Array.from({ length: ROWS }, (_, row) => [row, 1, String(row + 1)])

		expect((await session.command({ type: "cells.set", sheet: 0, cells: cells(many) })).ok).toBe(true)
		expect([store.one(), store.many()]).toEqual([0, 1])
		expect(tileEntry(await firstTile(session, 0), ROWS - 1, 1)?.text).toBe(String(ROWS))

		const ranges = [{ startRow: 0, startCol: 1, endRow: ROWS - 1, endCol: 1 }]

		expect((await session.command({ type: "cells.clear", sheet: 0, ranges, what: "contents" })).ok).toBe(true)
		expect([store.one(), store.many()]).toEqual([0, 2])
		expect(tileEntry(await firstTile(session, 0), 3, 1)).toBeNull()

		expect((await session.command({ type: "cells.set", sheet: 0, cells: cells([[0, 2, "7"]]) })).ok).toBe(true)
		expect([store.one(), store.many()]).toEqual([1, 2])
	})
})
