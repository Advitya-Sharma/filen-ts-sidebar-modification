import { describe, expect, it } from "vitest"
import { TILE_ROWS, type SheetCommand, type SheetSessionApi } from "../core/sheet/protocol"
import { cells, savedBytes, tileEntry } from "./sheetSessionFixtures"
import { openKeys, sheetBody } from "./sheetKeysFixtures"

// A large paste, fill or clear in the in-thread host runs in slices, yielding to the page between them: the
// session answers nothing else meanwhile (a tile asked for during it shows the whole edit), the edit is one step of
// history, and what it, its undo and its redo leave is exactly what the same edit unsliced leaves.

const ROWS = 20_000

// A clock whose slice is up at every reading: every stage of a large edit yields at each look at it.
function counting(): { slicing: { sliceMs: number; now: () => number; pause: () => Promise<void> }; pauses: () => number } {
	let ticks = 0
	let pauses = 0

	return {
		slicing: {
			sliceMs: 1,
			now: () => ticks++,
			pause: () => {
				pauses++

				return new Promise(resolve => setTimeout(resolve, 0))
			}
		},
		pauses: () => pauses
	}
}

// Two columns of numbers and a total over the first.
const BODY = sheetBody([[1, 2, { f: `SUM(A1:A${String(ROWS + 1)})`, v: 1 }], ...Array.from({ length: ROWS }, (_, k) => [k + 2, k % 5])])

const column = (col: number, input: (row: number) => string): ReturnType<typeof cells> =>
	cells(Array.from({ length: ROWS }, (_, k) => [k + 1, col, input(k)] as [number, number, string]))

const EDITS: readonly [string, SheetCommand][] = [
	["a paste into new cells", { type: "cells.set", sheet: 0, cells: column(3, k => String(k % 997)) }],
	["a paste over cells", { type: "cells.set", sheet: 0, cells: column(0, k => String(k % 13)) }],
	["a fill", { type: "fill.copy", sheet: 0, ranges: [{ startRow: 0, startCol: 0, endRow: ROWS, endCol: 1 }], direction: "down" }],
	["a clear", { type: "cells.clear", sheet: 0, ranges: [{ startRow: 1, startCol: 0, endRow: ROWS, endCol: 1 }], what: "contents" }]
]

// What the first four columns show in a few rows, from one tile request.
async function shown(session: SheetSessionApi, rows: readonly number[]): Promise<(string | undefined)[][]> {
	const { tiles } = await session.tiles(0, Uint32Array.from(rows.flatMap(row => [Math.floor(row / TILE_ROWS), 0])))

	return rows.map((row, k) => [0, 1, 2, 3].map(col => (tiles[k] === undefined ? undefined : tileEntry(tiles[k], row, col)?.text)))
}

describe("large edits in slices", () => {
	it.each(EDITS)("%s yields between slices and leaves what the unsliced edit leaves", { timeout: 120_000 }, async (_name, command) => {
		const plain = await openKeys({ sheets: [{ name: "S", body: BODY }] })
		const clock = counting()
		const sliced = await openKeys({ sheets: [{ name: "S", body: BODY }] }, { slicing: clock.slicing })

		for (const { session } of [plain, sliced]) await session.intent()

		const before = await savedBytes(sliced.session)
		const paused = clock.pauses()
		const result = sliced.session.command(command)
		// Sent while the edit runs: answered once it is done.
		const during = shown(sliced.session, [ROWS, 70, 0])

		expect((await result).ok).toBe(true)
		expect(clock.pauses() - paused).toBeGreaterThan(10)
		expect((await plain.session.command(command)).ok).toBe(true)
		expect(await during).toEqual(await shown(plain.session, [ROWS, 70, 0]))

		const after = await savedBytes(plain.session)

		expect(await savedBytes(sliced.session)).toEqual(after)

		expect((await sliced.session.undo()).ok).toBe(true)
		expect(await savedBytes(sliced.session)).toEqual(before)
		const undone = clock.pauses()

		expect((await sliced.session.redo()).ok).toBe(true)
		expect(clock.pauses() - undone).toBeGreaterThan(3)
		expect(await savedBytes(sliced.session)).toEqual(after)

		for (const { session } of [plain, sliced]) session.close()
	})
})
