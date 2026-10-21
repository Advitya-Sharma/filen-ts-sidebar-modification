import { HyperFormula } from "hyperformula"
import { describe, expect, it, vi } from "vitest"
import { FormulaEngine, type EngineSheet, type RecalculatedCell } from "../core/sheet/formula/index"
import { SliceClock } from "../core/slicing"

// Sets made while evaluation is held only write: the release recalculates once and reports what one set of all the
// cells reports, reading the changed cells back between turns of the caller's clock.

function sheet(rows: number): EngineSheet {
	const data: EngineSheet["rows"] = Array.from({ length: rows }, (_, r) => [r, `=A${String(r + 1)}*2`])

	data[0]?.push(`=SUM(A1:A${String(rows)})`)

	return { name: "S", rows: data, formulas: [] }
}

const sorted = (cells: readonly RecalculatedCell[]): string[] =>
	cells.map(cell => `${String(cell.row)},${String(cell.col)}=${String(cell.value)}`).sort()

describe("evaluation held for a sliced write", () => {
	it("reports on release what one set of every cell reports, recalculating once", async () => {
		const once = new FormulaEngine([sheet(2_000)], [])
		const held = new FormulaEngine([sheet(2_000)], [])
		const cells = Array.from({ length: 2_000 }, (_, r) => ({ row: r, col: 0, content: 2_000 - r }))

		try {
			const expected = once.set(0, cells)

			held.holdEvaluation()

			for (let at = 0; at < cells.length; at += 256) expect(held.set(0, cells.slice(at, at + 256))).toEqual([])

			expect(sorted(await held.releaseEvaluation())).toEqual(sorted(expected))
			expect(held.result(0, 0, 2)).toEqual(once.result(0, 0, 2))
		} finally {
			once.destroy()
			held.destroy()
		}
	})

	it("reads back the changed cells in turns of the clock, and none when not asked to report", async () => {
		const engine = new FormulaEngine([sheet(30_000)], [])
		let time = 0
		let turns = 0
		const clock = new SliceClock({
			sliceMs: 1,
			now: () => (time += 1),
			pause: () => {
				turns++

				return Promise.resolve()
			}
		})

		try {
			engine.holdEvaluation()
			engine.set(
				0,
				Array.from({ length: 30_000 }, (_, r) => ({ row: r, col: 0, content: r + 1 }))
			)

			expect((await engine.releaseEvaluation(true, clock)).length).toBeGreaterThan(60_000)
			expect(turns).toBeGreaterThan(5)

			engine.holdEvaluation()
			engine.set(0, [{ row: 0, col: 0, content: 7 }])

			expect(await engine.releaseEvaluation(false)).toEqual([])
			expect(engine.result(0, 0, 1)).toEqual({ value: 14 })
		} finally {
			engine.destroy()
		}
	})

	// A formula reading its own cell is a circle Excel does not calculate: held at its stored result as it is written,
	// so the release recalculates once (not again over everything that reads it once the engine finds the circle).
	it("recalculates once on release when a formula reading its own cell is written", async () => {
		const rows: EngineSheet["rows"] = Array.from({ length: 50 }, (_, r) => [r + 1, `=A${String(r + 1)}+B${String(Math.max(1, r))}`])
		const engine = new FormulaEngine([{ name: "S", rows, formulas: [] }], [], false, () => undefined)

		try {
			const resumes = vi.spyOn(HyperFormula.prototype, "resumeEvaluation")

			engine.holdEvaluation()
			engine.set(
				0,
				rows.map((row, r) => ({ row: r, col: 1, content: row[1] ?? null, fallback: 5 }))
			)
			await engine.releaseEvaluation()

			expect(resumes).toHaveBeenCalledTimes(1)
			expect([engine.result(0, 0, 1), engine.result(0, 1, 1)]).toEqual([{ value: 5 }, { value: 7 }])
		} finally {
			engine.destroy()
		}
	})
})
