import { describe, expect, it } from "vitest"
import { FormulaEngine, type EngineSheet } from "../core/sheet/formula/index"

// A range where one value is wanted (=B:B in an ordinary formula, written FILEN.INTERSECT(B:B) for the engine) reads
// the one cell in the formula's row: never the range's every value, which made each such formula cost the range.

describe("implicit intersection", () => {
	it("reads one cell per formula, so thousands of them over a whole column recalculate quickly", { timeout: 60_000 }, () => {
		const rows: EngineSheet["rows"] = []

		for (let row = 0; row < 4_000; row++) rows.push(["=FILEN.INTERSECT(B:B)", row])

		const engine = new FormulaEngine([{ name: "S", rows, formulas: [] }], [])

		try {
			expect(engine.result(0, 3_999, 0)).toEqual({ value: 3_999 })

			const start = performance.now()
			const changes = engine.set(0, [{ row: 2_000, col: 1, content: -1 }])

			expect(performance.now() - start).toBeLessThan(500)
			expect(changes).toContainEqual({ sheet: 0, row: 2_000, col: 0, value: -1 })
		} finally {
			engine.destroy()
		}
	})

	it("gives #VALUE! where the formula's row or column misses the range, and the cell of a row or a column", () => {
		const engine = new FormulaEngine(
			[
				{
					name: "S",
					rows: [
						[1, 2, 3],
						[4, null, "=FILEN.INTERSECT(A1:C1)"],
						["=FILEN.INTERSECT(A1:A2)", "=FILEN.INTERSECT(B5:B9)"]
					],
					formulas: []
				}
			],
			[]
		)

		try {
			expect(engine.result(0, 1, 2)).toEqual({ value: 3 })
			expect(engine.result(0, 2, 0)).toEqual({ value: "#VALUE!", error: true })
			expect(engine.result(0, 2, 1)).toEqual({ value: "#VALUE!", error: true })
		} finally {
			engine.destroy()
		}
	})
})
