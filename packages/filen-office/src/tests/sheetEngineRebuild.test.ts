import { describe, expect, it } from "vitest"
import { FormulaEngine } from "../core/sheet/formula/index"

// The engine built again with a far cell written (a formula reading a wide whole-column span held): its formulas read
// as before, whatever the sheets are named, and SUBTOTAL leaves out the rows hidden on the sheet it reads.

describe("the engine built again after a cell at the grid's far edge", () => {
	it("keeps formulas naming a sheet whose name needs quotes, and the hidden rows of the sheet SUBTOTAL reads", () => {
		const hidden = (sheet: number, row: number): number => (sheet === 1 && row === 1 ? 1 : 0)
		const engine = new FormulaEngine([{ name: "S", rows: [[null, "=SUM(C:Z)"]], formulas: [] }], [], false, () => undefined, hidden)

		try {
			engine.addSheet("2024")
			engine.set(1, [
				{ row: 0, col: 0, content: 1 },
				{ row: 1, col: 0, content: 2 },
				{ row: 2, col: 0, content: 3 }
			])
			engine.set(0, [
				{ row: 0, col: 0, content: "=SUBTOTAL(109,'2024'!A1:A3)" },
				{ row: 0, col: 2, content: "='2024'!A3*2" }
			])

			expect(engine.result(0, 0, 0)).toEqual({ value: 4 })

			engine.set(0, [{ row: 1_048_575, col: 25, content: 1 }])

			expect(engine.result(0, 0, 0)).toEqual({ value: 4 })
			expect(engine.result(0, 0, 2)).toEqual({ value: 6 })
		} finally {
			engine.destroy()
		}
	})
})
