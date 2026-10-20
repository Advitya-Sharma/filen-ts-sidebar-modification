import { HyperFormula } from "hyperformula"
import { describe, expect, it } from "vitest"
import { fellBack } from "./fuzz/sheet/recalc"

// The recalculation oracle counts a cell as calculated only when the engine was given its formula: a value, or a
// formula that only holds the file's result, matches Excel by construction.

describe("the recalculation oracle's fallbacks", () => {
	it("counts values and held results as fallbacks, and formulas as calculated", () => {
		const hf = HyperFormula.buildFromArray([[1, "=FILEN.HELD(0,A1)", "=FILEN.HELD.ANY(1)", "=A1*2"]], { licenseKey: "gpl-v3" })
		const at = (col: number) => ({ sheet: 0, row: 0, col })

		expect([0, 1, 2, 3].map(col => fellBack(hf, at(col)))).toEqual([true, true, true, false])
		hf.destroy()
	})
})
