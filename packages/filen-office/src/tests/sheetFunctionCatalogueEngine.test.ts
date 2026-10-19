import { HyperFormula } from "hyperformula"
import { describe, expect, it } from "vitest"
import { sheetFunctions } from "../core/sheet/functions/catalogue"

// The catalogue's "hyperformula" support against the engine itself: exactly the names it registers.
describe("sheet function catalogue against HyperFormula", () => {
	it("is read from the engine version the app ships", () => {
		expect(HyperFormula.version).toBe("3.4.0")
	})

	it("marks as calculated by HyperFormula exactly the functions the engine registers", () => {
		const registered = new Set(HyperFormula.getRegisteredFunctionNames("enGB"))

		for (const entry of sheetFunctions()) {
			expect(`${entry.name}: ${String(entry.support === "hyperformula")}`).toBe(
				`${entry.name}: ${String(registered.has(entry.name))}`
			)
		}
	})
})
