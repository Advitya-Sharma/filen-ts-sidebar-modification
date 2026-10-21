import { HyperFormula } from "hyperformula"
import { describe, expect, it } from "vitest"
import { pluginFunctionNames, registerSheetFunctions } from "../core/sheet/formula/functions"
import { sheetFunctions } from "../core/sheet/functions/catalogue"

// Read before the app's own functions are registered.
const BUILT_IN = new Set(HyperFormula.getRegisteredFunctionNames("enGB"))

// The catalogue's "hyperformula" support against the engine itself: exactly the names it registers.
describe("sheet function catalogue against HyperFormula", () => {
	it("is read from the engine version the app ships", () => {
		expect(HyperFormula.version).toBe("3.4.0")
	})

	it("marks as calculated by HyperFormula exactly the functions the engine registers", () => {
		for (const entry of sheetFunctions()) {
			expect(`${entry.name}: ${String(entry.support === "hyperformula")}`).toBe(`${entry.name}: ${String(BUILT_IN.has(entry.name))}`)
		}
	})

	it("marks as a plugin's exactly the functions the app's plugins add, and they are registered", () => {
		registerSheetFunctions()

		const registered = new Set(HyperFormula.getRegisteredFunctionNames("enGB"))
		const added = new Set(pluginFunctionNames().filter(name => !BUILT_IN.has(name) && !name.startsWith("FILEN.")))

		expect(
			sheetFunctions()
				.filter(entry => entry.support === "plugin")
				.map(entry => entry.name)
		).toEqual([...added].sort((a, z) => (a < z ? -1 : 1)))

		for (const name of pluginFunctionNames()) expect(registered.has(name), name).toBe(true)
	})
})
