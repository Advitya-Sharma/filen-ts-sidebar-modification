import { describe, expect, it } from "vitest"
import { FormulaEngine } from "../core/sheet/formula/index"
import type { SheetLocale } from "../core/sheet/protocol"

// Text read as a date inside a formula, as Excel reads it by the system's date settings and as a typed entry reads
// it here: support.microsoft.com "DATEVALUE function" gives =DATEVALUE("8/22/2011") as 40777 (en-US).

function results(formulas: readonly string[], locale?: SheetLocale): unknown[] {
	const engine = new FormulaEngine([{ name: "S", rows: [[...formulas]], formulas: [] }], [], false, undefined, undefined, locale)

	try {
		return formulas.map((_, col) => engine.result(0, 0, col).value)
	} finally {
		engine.destroy()
	}
}

describe("text dates in formulas", () => {
	it("read month first in en-US, and ISO dates", () => {
		expect(
			results([
				'=DATEVALUE("8/22/2011")',
				'=MONTH("5/1/2020")',
				'=DATEVALUE("2011-08-22")',
				'=MONTH("13/1/2020")',
				'=DATEVALUE("2/30/2020")'
			])
		).toEqual([40777, 5, 40777, "#VALUE!", "#VALUE!"])
	})

	it("read day first where the locale writes dates so, ISO dates still", () => {
		const de: SheetLocale = { decimal: ",", group: ".", argument: ";", dateOrder: "dmy" }

		expect(results(['=DATEVALUE("22.8.2011")', '=MONTH("5/1/2020")', '=DATEVALUE("2011-08-22")'], de)).toEqual([40777, 1, 40777])
	})

	it("read times alone and after a date, with AM and PM", () => {
		const [half, evening, both, bad] = results([
			'=TIMEVALUE("12:00")',
			'=TIMEVALUE("6:30 PM")',
			'=DATEVALUE("8/22/2011 6:30 PM")+TIMEVALUE("8/22/2011 6:30 PM")',
			'=TIMEVALUE("10:75")'
		])

		expect(half).toBe(0.5)
		expect(evening).toBeCloseTo(18.5 / 24, 12)
		expect(both).toBeCloseTo(40777 + 18.5 / 24, 9)
		expect(bad).toBe("#VALUE!")
	})

	it("read text in arithmetic as a date, as a typed date reads", () => {
		expect(results(['="1/2/2020"+1', '="2020-01-02"-1'])).toEqual([43833, 43831])
	})
})
