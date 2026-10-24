import { describe, expect, it } from "vitest"
import { numericPicture } from "../../core/docx/fields/evaluate"
import { budget } from "../timing"

// A field's \# picture comes from the document: no picture makes an update throw or take long.

describe("numeric pictures from documents", () => {
	it("formats more decimal places than toFixed allows instead of throwing", () => {
		const places = 150

		expect(numericPicture(1.5, `0.${"0".repeat(places)}`)).toBe(`1.5${"0".repeat(places - 1)}`)
	})

	it("groups a very long run of digits in linear time", () => {
		const started = performance.now()
		const out = numericPicture(1, `#,${"0".repeat(60_000)}`)

		expect(performance.now() - started).toBeLessThan(budget(500))
		expect(out.replace(/[ ,]/g, "")).toBe(`${"0".repeat(59_999)}1`)
		expect(
			out
				.split(",")
				.slice(1)
				.every(group => group.length === 3)
		).toBe(true)
	})

	it("still groups and rounds as before", () => {
		expect(numericPicture(1234567.891, "#,##0.00")).toBe("1,234,567.89")
		expect(numericPicture(-12, "#,##0")).toBe("-  12")
		expect(numericPicture(1000, "#,##0")).toBe("1,000")
		expect(numericPicture(1e21, "#,##0")).toBe("1e+21")
		expect(numericPicture(5, "000,000")).toBe("000,005")
		expect(numericPicture(1.5, "0.000")).toBe("1.500")
	})
})
