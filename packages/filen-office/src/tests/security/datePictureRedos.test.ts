import { describe, expect, it } from "vitest"
import { parseControlDate } from "../../core/docx/forms/dates"

// A date control's picture comes from the document and becomes a pattern for reading typed text back: no
// picture may make that read take more than linear-ish time.

function timed(run: () => unknown): number {
	const started = performance.now()

	run()

	return performance.now() - started
}

describe("date pictures from documents", () => {
	it("reads back text against a picture of many adjacent numeric fields quickly", () => {
		// Exponential backtracking: 13 adjacent day/month groups took seconds, each more one about four times as long.
		expect(timed(() => parseControlDate(`${"1".repeat(52)}1x`, "dM".repeat(13), "en-US"))).toBeLessThan(250)
	})

	it("reads back text against a picture of many whitespace runs quickly", () => {
		expect(timed(() => parseControlDate(`1${" ".repeat(30)}x`, `d${" \t".repeat(12)}M`, "en-US"))).toBeLessThan(250)
	})

	it("reads back text against a picture alternating quoted and bare spaces quickly", () => {
		// Quoted spaces were literal spaces between \s* groups: a run of spaces before a non-digit split every way.
		expect(timed(() => parseControlDate(`1${" ".repeat(42)}x`, `d${" ' '".repeat(14)} d`, "en-US"))).toBeLessThan(250)
		expect(parseControlDate("16 de octubre de 2026", "d' de 'MMMM' de 'yyyy", "es-ES")).toBe("2026-10-16")
	})

	it("still reads the pictures Word writes", () => {
		expect(parseControlDate("16/10/2026", "dd/MM/yyyy", "en-GB")).toBe("2026-10-16")
		expect(parseControlDate("10/16/2026", "M/d/yyyy", "en-US")).toBe("2026-10-16")
		expect(parseControlDate("16. Oktober 2026", "d. MMMM yyyy", "de-DE")).toBe("2026-10-16")
		expect(parseControlDate("Friday, October 16, 2026", "dddd, MMMM d, yyyy", "en-US")).toBe("2026-10-16")
		expect(parseControlDate("16  10   2026", "d M yyyy", "en-US")).toBe("2026-10-16")
	})
})
