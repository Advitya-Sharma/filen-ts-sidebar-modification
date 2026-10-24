import { describe, expect, it } from "vitest"
import { BASIC, openFixture } from "./sheetSessionFixtures"

// An Excel 97–2003 workbook is opened from an .xlsx made of its values alone (the app's .xls reader keeps no
// formulas, number formats, sizes or names). A copy in another format says so, as Save as .xlsx does before it
// writes one: nothing of the source is dropped unnamed.

describe("Download as from an Excel 97–2003 workbook", () => {
	it.each(["xlsx", "ods"] as const)("names the values-only loss in its .%s copy", async format => {
		const { session } = await openFixture(BASIC, { kind: "xls" })

		try {
			const answer = await session.query("saveAs.format", { format })

			if (!answer.ok) throw new Error(answer.message)

			expect(answer.lossy.filter(feature => feature.kind === "valuesOnly")).toEqual([{ kind: "valuesOnly", count: 1, sheets: [] }])
		} finally {
			session.close()
		}
	})

	it("names no such loss for a workbook read with its formulas", async () => {
		const { session } = await openFixture(BASIC)

		try {
			const answer = await session.query("saveAs.format", { format: "xlsx" })

			expect(answer.ok && answer.lossy).toEqual([])
		} finally {
			session.close()
		}
	})
})
