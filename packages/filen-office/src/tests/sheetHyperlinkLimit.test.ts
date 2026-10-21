import { describe, expect, it } from "vitest"
import { MAX_SHEET_LINKS } from "../core/xlsx/hyperlinks/edit"
import { buildFixture } from "./sheetSessionFixtures"
import { changeOf, openAnnotated, range, refusalOf } from "./sheetNotesFixtures"

// Excel's "Excel specifications and limits": 65,530 hyperlinks in a worksheet, counted as the sheet has them after
// the edit (a link cut through the middle stays as its pieces).

const SINGLES = MAX_SHEET_LINKS - 2

function links(): string {
	let markup = '<hyperlink ref="A1:C3" location="A1"/>'

	for (let k = 0; k < SINGLES; k++) markup += `<hyperlink ref="D${String(k + 10)}" location="A1"/>`

	return `<sheetData/><hyperlinks>${markup}</hyperlinks>`
}

const FULL = buildFixture({ sheets: [{ name: "One", body: links() }] })

describe("hyperlinks in a worksheet", () => {
	it("is 65,530 at most", () => {
		expect(MAX_SHEET_LINKS).toBe(65_530)
	})

	it("refuses a link whose cut leaves the sheet past the limit, and takes one up to it", async () => {
		const { session, sheets } = await openAnnotated(FULL)
		const sheet = sheets[0] ?? 0
		const target = { location: "A1" }

		// B2 cuts A1:C3 into four pieces: SINGLES + 4 + 1 links.
		expect(refusalOf(await session.command({ type: "link.set", sheet, ranges: [range(1, 1)], target }))).toBe("tooLarge")
		// A new link on E1: SINGLES + 1 + 1, the limit itself.
		changeOf(await session.command({ type: "link.set", sheet, ranges: [range(0, 4)], target }))
		expect(refusalOf(await session.command({ type: "link.set", sheet, ranges: [range(0, 5)], target }))).toBe("tooLarge")
		session.close()
	})
})
