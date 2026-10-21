import { describe, expect, it } from "vitest"
import { CellFlag, type SheetLocale } from "../core/sheet/protocol"
import { firstTile, tileEntry } from "./sheetSessionFixtures"
import { openKeys, sheetBody } from "./sheetKeysFixtures"

// A text file's copy as clipboard HTML, pasted into a workbook: what the view shows comes across, numbers as numbers.

const GERMAN: SheetLocale = { decimal: ",", group: ".", argument: ";", dateOrder: "dmy" }

async function pasted(
	csv: string,
	locale?: SheetLocale
): Promise<{ text: (col: number) => string | undefined; number: (col: number) => boolean }> {
	const source = await openKeys(csv)

	if (locale !== undefined) source.session.configure({ locale })

	const range = { startRow: 0, startCol: 0, endRow: 0, endCol: 7 }
	const copy = await source.session.query("copy", { sheet: 0, range, cut: false })
	const { html } = await source.session.query("copyHtml", { sheet: 0, range })

	if ("refused" in copy || html === null) throw new Error("no copy")

	const book = await openKeys({ sheets: [{ name: "S", body: sheetBody([[]]) }] })

	if (locale !== undefined) book.session.configure({ locale })

	expect((await book.session.command({ type: "clipboard.pasteHtml", sheet: 0, at: { row: 0, col: 0 }, html, text: copy.text })).ok).toBe(
		true
	)

	const tile = await firstTile(book.session, 0)

	source.session.close()
	book.session.close()

	return {
		text: col => tileEntry(tile, 0, col)?.text,
		number: col => ((tileEntry(tile, 0, col)?.flags ?? 0) & CellFlag.NUMBER) !== 0
	}
}

describe("a text file's copy pasted into a workbook", () => {
	it("pastes the numbers it shows as numbers, and text it shows as text", async () => {
		const cells = await pasted("42,1.5,abc,=1+1,1e5,12%,'7,2024-01-15\n")

		expect([0, 1, 2, 3, 4, 5, 6, 7].map(cells.number)).toEqual([true, true, false, false, false, true, false, true])
		expect(cells.text(3)).toBe("=1+1")
		expect(cells.text(4)).toBe("1e5")
		expect(cells.text(5)).toBe("12%")
		expect(cells.text(6)).toBe("'7")
	})

	it("reads the numbers in the session's locale", async () => {
		const cells = await pasted('"1,5","1.234,5",x\n', GERMAN)

		expect([0, 1, 2].map(cells.number)).toEqual([true, true, false])
		expect(cells.text(0)).toBe("1,5")
		// As typed into a workbook: grouping implies "#,##0.00".
		expect(cells.text(1)).toBe("1.234,50")
	})
})
