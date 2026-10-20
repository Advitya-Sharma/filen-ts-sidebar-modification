import { describe, expect, it } from "vitest"
import { XlsxWriteError } from "../core/xlsx/write/failure"
import { writeXlsx } from "../core/xlsx/write/save"
import { verifyXlsxParts } from "../core/xlsx/verify/verify"
import { buildFixture, openWorkbook } from "./xlsxWriteFixtures"

// The save's check counts each worksheet section against the model: a section written anew appears once (its
// markup holds every child of its name), a copied one as often as the source holds it.

const RULES =
	'<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;1</formula></cfRule></conditionalFormatting>' +
	'<conditionalFormatting sqref="B1"><cfRule type="expression" priority="2"><formula>B1&gt;1</formula></cfRule></conditionalFormatting>'

describe("the save check's section counts", () => {
	it("refuses a section written more often than the model has it", async () => {
		const workbook = await openWorkbook(buildFixture({ sheets: [{ name: "Data", body: `<sheetData/>${RULES}` }] }))

		await workbook.load(0)
		workbook.changes.sheetSection(0, "conditionalFormatting", () => RULES)

		const draft = await writeXlsx({ opc: workbook.opc, workbook })

		await expect(verifyXlsxParts(draft)).resolves.toBeUndefined()

		// The assembly as it was once: the markup written in place of each of the two sources.
		const sheet = draft.sheets[0]

		if (sheet === undefined) throw new Error("no sheet written")

		const doubled = { ...draft, sheets: [{ ...sheet, names: ["sheetData", "conditionalFormatting", "conditionalFormatting"] }] }
		const refusal = await verifyXlsxParts(doubled).then(
			() => null,
			(error: unknown) => error
		)

		expect(refusal).toBeInstanceOf(XlsxWriteError)
		expect((refusal as XlsxWriteError).reason).toBe("model")
		expect((refusal as XlsxWriteError).message).toContain("conditionalFormatting written 2 times")
	})

	it("accepts a copied section as often as the source holds it", async () => {
		const workbook = await openWorkbook(buildFixture({ sheets: [{ name: "Data", body: `<sheetData/>${RULES}` }] }))

		await workbook.load(0)
		workbook.changes.sheetRows(0)

		await expect(verifyXlsxParts(await writeXlsx({ opc: workbook.opc, workbook }))).resolves.toBeUndefined()
	})
})
