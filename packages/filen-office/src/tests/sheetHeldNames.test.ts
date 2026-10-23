import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { CellFlag } from "../core/sheet/protocol"
import { HeldResult } from "../core/sheet/formula/formulaEngine"
import { RenderStyleTable } from "../core/sheet/session"
import { XlsxDocument } from "../core/xlsx/edit/document"
import { openXlsx } from "../core/xlsx/read/index"
import { buildFixture } from "./xlsxWriteFixtures"

async function document(workbookTail: string, body: string) {
	const opened = await openXlsx(bytesSource(buildFixture({ sheets: [{ name: "Data", body }], workbookTail })))

	if (!opened.ok) throw new Error(opened.reason)

	await opened.workbook.loadAll(0)

	const sheet = opened.workbook.sheets[0]?.model

	if (sheet === undefined || sheet === null) throw new Error("no sheet")

	const opening = new XlsxDocument(opened.workbook, new RenderStyleTable(), { kind: "xlsx" })

	opening.buildEngine()

	return { document: opening, sheet }
}

describe("defined names over huge areas", () => {
	it("holds a formula reading a name over more than four million cells at its cached result, flagged as cached", async () => {
		// The cached 7 differs from what the sheet's one value gives: only a held result shows it.
		const { document: opened, sheet } = await document(
			'<definedNames><definedName name="Huge">Data!$A$1:$XFD$21949</definedName></definedNames>',
			'<sheetData><row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>SUM(Huge)</f><v>7</v></c></row></sheetData>'
		)

		expect(opened.engineCell(sheet, 0, 1).content).toEqual(new HeldResult(7))
		expect(opened.engine?.result(0, 0, 1).value).toBe(7)

		const view = opened.cellView(sheet.id, 0, 1)

		expect(view.text).toBe("7")
		expect(view.flags & CellFlag.CACHED).toBe(CellFlag.CACHED)
		opened.dropEngine()
	})

	it("calculates a formula reading a name over a small area, not flagged as cached", async () => {
		const { document: opened, sheet } = await document(
			'<definedNames><definedName name="Small">Data!$A$1:$A$2</definedName></definedNames>',
			'<sheetData><row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>SUM(Small)</f><v>7</v></c></row></sheetData>'
		)

		expect(opened.engineCell(sheet, 0, 1).content).not.toBeInstanceOf(HeldResult)
		expect(opened.engine?.result(0, 0, 1).value).toBe(3)
		expect(opened.cellView(sheet.id, 0, 1).flags & CellFlag.CACHED).toBe(0)
		opened.dropEngine()
	})
})
