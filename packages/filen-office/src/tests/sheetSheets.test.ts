import { describe, expect, it } from "vitest"
import type { SheetChange, SheetCommandResult } from "../core/sheet/protocol"
import { unzipSync, zipSync } from "fflate"
import { encodeUtf8 } from "../core/xml/index"
import { buildFixture, parts, type Fixture } from "./xlsxWriteFixtures"
import { ZIP_ENTRY_TIME } from "./zipTime"
import { firstTile, openFixture, savedBytes, tileEntry } from "./sheetSessionFixtures"

// Sheets deleted, moved, hidden and shown, and tab colours: what the saved workbook holds, and that an undo gives
// the source back.

function changeOf(result: SheetCommandResult): SheetChange {
	if (!result.ok)
		throw new Error(`refused: ${result.refusal.reason}${result.refusal.detail === undefined ? "" : ` (${result.refusal.detail})`}`)

	return result.change
}

const THREE = {
	sheets: [
		{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' },
		{ name: "Two", body: '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>' },
		{
			name: "Three",
			body: '<sheetData><row r="1"><c r="A1"><f>Two!A1+1</f><v>3</v></c><c r="B1"><f>SUM(One:Two!A1)</f><v>3</v></c></row></sheetData>'
		}
	],
	workbookTail:
		'<definedNames><definedName name="Local" localSheetId="1">Two!$A$1</definedName><definedName name="Global">Two!$A$1</definedName></definedNames>'
}

describe("sheet session: sheets deleted", () => {
	it("deletes a sheet, turns references to it into #REF! and drops its local names, and undoes to the source", async () => {
		const { session, info, bytes } = await openFixture(THREE)
		const [one, two, three] = info.sheets.map(sheet => sheet.id)
		const change = changeOf(await session.command({ type: "sheet.delete", sheet: two! }))

		expect(change.sheets?.map(sheet => sheet.name)).toEqual(["One", "Three"])

		const saved = parts(await savedBytes(session))

		expect(saved["xl/worksheets/sheet2.xml"]).toBeUndefined()
		expect(saved["xl/workbook.xml"]).not.toContain('name="Two"')
		expect(saved["xl/workbook.xml"]).not.toContain('name="Local"')
		expect(saved["xl/workbook.xml"]).toContain('<definedName name="Global">#REF!$A$1</definedName>')
		expect(saved["xl/worksheets/sheet3.xml"]).toContain("<f>#REF!A1+1</f>")
		// A 3-D span losing its last sheet narrows to the one before it.
		expect(saved["xl/worksheets/sheet3.xml"]).toContain("<f>SUM('One'!A1)</f>")
		expect(saved["[Content_Types].xml"]).not.toContain("sheet2.xml")
		expect(tileEntry(await firstTile(session, three!), 0, 0)?.input).toBe("=#REF!A1+1")

		changeOf(await session.undo())

		const undone = parts(await savedBytes(session))
		const source = parts(bytes)

		for (const name of Object.keys(source)) expect(undone[name], name).toBe(source[name])

		expect(one).toBeDefined()
		session.close()
	})

	it("refuses to delete the last visible sheet", async () => {
		const { session, info } = await openFixture({
			sheets: [
				{ name: "One", body: "<sheetData/>" },
				{ name: "Two", body: "<sheetData/>" }
			],
			workbookTail: ""
		})

		changeOf(await session.command({ type: "sheet.state", sheet: info.sheets[1]!.id, state: "hidden" }))

		const result = await session.command({ type: "sheet.delete", sheet: info.sheets[0]!.id })

		expect(result.ok ? null : result.refusal.reason).toBe("lastSheet")
		session.close()
	})
})

describe("sheet session: sheets moved", () => {
	it("moves a sheet, writes a 3-D span whose ends swapped in sheet order, and keeps local names on their sheet", async () => {
		const { session, info, bytes } = await openFixture(THREE)
		const one = info.sheets[0]!.id

		changeOf(await session.command({ type: "sheet.move", sheet: one, to: 2 }))

		const saved = parts(await savedBytes(session))
		const workbook = saved["xl/workbook.xml"] ?? ""

		expect(workbook.indexOf('name="Two"')).toBeLessThan(workbook.indexOf('name="One"'))
		expect(workbook).toContain('<definedName name="Local" localSheetId="0">Two!$A$1</definedName>')
		expect(saved["xl/worksheets/sheet3.xml"]).toContain("<f>SUM('Two:One'!A1)</f>")

		changeOf(await session.undo())
		expect(parts(await savedBytes(session))["xl/workbook.xml"]).toBe(parts(bytes)["xl/workbook.xml"])
		session.close()
	})
})

describe("sheet session: sheets hidden and shown", () => {
	it("hides the active sheet, moving the active tab to the next visible one, and shows it again", async () => {
		const { session, info, bytes } = await openFixture({
			sheets: [
				{ name: "One", body: '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/>' },
				{ name: "Two", body: '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/>' }
			]
		})
		const one = info.sheets[0]!.id
		const change = changeOf(await session.command({ type: "sheet.state", sheet: one, state: "hidden" }))

		expect(change.sheets?.find(sheet => sheet.id === one)?.state).toBe("hidden")

		const saved = parts(await savedBytes(session))

		expect(saved["xl/workbook.xml"]).toContain('<sheet name="One" sheetId="1" r:id="rId1" state="hidden"/>')
		expect(saved["xl/workbook.xml"]).toMatch(/activeTab="1"/)
		expect(saved["xl/worksheets/sheet1.xml"]).not.toContain('tabSelected="1"')
		expect(saved["xl/worksheets/sheet2.xml"]).toContain('tabSelected="1"')

		changeOf(await session.command({ type: "sheet.state", sheet: one, state: "visible" }))
		expect(parts(await savedBytes(session))["xl/workbook.xml"]).toContain('<sheet name="One" sheetId="1" r:id="rId1"/>')

		changeOf(await session.undo())
		changeOf(await session.undo())

		const undone = parts(await savedBytes(session))

		for (const name of ["xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"])
			expect(undone[name], name).toBe(parts(bytes)[name])

		session.close()
	})

	it("refuses to hide the last visible sheet, and refuses under workbook structure protection", async () => {
		const single = await openFixture({ sheets: [{ name: "One", body: "<sheetData/>" }] })
		const hidden = await single.session.command({ type: "sheet.state", sheet: single.info.sheets[0]!.id, state: "hidden" })

		expect(hidden.ok ? null : hidden.refusal.reason).toBe("lastSheet")
		single.session.close()

		const locked = await openFixture({
			sheets: [
				{ name: "One", body: "<sheetData/>" },
				{ name: "Two", body: "<sheetData/>" }
			],
			workbookTail: '<workbookProtection lockStructure="1"/>'
		})
		const refused = await locked.session.command({ type: "sheet.state", sheet: locked.info.sheets[0]!.id, state: "hidden" })

		expect(refused.ok ? null : refused.refusal.reason).toBe("protectedWorkbook")
		locked.session.close()
	})
})

describe("sheet session: tab colours", () => {
	it("writes a tab colour into the sheet's properties and takes it off", async () => {
		const { session, info, bytes } = await openFixture({ sheets: [{ name: "One", body: "<sheetData/>" }] })
		const one = info.sheets[0]!.id
		const change = changeOf(await session.command({ type: "sheet.tabColor", sheet: one, color: 0xff0000 }))

		expect(change.sheets?.[0]?.tabColor).toBe(0xff0000)
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toContain('<tabColor rgb="FFFF0000"/>')

		changeOf(await session.command({ type: "sheet.tabColor", sheet: one, color: null }))
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toBe(parts(bytes)["xl/worksheets/sheet1.xml"])
		session.close()
	})
})

function withPart(fixture: Fixture, name: string, text: string): Uint8Array {
	const files = unzipSync(buildFixture(fixture))

	files[name] = Uint8Array.from(encodeUtf8(text))

	return new Uint8Array(zipSync(files, { mtime: ZIP_ENTRY_TIME }))
}

describe("sheet session: the sheet list after an undo", () => {
	it("writes the workbook's sheet list as it was, white space included, once a move is undone", async () => {
		const fixture = {
			sheets: [
				{ name: "One", body: "<sheetData/>" },
				{ name: "Two", body: "<sheetData/>" }
			]
		}
		const source = parts(buildFixture(fixture))
		const spaced = (source["xl/workbook.xml"] ?? "").replace("<sheets>", "<sheets>\n    ").replace("</sheets>", "\n  </sheets>")
		const { session, info, bytes } = await openFixture(withPart(fixture, "xl/workbook.xml", spaced))

		changeOf(await session.command({ type: "sheet.move", sheet: info.sheets[0]!.id, to: 1 }))
		changeOf(await session.undo())
		expect(parts(await savedBytes(session))["xl/workbook.xml"]).toBe(parts(bytes)["xl/workbook.xml"])
		session.close()
	})
})
