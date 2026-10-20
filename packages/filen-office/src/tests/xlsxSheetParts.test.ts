import { unzipSync, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { saveXlsx } from "../core/xlsx/write/save"
import { CellKind, type XlsxSheetEntry } from "../core/xlsx/model/types"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { buildFixture, openFixture } from "./sheetSessionFixtures"
import { openWorkbook } from "./xlsxWriteFixtures"
import { ZIP_ENTRY_TIME } from "./zipTime"

// A sheet's part is its own and a sheet part: a sheet sharing its part with another, or related to a part of
// another kind, would take edits a save then loses. Such a workbook opens view-only, and the gate refuses to
// write it.

function retarget(from: string, to: string): Uint8Array {
	const files = unzipSync(
		buildFixture({
			sheets: [
				{ name: "One", body: '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>' },
				{ name: "Two", body: '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>' }
			]
		})
	)

	const rels = decodeUtf8(files["xl/_rels/workbook.xml.rels"]!).replace(`Target="${from}"`, `Target="${to}"`)

	return zipSync({ ...files, "xl/_rels/workbook.xml.rels": encodeUtf8(rels) }, { mtime: ZIP_ENTRY_TIME })
}

describe("a sheet's part", () => {
	it.each([
		["another sheet's part", "worksheets/sheet1.xml"],
		["the styles part", "styles.xml"],
		["the workbook part", "workbook.xml"]
	])("makes the workbook view-only when it is %s", async (_, target) => {
		const { session } = await openFixture(retarget("worksheets/sheet2.xml", target))

		expect(await session.intent()).toEqual({ state: "viewOnly", reason: "lossy" })
		session.close()
	})

	it("is never written for two sheets at once", async () => {
		const workbook = await openWorkbook(retarget("worksheets/sheet2.xml", "worksheets/sheet1.xml"))
		const two = workbook.sheets[1] as { part: string } & XlsxSheetEntry

		// The open left the second sheet without a part; force the alias past it.
		expect(two.part).toBe("")
		two.part = workbook.sheets[0]!.part

		const model = await workbook.load(0)

		model.store.setCell(0, 1, CellKind.NUMBER, 5, 0, -1)
		workbook.changes.sheetRows(0)

		const result = await saveXlsx({ opc: workbook.opc, workbook })

		expect(result.ok ? "saved" : result.reason).toBe("relationship")
	})

	it("is added to a workbook whose relationships already name a part the package lacks", async () => {
		const files = unzipSync(
			buildFixture({ sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }] })
		)
		const rels = decodeUtf8(files["xl/_rels/workbook.xml.rels"]!).replace(
			"</Relationships>",
			'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/></Relationships>'
		)
		const { session } = await openFixture(
			zipSync({ ...files, "xl/_rels/workbook.xml.rels": encodeUtf8(rels) }, { mtime: ZIP_ENTRY_TIME })
		)

		await session.intent()

		const added = await session.command({ type: "sheet.add", name: "Two" })
		const saved = await session.save()

		session.close()
		expect(added.ok).toBe(true)
		expect(saved.ok ? "saved" : `${saved.reason}: ${saved.message}`).toBe("saved")
	})
})
