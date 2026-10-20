import { describe, expect, it } from "vitest"
import type { SheetChange, SheetCommandResult } from "../core/sheet/protocol"
import { parts } from "./xlsxWriteFixtures"
import { firstTile, openFixture, savedBytes, tileEntry } from "./sheetSessionFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Cells merged and unmerged: the value-loss refusal, what a merge keeps, merge and centre, merge across, and the
// typed refusals.

function changeOf(result: SheetCommandResult): SheetChange {
	if (!result.ok)
		throw new Error(`refused: ${result.refusal.reason}${result.refusal.detail === undefined ? "" : ` (${result.refusal.detail})`}`)

	return result.change
}

const SHEET = {
	sheets: [
		{
			name: "One",
			body:
				'<sheetData><row r="1"><c r="A1" s="1"><v>45000</v></c><c r="B1"><v>2</v></c></row>' +
				'<row r="2"><c r="A2"><f>B1*2</f><v>4</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells>'
		}
	]
}

const range = (startRow: number, startCol: number, endRow: number, endCol: number) => ({ startRow, startCol, endRow, endCol })

describe("sheet session: cells merged", () => {
	it("asks before dropping values, then keeps the upper-left value and gives every cell its format", async () => {
		const { session, info, bytes } = await openFixture(SHEET)
		const sheet = info.sheets[0]!.id
		const asked = await session.command({ type: "merge.cells", sheet, ranges: [range(0, 0, 0, 2)], mode: "merge" })

		expect(asked.ok ? null : asked.refusal.reason).toBe("mergeLosesValues")

		const merged = changeOf(
			await session.command({ type: "merge.cells", sheet, ranges: [range(0, 0, 0, 2)], mode: "merge", discard: true })
		)

		expect([...(merged.layouts[0]?.layout.merges ?? [])]).toEqual([0, 3, 0, 4, 0, 0, 0, 2])

		const sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml).toContain('<c r="A1" s="1"><v>45000</v></c><c r="B1" s="1"/><c r="C1" s="1"/>')
		expect(sheetXml).toContain('<mergeCell ref="A1:C1"/>')
		// A formula reading a dropped value is recalculated.
		expect(tileEntry(await firstTile(session, sheet), 1, 0)?.text).toBe("0")

		changeOf(await session.undo())
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toBe(parts(bytes)["xl/worksheets/sheet1.xml"])
		session.close()
	})

	it("takes in merges the range holds whole, centres it, and merges across row by row", async () => {
		const { session, info } = await openFixture(SHEET)
		const sheet = info.sheets[0]!.id

		changeOf(await session.command({ type: "merge.cells", sheet, ranges: [range(0, 3, 1, 5)], mode: "center" }))

		const saved = parts(await savedBytes(session))
		const sheetXml = saved["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml).toContain('<mergeCells count="1"><mergeCell ref="D1:F2"/></mergeCells>')
		expect(saved["xl/styles.xml"]).toContain('<alignment horizontal="center"/>')

		changeOf(await session.command({ type: "merge.cells", sheet, ranges: [range(4, 0, 6, 2)], mode: "across" }))
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toContain(
			'<mergeCell ref="D1:F2"/><mergeCell ref="A5:C5"/><mergeCell ref="A6:C6"/><mergeCell ref="A7:C7"/>'
		)
		session.close()
	})

	it("refuses to cut through a merge, a table or an array formula, and on a protected sheet", async () => {
		const { session, info } = await openFixture({
			sheets: [
				{
					name: "One",
					body:
						'<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="4"><c r="A4"><f t="array" ref="A4:A5">B4:B5</f><v>0</v></c></row></sheetData>' +
						'<mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells><tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
					rels: relationship("rId1", "table", "../tables/table1.xml")
				}
			],
			parts: {
				"xl/tables/table1.xml": {
					text:
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="H1:I3">' +
						'<autoFilter ref="H1:I3"/><tableColumns count="2"><tableColumn id="1" name="K"/><tableColumn id="2" name="V"/></tableColumns></table>',
					type: TYPES.table
				}
			}
		})
		const sheet = info.sheets[0]!.id
		const reason = async (ranges: ReturnType<typeof range>[]) => {
			const result = await session.command({ type: "merge.cells", sheet, ranges, mode: "merge", discard: true })

			return result.ok ? null : result.refusal
		}

		expect(await reason([range(0, 4, 0, 5)])).toEqual({ reason: "merged" })
		expect(await reason([range(0, 6, 0, 7)])).toEqual({ reason: "merged", detail: "table" })
		expect(await reason([range(3, 0, 3, 1)])).toEqual({ reason: "arrayFormula" })
		session.close()

		const locked = await openFixture({ sheets: [{ name: "One", body: '<sheetData/><sheetProtection sheet="1" formatCells="0"/>' }] })
		const refused = await locked.session.command({
			type: "merge.cells",
			sheet: locked.info.sheets[0]!.id,
			ranges: [range(0, 0, 0, 1)],
			mode: "merge"
		})

		expect(refused.ok ? null : refused.refusal.reason).toBe("protectedSheet")
		locked.session.close()
	})
})

describe("sheet session: cells unmerged", () => {
	it("unmerges every merge the selection touches, and says when there is none", async () => {
		const { session, info, bytes } = await openFixture(SHEET)
		const sheet = info.sheets[0]!.id
		const state = await session.query("commandState", { sheet, ranges: [range(5, 5, 5, 5)], types: ["merge.unmerge"] })

		expect(state.enabled).toEqual([false])

		const change = changeOf(await session.command({ type: "merge.unmerge", sheet, ranges: [range(0, 4, 0, 4)] }))

		expect([...(change.layouts[0]?.layout.merges ?? [])]).toEqual([])
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).not.toContain("mergeCell")

		changeOf(await session.undo())
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toBe(parts(bytes)["xl/worksheets/sheet1.xml"])
		session.close()
	})
})
