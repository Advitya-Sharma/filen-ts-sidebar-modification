import { describe, expect, it } from "vitest"
import type { SheetCommandResult, SheetSessionApi } from "../core/sheet/protocol"
import { firstTile, tileEntry } from "./sheetSessionFixtures"
import { openKeys, sheetBody } from "./sheetKeysFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// SUBTOTAL leaves out the rows a filter hides for every function, and rows hidden by hand too for 101-111
// (support.microsoft.com "SUBTOTAL function"); hiding, showing, filtering and their undo and redo recalculate it.

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason} (${result.refusal.detail ?? ""})`)
}

async function shown(session: SheetSessionApi, row: number, col: number): Promise<string | undefined> {
	return tileEntry(await firstTile(session, 0), row, col)?.text
}

const NUMBERS = sheetBody([
	[1],
	[2],
	[4],
	[8],
	[
		{ f: "SUBTOTAL(109,A1:A4)", v: 15 },
		{ f: "SUBTOTAL(9,A1:A4)", v: 15 }
	]
])

describe("SUBTOTAL over hidden rows", () => {
	it("leaves rows hidden by hand out of 101-111 only, through undo and redo", async () => {
		const { session } = await openKeys({ sheets: [{ name: "S", body: NUMBERS }] })

		await session.intent()
		ok(await session.command({ type: "axis.hide", sheet: 0, axis: "rows", spans: [{ start: 1, end: 1 }], hidden: true }))
		expect([await shown(session, 4, 0), await shown(session, 4, 1)]).toEqual(["13", "15"])

		ok(await session.undo())
		expect(await shown(session, 4, 0)).toBe("15")

		ok(await session.redo())
		expect(await shown(session, 4, 0)).toBe("13")

		ok(await session.command({ type: "axis.hide", sheet: 0, axis: "rows", spans: [{ start: 1, end: 1 }], hidden: false }))
		expect(await shown(session, 4, 0)).toBe("15")
		session.close()
	})

	it("leaves the rows an autofilter hides out of every function", async () => {
		const { session } = await openKeys({ sheets: [{ name: "S", body: NUMBERS }] })
		const range = { startRow: 0, startCol: 0, endRow: 3, endCol: 0 }

		await session.intent()
		ok(await session.command({ type: "filter.toggle", sheet: 0, range, active: { row: 0, col: 0 } }))
		ok(
			await session.command({
				type: "filter.column",
				sheet: 0,
				row: 0,
				col: 0,
				criteria: { kind: "values", values: ["2", "8"], blank: false, dates: [] }
			})
		)
		expect([await shown(session, 4, 0), await shown(session, 4, 1)]).toEqual(["11", "11"])

		ok(await session.undo())
		expect([await shown(session, 4, 0), await shown(session, 4, 1)]).toEqual(["15", "15"])

		ok(await session.redo())
		expect(await shown(session, 4, 1)).toBe("11")

		ok(await session.command({ type: "filter.clear", sheet: 0, row: 0, col: 0 }))
		expect(await shown(session, 4, 1)).toBe("15")
		session.close()
	})

	it("follows a filter put on rows hidden by hand, which shows them", async () => {
		const { session } = await openKeys({ sheets: [{ name: "S", body: NUMBERS }] })

		await session.intent()
		ok(await session.command({ type: "axis.hide", sheet: 0, axis: "rows", spans: [{ start: 2, end: 2 }], hidden: true }))
		expect(await shown(session, 4, 0)).toBe("11")

		ok(
			await session.command({
				type: "filter.toggle",
				sheet: 0,
				range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 },
				active: { row: 0, col: 0 }
			})
		)
		expect(await shown(session, 4, 0)).toBe("15")

		ok(await session.undo())
		expect(await shown(session, 4, 0)).toBe("11")

		ok(await session.redo())
		expect(await shown(session, 4, 0)).toBe("15")
		session.close()
	})

	it("keeps a table's total row on what its filter shows", async () => {
		const rows: (number | string | { f: string; v: number })[][] = [
			["Fruit", "Count"],
			["apple", 3],
			["pear", 7],
			["apple", 1],
			["fig", 5],
			["pear", 9]
		]
		const body = sheetBody([...rows, ["Total", { f: "SUBTOTAL(109,B2:B6)", v: 25 }]])
		const { session } = await openKeys({
			sheets: [
				{
					name: "S",
					body: `${body}<tableParts count="1"><tablePart r:id="rId1"/></tableParts>`,
					rels: relationship("rId1", "table", "../tables/table1.xml")
				}
			],
			parts: {
				"xl/tables/table1.xml": {
					type: TYPES.table,
					text: '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Fruit" displayName="Fruit" ref="A1:B7" totalsRowCount="1"><autoFilter ref="A1:B6"/><tableColumns count="2"><tableColumn id="1" name="Fruit" totalsRowLabel="Total"/><tableColumn id="2" name="Count" totalsRowFunction="sum"/></tableColumns></table>'
				}
			}
		})

		await session.intent()
		ok(
			await session.command({
				type: "filter.column",
				sheet: 0,
				row: 0,
				col: 1,
				criteria: { kind: "dynamic", type: "aboveAverage" }
			})
		)
		expect(await shown(session, 6, 1)).toBe("16")

		ok(await session.undo())
		expect(await shown(session, 6, 1)).toBe("25")
		session.close()
	})
})
