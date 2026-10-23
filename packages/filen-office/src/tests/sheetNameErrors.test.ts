import { describe, expect, it } from "vitest"
import type { SheetSessionApi } from "../core/sheet/protocol"
import { inThreadSheetHost } from "../dom/sheet/host"
import { odsBytes, table } from "./odsFixtures"
import { buildFixture, cells, firstTile, IMMEDIATE, tileTexts } from "./sheetSessionFixtures"

// A #NAME? a formula gives over a name the workbook does not define, or a function Excel does not have, is that
// formula's error, which every formula reading it gives too, as Excel does; a function Excel has but the engine
// lacks is the engine failing, and a formula reading it keeps what it showed.

const XLSX = buildFixture({
	sheets: [
		{
			name: "One",
			body:
				'<dimension ref="A1:B3"/><sheetData>' +
				'<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>SUM(A1:A3)</f><v>5</v></c></row>' +
				'<row r="2"><c r="A2"><v>3</v></c></row>' +
				"</sheetData>"
		}
	]
})

const number = (value: string): string =>
	`<table:table-cell office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell>`
const ODS = odsBytes({
	body: table(
		"One",
		`<table:table-row>${number("2")}<table:table-cell table:formula="of:=SUM([.A1:.A3])" office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell></table:table-row><table:table-row>${number("3")}<table:table-cell/></table:table-row>`,
		'<table:table-column table:number-columns-repeated="2"/>'
	)
})

async function open(kind: "xlsx" | "ods"): Promise<{ session: SheetSessionApi; sheet: number }> {
	const result = await inThreadSheetHost().open(
		(kind === "xlsx" ? XLSX : ODS).slice(),
		{ kind, slicing: IMMEDIATE, settings: { author: "Test" } },
		() => undefined
	)

	if (result.kind !== "opened") throw new Error(result.kind)

	expect((await result.session.intent()).state).toBe("writable")

	return { session: result.session, sheet: result.info.sheets[0]?.id ?? 0 }
}

async function sumAfter(session: SheetSessionApi, sheet: number, input: string): Promise<string | undefined> {
	const result = await session.command({ type: "cells.set", sheet, cells: cells([[2, 0, input]]) })

	expect(result.ok).toBe(true)

	return tileTexts(await firstTile(session, sheet)).get("0,1")
}

describe.each(["xlsx", "ods"] as const)("a #NAME? read by a SUM (%s)", kind => {
	it("is the SUM's result where the name is not defined or the function unknown to Excel", async () => {
		const { session, sheet } = await open(kind)

		expect(await sumAfter(session, sheet, "=Total+1")).toBe("#NAME?")
		expect(await sumAfter(session, sheet, "4")).toBe("9")
		expect(await sumAfter(session, sheet, "=ZORK(1)")).toBe("#NAME?")
		expect(await sumAfter(session, sheet, "")).toBe("5")
		session.close()
	})

	it("leaves the SUM as it showed where the function is Excel's and the engine lacks it", async () => {
		const { session, sheet } = await open(kind)

		expect(await sumAfter(session, sheet, '=CELL("width",A1)')).toBe("5")
		session.close()
	})
})
