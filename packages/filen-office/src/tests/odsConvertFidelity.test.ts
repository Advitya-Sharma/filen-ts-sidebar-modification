import { describe, expect, it } from "vitest"
import { strFromU8, unzipSync } from "fflate"
import { convertSheet, type SheetConversion } from "../core/ods/convert/index"
import { buildWorkbook } from "./xlsxReadFixtures"
import { odsBytes, table } from "./odsFixtures"
import { parseCsvFile } from "../core/sheet/csv/csvView"
import { textWorkbookXlsx } from "../core/xlsx/convert/text"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE } from "./sheetSessionFixtures"

// Conversions between SpreadsheetML and OpenDocument keeping what both can say: carriage returns in text are reported
// where a paragraph cannot hold them, the iteration limit goes by its ODF name, a 1904 workbook's early dates keep
// their day, and a sheet name Excel takes is not renamed.

function converted(result: SheetConversion): Extract<SheetConversion, { ok: true }> {
	if (!result.ok) throw new Error(`${result.reason}: ${result.message}`)

	return result
}

function part(bytes: Uint8Array, name: string): string {
	return strFromU8(unzipSync(bytes)[name] ?? new Uint8Array(0))
}

describe("an xlsx to ods conversion", () => {
	it("reports each cell whose carriage returns a paragraph cannot hold, before a line feed too", async () => {
		const source = buildWorkbook({
			sheets: [
				{
					name: "S",
					body: '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row></sheetData>'
				}
			],
			sharedStrings: ["lone_x000D_cr", "line_x000D_\nnext", "plain\nlines"]
		})
		const result = converted(await convertSheet(source, "xlsx", "ods"))

		expect(result.lossy).toContainEqual({ kind: "other", count: 2, sheets: ["S"] })
		expect(part(result.bytes, "content.xml")).toContain("<text:p>line</text:p><text:p>next</text:p>")
	})

	it("reports the carriage returns of a CSV's quoted fields on Download as .ods, as it reports a workbook's", async () => {
		const rows = parseCsvFile(new TextEncoder().encode('a,"x\r\ny"\r\nb,"lone\rcr"\r\n'), false).rows
		const xlsx = await textWorkbookXlsx(rows, "t")

		if (!xlsx.ok) throw new Error(xlsx.refusal.reason)

		const opened = await inThreadSheetHost().open(xlsx.bytes, { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const result = converted(await opened.session.query("saveAs.format", { format: "ods", language: "en-US" }))

		expect(result.lossy).toContainEqual({ kind: "other", count: 2, sheets: ["t"] })
		expect(part(result.bytes, "content.xml")).toContain("<text:p>x</text:p><text:p>y</text:p>")
		opened.session.close()
	})

	it("writes the iteration limit as table:maximum-difference", async () => {
		const source = buildWorkbook({
			sheets: [{ name: "S", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }],
			workbookTail: '<calcPr iterate="1" iterateCount="50" iterateDelta="0.01"/>'
		})
		const content = part(converted(await convertSheet(source, "xlsx", "ods")).bytes, "content.xml")

		expect(content).toContain('table:maximum-difference="0.01"')
		expect(content).not.toContain("minimum-difference")
	})
})

describe("an ods to xlsx conversion", () => {
	const iteration = (attribute: string): Uint8Array =>
		odsBytes({
			body: `<table:calculation-settings><table:iteration table:status="enable" table:steps="50" ${attribute}="0.01"/></table:calculation-settings>${table("S", '<table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row>')}`
		})

	it("reads the iteration limit from table:maximum-difference, and from the name earlier saves wrote", async () => {
		for (const attribute of ["table:maximum-difference", "table:minimum-difference"]) {
			const workbook = part(converted(await convertSheet(iteration(attribute), "ods", "xlsx")).bytes, "xl/workbook.xml")

			expect(workbook).toContain('iterateDelta="0.01"')
		}
	})

	it("keeps the day of a date before March 1904 in a 1904 date system workbook", async () => {
		const source = odsBytes({
			automatic:
				'<number:date-style style:name="N1"><number:year/><number:text>-</number:text><number:month/><number:text>-</number:text><number:day/></number:date-style><style:style style:name="ce1" style:family="table-cell" style:data-style-name="N1"/>',
			body: `<table:calculation-settings><table:null-date table:date-value="1904-01-01"/></table:calculation-settings>${table(
				"S",
				'<table:table-row><table:table-cell table:style-name="ce1" office:value-type="date" office:date-value="1904-03-01"/><table:table-cell table:style-name="ce1" office:value-type="date" office:date-value="1904-01-02"/></table:table-row>'
			)}`
		})
		const result = converted(await convertSheet(source, "ods", "xlsx"))
		const sheet = part(result.bytes, "xl/worksheets/sheet1.xml")

		expect(part(result.bytes, "xl/workbook.xml")).toContain('date1904="1"')
		expect(sheet).toMatch(/<c r="A1"[^>]*><v>60<\/v>/)
		expect(sheet).toMatch(/<c r="B1"[^>]*><v>1<\/v>/)
		expect(result.lossy.find(feature => feature.kind === "other")).toBeUndefined()
	})

	it("keeps a sheet name that starts or ends with a space", async () => {
		const result = converted(
			await convertSheet(
				odsBytes({
					body: table(
						"2 BUDGET ",
						'<table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row>'
					)
				}),
				"ods",
				"xlsx"
			)
		)

		expect(part(result.bytes, "xl/workbook.xml")).toContain('name="2 BUDGET "')
		expect(result.lossy.find(feature => feature.kind === "sheetName")).toBeUndefined()
	})
})
