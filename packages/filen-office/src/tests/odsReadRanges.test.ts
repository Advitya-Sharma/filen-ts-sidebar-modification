import { describe, expect, it } from "vitest"
import { strFromU8, unzipSync } from "fflate"
import { bytesSource } from "../core/opc/index"
import { convertSheet } from "../core/ods/convert/index"
import { openOds } from "../core/ods/open"
import { inThreadSheetHost } from "../dom/sheet/host"
import { odsBytes, table } from "./odsFixtures"
import { IMMEDIATE } from "./sheetSessionFixtures"

// Cell range addresses an .ods gives (ODF 1.3 Part 3 §9.2.5): a list of ranges whose sheet names may hold the
// characters that separate them, and a spreadsheet of no table at all, which reads as one empty sheet.

const SHEET = "LP-Inv, JVs & Assoc"
const QUOTED = `'${SHEET.replaceAll("&", "&amp;")}'`
const ROWS = [1, 7, 3]
	.map(
		value =>
			`<table:table-row><table:table-cell office:value-type="float" office:value="${String(value)}"><text:p>${String(value)}</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="${String(value)}"><text:p>${String(value)}</text:p></table:table-cell></table:table-row>`
	)
	.join("")
const STYLES =
	'<style:style style:name="Default" style:family="table-cell"/><style:style style:name="Bad" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style>'

function conditional(target: string): string {
	return `<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="${target}"><calcext:condition calcext:apply-style-name="Bad" calcext:value="&gt;5" calcext:base-cell-address="${QUOTED}.A1"/></calcext:conditional-format></calcext:conditional-formats>`
}

async function sections(bytes: Uint8Array): Promise<string> {
	const opened = await openOds(bytesSource(bytes))

	if (!opened.ok) throw new Error(`${opened.reason}: ${opened.message}`)

	const sheet = opened.workbook.sheets[0]?.model

	return (sheet?.sections ?? []).map(section => (section.bytes === null ? "" : strFromU8(section.bytes))).join("")
}

describe("an .ods range list", () => {
	it("keeps a conditional format on its ranges where the sheet's name holds a comma and a space", async () => {
		const bytes = odsBytes({
			styles: STYLES,
			body: table(SHEET.replaceAll("&", "&amp;"), ROWS).replace(
				"</table:table>",
				`${conditional(`${QUOTED}.A1:${QUOTED}.A3 ${QUOTED}.B2:${QUOTED}.B3`)}</table:table>`
			)
		})

		expect(await sections(bytes)).toContain('sqref="A1:A3 B2:B3"')

		const result = await convertSheet(bytes, "ods", "xlsx")

		if (!result.ok) throw new Error(`${result.reason}: ${result.message}`)

		const sheet = strFromU8(unzipSync(result.bytes)["xl/worksheets/sheet1.xml"] ?? new Uint8Array(0))

		expect(sheet).toContain('sqref="A1:A3 B2:B3"')
	})

	it("gives a print range on such a sheet as its print area", async () => {
		const bytes = odsBytes({
			body: table(
				SHEET.replaceAll("&", "&amp;"),
				ROWS,
				"",
				` table:print-ranges="${QUOTED}.A1:${QUOTED}.B2 ${QUOTED}.A3:${QUOTED}.B3"`
			)
		})
		const opened = await openOds(bytesSource(bytes))

		if (!opened.ok) throw new Error(opened.message)

		const area = opened.workbook.definedNames.find(name => name.name === "_xlnm.Print_Area")

		expect(area?.formula).toBe(`'${SHEET}'!$A$1:$B$2,'${SHEET}'!$A$3:$B$3`)
	})
})

describe("an .ods spreadsheet without a table", () => {
	it("opens as one empty sheet, as LibreOffice opens it", async () => {
		const opened = await openOds(bytesSource(odsBytes({ body: "" })))

		if (!opened.ok) throw new Error(`${opened.reason}: ${opened.message}`)

		expect(opened.workbook.sheets.map(sheet => sheet.kind)).toEqual(["worksheet"])
		expect(opened.workbook.sheets[0]?.model?.store.cellCount ?? 0).toBe(0)
	})

	it("saves untouched as its bytes, and an edit as a table holding it", async () => {
		const bytes = odsBytes({ body: "" })
		const opened = await inThreadSheetHost().open(bytes.slice(), { kind: "ods", slicing: IMMEDIATE }, () => undefined)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const sheet = opened.info.sheets[0]?.id ?? 0

		expect(opened.info.sheets.map(entry => entry.name)).toEqual(["Sheet1"])
		expect(await opened.session.intent()).toEqual({ state: "writable" })

		const untouched = await opened.session.save()

		expect(untouched.ok && untouched.bytes).toEqual(bytes)
		expect(
			(
				await opened.session.command({
					type: "cells.set",
					sheet,
					cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(0), inputs: ["7"] }
				})
			).ok
		).toBe(true)

		const saved = await opened.session.save()

		if (!saved.ok) throw new Error(saved.message)

		const back = await openOds(bytesSource(saved.bytes))

		if (!back.ok) throw new Error(back.message)

		expect(back.workbook.sheets.map(entry => entry.name)).toEqual(["Sheet1"])
		expect(strFromU8(unzipSync(saved.bytes)["content.xml"] ?? new Uint8Array(0))).toContain('office:value="7"')
		opened.session.close()
	})

	it("saves the new table with its validations before the elements that follow the tables", async () => {
		const opened = await inThreadSheetHost().open(
			odsBytes({ body: "<table:named-expressions/>" }),
			{ kind: "ods", slicing: IMMEDIATE },
			() => undefined
		)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const sheet = opened.info.sheets[0]?.id ?? 0

		await opened.session.intent()

		for (const command of [
			{ type: "cells.set", sheet, cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(0), inputs: ["a"] } },
			{
				type: "validation.set",
				sheet,
				ranges: [{ startRow: 1, startCol: 0, endRow: 1, endCol: 0 }],
				rule: {
					type: "list",
					operator: "between",
					value1: "a,b",
					value2: "",
					allowBlank: true,
					inCellDropdown: true,
					showInputMessage: false,
					promptTitle: "",
					prompt: "",
					showErrorMessage: true,
					errorTitle: "",
					error: "",
					errorStyle: "stop"
				}
			}
		] as const)
			expect((await opened.session.command(command)).ok).toBe(true)

		const saved = await opened.session.save()

		if (!saved.ok) throw new Error(saved.message)

		const xml = strFromU8(unzipSync(saved.bytes)["content.xml"] ?? new Uint8Array(0))

		expect(xml.indexOf("<table:content-validations>")).toBeGreaterThan(0)
		expect(xml.indexOf("<table:content-validations>")).toBeLessThan(xml.indexOf("<table:table "))
		expect(xml.indexOf("<table:table ")).toBeLessThan(xml.indexOf("<table:named-expressions"))
		opened.session.close()
	})

	it("converts to a workbook of one empty sheet", async () => {
		const result = await convertSheet(odsBytes({ body: "" }), "ods", "xlsx")

		expect(result.ok).toBe(true)
	})
})
