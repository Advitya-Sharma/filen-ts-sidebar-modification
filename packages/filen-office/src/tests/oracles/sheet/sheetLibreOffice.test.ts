import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { CORPUS_DIR } from "../../corpusDir"
import { alignHtml, exportShown, parseCalcHtml, parseCsv, sheetOffices, sheetProfile, unoOffice } from "./libreOffice"
import { buildXlsx } from "./synthetic"
import { UnoSheets, unoSheetByName } from "./unoSheets"
import { readXlsxCells } from "./xlsxCells"

// The LibreOffice spreadsheet oracles: the "as shown" CSV and HTML exports (--convert-to, every install) and the
// UNO runner (hard recalculation, cell dumps, structure operations, xlsx re-save; the first install whose
// bundled Python runs). Opt-in: FILEN_OFFICE_SOFFICE=<soffice>[,<soffice>]; skips without it.

const INSTALLS = sheetOffices()
const UNO = unoOffice(INSTALLS)

// Number formats, a conditional fill (green above 2, ECMA-376 Part 1 §18.3.1.10), a formula error, a boolean,
// a cell fill at B2 (so the data area starts there), a stale cached result (E4) and a second sheet.
const STYLES =
	'<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>' +
	'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
	'<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
	'<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills>' +
	'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
	'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
	'<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
	'<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
	'<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>' +
	'<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>'
const DATA =
	'<sheetData><row r="2"><c r="B2" s="2"/></row>' +
	'<row r="3"><c r="C3" s="1"><v>12.344</v></c><c r="D3" t="inlineStr"><is><t>x</t></is></c><c r="E3" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="F3" t="b"><f>C3&gt;1</f><v>1</v></c></row>' +
	'<row r="4"><c r="C4"><f>C3*2</f><v>24.688</v></c><c r="D4"><v>3</v></c><c r="E4"><f>C4+1</f><v>7</v></c></row>' +
	'<row r="5"><c r="C5"><v>1</v></c><c r="D5"><v>1</v></c></row></sheetData>' +
	'<conditionalFormatting sqref="C3:C5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>'
const WORKBOOK = buildXlsx({
	sheets: [
		{ name: "Data", body: DATA },
		{ name: "Other", body: '<sheetData><row r="1"><c r="A1"><f>Data!C5+1</f><v>0</v></c></row></sheetData>' }
	],
	styles: STYLES
})

describe("the LibreOffice export parsers", () => {
	it("reads CSV fields", () => {
		expect(parseCsv('﻿a,"b,c"\r\n"d""e","f\ng",\n')).toEqual([
			["a", "b,c"],
			['d"e', "f\ng", ""]
		])
	})

	it("reads HTML tables and aligns them with the CSV", () => {
		const html =
			'<h1>Sheet 1: <em>S&amp;1</em></h1><table><tr><td colspan=2 bgcolor="#ff0000">a</td><td rowspan="2">b</td></tr>' +
			'<tr><td>c<br>d</td><td bgcolor="#00FF00"></td></tr></table>'
		const [sheet] = parseCalcHtml(html)

		expect(sheet?.name).toBe("S&1")
		expect(sheet?.cells.map(cell => [cell.row, cell.column, cell.text, cell.fill ?? ""])).toEqual([
			[0, 0, "a", "#FF0000"],
			[0, 2, "b", ""],
			[1, 0, "c\nd", ""],
			[1, 1, "", "#00FF00"]
		])
		expect(alignHtml(sheet as NonNullable<typeof sheet>, [[], ["", "", "a", "", "b"], ["", "", "c d"]])).toEqual({ row: 1, column: 2 })
	})
})

describe.skipIf(INSTALLS.length === 0)("LibreOffice as-shown exports", () => {
	let work: string

	beforeAll(() => {
		work = mkdtempSync(join(tmpdir(), "filen-office-sheet-lo-"))
		writeFileSync(join(work, "synthetic.xlsx"), WORKBOOK)
	})

	afterAll(() => {
		rmSync(work, { recursive: true, force: true })
	})

	for (const office of INSTALLS) {
		it(`exports text as shown and fills (${office.name})`, { timeout: 300_000 }, async () => {
			if (office.unoMissing !== undefined) console.log(`sheet-lo: ${office.name} cannot run the UNO runner: ${office.unoMissing}`)

			const profile = sheetProfile(work, office.name)
			const inputs = [join(work, "synthetic.xlsx")]

			for (const relative of ["xlsx/poi__NumberFormatTests.xlsx", "xlsx/lo__tdf105272.xlsx"])
				if (existsSync(join(CORPUS_DIR, relative))) inputs.push(join(CORPUS_DIR, relative))

			const [synthetic, ...corpus] = await exportShown(office, profile, inputs, join(work, office.name), 120_000)
			const data = synthetic?.sheets.find(sheet => sheet.name === "Data")

			expect(synthetic?.ok).toBe(true)
			expect(synthetic?.sheets.map(sheet => sheet.name).sort()).toEqual(["Data", "Other"])
			// en-US whatever the host's locale: 12.344 in format 0.00.
			expect(data?.text[2]?.slice(2, 6)).toEqual(["12.34", "x", "#DIV/0!", "TRUE"])
			// Cached results are shown as they are: E4 is not recalculated.
			expect(data?.text[3]?.[4]).toBe("7")
			expect(data?.fills?.get("1,1")).toBe("#FF0000")
			// C4 = 24.688 > 2: the rule's fill (both versions apply rule fills; only 26.8 color scales).
			expect(data?.fills?.get("3,2")).toBe("#00FF00")
			expect(data?.fills?.get("4,2")).toBeUndefined()

			for (const workbook of corpus) {
				expect(workbook.ok, workbook.input).toBe(true)
				expect(workbook.sheets.length).toBeGreaterThan(0)
				console.log(
					`sheet-lo ${office.name} ${workbook.input.slice(CORPUS_DIR.length + 1)}: ${workbook.sheets
						.map(
							sheet =>
								`${sheet.name} ${sheet.text.length} rows, ${sheet.fills === undefined ? "fills unaligned" : `${sheet.fills.size} fills`}`
						)
						.join("; ")}`
				)
			}
		})
	}
})

describe.skipIf(UNO === undefined)("the LibreOffice UNO runner", () => {
	let work: string
	let uno: UnoSheets

	beforeAll(() => {
		work = mkdtempSync(join(tmpdir(), "filen-office-sheet-uno-"))
		writeFileSync(join(work, "synthetic.xlsx"), WORKBOOK)
		uno = new UnoSheets(UNO as NonNullable<typeof UNO>, work)
	})

	afterAll(async () => {
		await uno.close()
		rmSync(work, { recursive: true, force: true })
	})

	it("recalculates, dumps, replays operations and saves", { timeout: 300_000 }, async () => {
		const saved = join(work, "saved.xlsx")
		const result = await uno.run({
			input: join(work, "synthetic.xlsx"),
			ops: [
				{ op: "insertRows", sheet: "Data", at: 0, count: 1 },
				{ op: "sort", sheet: "Data", range: "C4:D6", keys: [{ column: 1, ascending: true }] },
				{ op: "moveSheet", sheet: "Other", to: 0 },
				{ op: "deleteColumns", sheet: 1, at: 0, count: 1 },
				{ op: "deleteSheet", sheet: "Nope" }
			],
			saveAs: saved
		})

		expect(result.ok, result.error).toBe(true)
		expect(result.calculated).toBe(true)
		expect(result.ops?.map(op => op.ok)).toEqual([true, true, true, true, false])

		const before = unoSheetByName(result.before ?? { sheets: [], names: [] }, "Data")
		const cell = (sheet: typeof before, row: number, column: number): unknown =>
			sheet?.cells.find(item => item[0] === row && item[1] === column)?.slice(2, 5)

		// Hard recalculation: E4's stale cached 7 becomes C4+1.
		expect(cell(before, 3, 4)).toEqual(["n", 25.688, "25.688"])
		expect(cell(before, 2, 2)).toEqual(["n", 12.344, "12.34"])
		expect(cell(before, 2, 4)).toEqual(["e", "#DIV/0!", "#DIV/0!"])
		expect(cell(before, 2, 5)).toEqual(["b", true, "TRUE"])
		expect(cell(before, 3, 2)).toEqual(["n", 24.688, "24.688"])

		// After: a row on top, rows 4-6 sorted by column D (numbers before text: 1, 3, "x"), column A gone, Other first.
		expect(result.after?.sheets.map(sheet => sheet.name)).toEqual(["Other", "Data"])

		const after = unoSheetByName(result.after ?? { sheets: [], names: [] }, "Data")

		expect(cell(after, 3, 1)).toEqual(["n", 1, "1"])
		expect(cell(after, 3, 2)).toEqual(["n", 1, "1"])
		expect(cell(after, 5, 1)).toEqual(["n", 12.344, "12.34"])
		expect(cell(after, 5, 2)).toEqual(["s", "x", "x"])
		expect((await readXlsxCells(new Uint8Array(readFileSync(saved)))).sheets.map(sheet => sheet.name)).toEqual(["Other", "Data"])
	})

	it("runs over corpus workbooks", { timeout: 600_000 }, async () => {
		for (const relative of ["xlsx/poi__FormatConditionTests.xlsx", "xlsx/govuk__25-26_CPID_list_.xlsx"]) {
			const path = join(CORPUS_DIR, relative)

			if (!existsSync(path)) continue

			const result = await uno.run({
				input: path,
				ops: [{ op: "insertRows", sheet: 0, at: 1, count: 1 }],
				saveAs: join(work, "corpus.xlsx")
			})

			expect(result.ok, result.error).toBe(true)
			expect(result.before?.sheets.length).toBeGreaterThan(0)
			console.log(
				`sheet-uno ${UNO?.name} ${relative}: ${result.cells} cells dumped${result.truncated === true ? " (truncated)" : ""}, calculated ${result.calculated}, ${result.ms} ms`
			)
		}
	})
})
