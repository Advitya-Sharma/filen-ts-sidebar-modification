import { unzipSync, zipSync } from "fflate"
import { decodeUtf8, encodeUtf8 } from "../../../core/xml/index"
import { buildFixture, SML, type Fixture } from "../../xlsxWriteFixtures"
import { ZIP_ENTRY_TIME } from "../../zipTime"

// Workbooks the sheet fuzz and end-to-end tests start from, built from part text: most of what a session edits
// meets something the file already holds (shared strings with runs and escapes, shared and cross-sheet
// formulas, an array formula, merges, column widths, row heights, styled empty cells, defined names, a hidden
// sheet, a protected sheet).

const STYLES =
	'<numFmts count="2"><numFmt numFmtId="164" formatCode="0.00"/><numFmt numFmtId="165" formatCode="[Red]#,##0;(#,##0)"/></numFmts>' +
	'<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
	'<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
	'<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>' +
	'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
	'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
	'<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
	'<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
	'<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>' +
	'<xf numFmtId="165" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>' +
	'<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
	'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'

const STRINGS = [
	"<si><t>North</t></si>",
	"<si><t>South</t></si>",
	'<si><r><rPr><b/></rPr><t>Rich</t></r><r><t xml:space="preserve"> text</t></r></si>',
	"<si><t>Line_x000D_break</t></si>",
	'<si><t>漢字</t><rPh sb="0" eb="2"><t>カンジ</t></rPh><phoneticPr fontId="0"/></si>',
	"<si><t>Item</t></si>"
]

function column(index: number): string {
	return String.fromCharCode(65 + index)
}

// "Data": 30 rows by 8 columns of mixed content, a shared formula down E, sums in F, an array formula, merges.
function dataSheet(): string {
	let rows = ""

	for (let r = 1; r <= 30; r++) {
		const attributes = r === 4 ? ' ht="24" customHeight="1"' : r === 9 ? ' s="3" customFormat="1"' : ""
		let row = `<row r="${r}"${attributes}>`

		for (let c = 0; c < 8; c++) {
			const ref = `${column(c)}${r}`

			if (c === 0) row += `<c r="${ref}"><v>${r * 3}</v></c>`
			else if (c === 1) row += `<c r="${ref}" s="${r % 3 === 0 ? 2 : 0}"><v>${(r * 1.25).toString()}</v></c>`
			else if (c === 2) row += `<c r="${ref}" t="s"><v>${r % STRINGS.length}</v></c>`
			else if (c === 3 && r % 5 === 0) row += `<c r="${ref}" t="b"><v>${r % 2}</v></c>`
			else if (c === 3 && r === 7) row += `<c r="${ref}" t="e"><f>1/0</f><v>#DIV/0!</v></c>`
			else if (c === 3 && r === 8) row += `<c r="${ref}" t="inlineStr"><is><t>inline</t></is></c>`
			else if (c === 3 && r === 11) row += `<c r="${ref}" s="1"><v>45000</v></c>`
			else if (c === 4 && r === 2)
				row += `<c r="${ref}"><f t="shared" ref="E2:E20" si="0">A2*B2+$A$1</f><v>${r * 3 * r * 1.25 + 3}</v></c>`
			else if (c === 4 && r > 2 && r <= 20) row += `<c r="${ref}"><f t="shared" si="0"/><v>${r * 3 * r * 1.25 + 3}</v></c>`
			else if (c === 5 && r % 4 === 0) row += `<c r="${ref}"><f>SUM(A${r}:B${r})</f><v>${r * 3 + r * 1.25}</v></c>`
			else if (c === 5 && r === 25) row += `<c r="${ref}" t="str"><f>C25&amp;"!"</f><v>${"x"}</v></c>`
			else if (c === 6 && r === 22) row += `<c r="${ref}"><f t="array" ref="G22:G23">A22:A23*2</f><v>132</v></c>`
			else if (c === 6 && r === 23) row += `<c r="${ref}"><v>138</v></c>`
			else if (c === 7 && r === 14) row += `<c r="${ref}" s="3"/>`
		}

		rows += `${row}</row>`
	}

	return (
		'<dimension ref="A1:H30"/><sheetViews><sheetView workbookViewId="0" tabSelected="1"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>' +
		'<sheetFormatPr defaultRowHeight="15"/>' +
		'<cols><col min="1" max="1" width="12.5" customWidth="1"/><col min="3" max="4" width="20" style="4" customWidth="1"/></cols>' +
		`<sheetData>${rows}</sheetData>` +
		'<mergeCells count="2"><mergeCell ref="G1:H2"/><mergeCell ref="A27:C28"/></mergeCells>' +
		'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
	)
}

const CALC =
	'<sheetData><row r="1"><c r="A1"><f>Data!A1+Data!B2</f><v>5.5</v></c><c r="B1"><f>SUM(Data!A1:A10)</f><v>165</v></c>' +
	'<c r="C1"><f>Total*2</f><v>330</v></c><c r="D1"><f>COUNTA(Data!C1:C30)</f><v>30</v></c></row>' +
	'<row r="2"><c r="A2" t="str"><f>IF(A1&gt;3,"big","small")</f><v>big</v></c><c r="B2" t="str"><f>VLOOKUP(6,Data!A1:C30,3,FALSE)</f><v>South</v></c>' +
	'<c r="C2"><f>ROUND(AVERAGE(Data!B1:B30),2)</f><v>19.38</v></c></row>' +
	'<row r="5"><c r="A5" t="s"><v>5</v></c><c r="B5"><v>-7.5</v></c><c r="C5" s="3"><v>1234.5</v></c></row></sheetData>'

const PROTECTED =
	'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" s="4" t="s"><v>0</v></c></row></sheetData>' +
	'<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"/>'

export const FUZZ_FIXTURE: Fixture = {
	sheets: [
		{ name: "Data", body: dataSheet() },
		{ name: "Calc", body: CALC },
		{ name: "Hidden", body: '<sheetData><row r="2"><c r="B2"><v>42</v></c></row></sheetData>' },
		{ name: "Locked", body: PROTECTED }
	],
	sharedStrings: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${SML}" count="31" uniqueCount="${STRINGS.length}">${STRINGS.join("")}</sst>`,
	styles: STYLES,
	workbookTail:
		'<definedNames><definedName name="Total">SUM(Data!$A$1:$A$10)</definedName><definedName name="Local" localSheetId="1">Calc!$A$1:$B$2</definedName></definedNames>' +
		'<calcPr calcId="191029"/>',
	calcChain: true
}

// A one-sheet workbook with no styles part and no shared strings: edits create both.
export const BARE_FIXTURE: Fixture = {
	sheets: [{ name: "Sheet1", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1+1</f><v>2</v></c></row></sheetData>' }],
	styles: null
}

// The rich workbook without defined names: rows and columns move without a name to shift.
export const UNNAMED_FIXTURE: Fixture = {
	...FUZZ_FIXTURE,
	sheets: FUZZ_FIXTURE.sheets.map(sheet =>
		sheet.name === "Calc" ? { ...sheet, body: (sheet.body ?? "").replace("Total*2", "SUM(Data!A1:A10)*2") } : sheet
	),
	workbookTail: '<calcPr calcId="191029"/>'
}

export const FIXTURES: Readonly<Record<string, Fixture>> = { rich: FUZZ_FIXTURE, unnamed: UNNAMED_FIXTURE, bare: BARE_FIXTURE }

// The fixture's package; the "Hidden" sheet's state is set in workbook.xml.
export function fixtureBytes(name: string): Uint8Array {
	const fixture = FIXTURES[name]

	if (fixture === undefined) throw new Error(`no fixture ${name}`)

	const files: Record<string, Uint8Array> = unzipSync(buildFixture(fixture))
	const workbook = files["xl/workbook.xml"]

	if (workbook !== undefined)
		files["xl/workbook.xml"] = encodeUtf8(decodeUtf8(workbook).replace('<sheet name="Hidden"', '<sheet name="Hidden" state="hidden"'))

	return zipSync(files, { mtime: ZIP_ENTRY_TIME })
}
