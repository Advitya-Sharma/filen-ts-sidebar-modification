import {
	CsvDocument,
	csvDoc,
	engineFormula,
	formulaTranslator,
	malformedFormula,
	parseCellInput,
	parseCsvFile,
	parseTsv,
	renameSheetInFormula,
	serializeCsv,
	shiftFormula,
	shownFormula,
	storedFormula,
	wellFormedFormula
} from "../../core/sheet/index"
import { FormulaEngine } from "../../core/sheet/formula/index"

// The spreadsheet core's text paths (references, CSV reading and writing in every encoding, typed input) and
// a HyperFormula workbook, run alike under V8 and Hermes: each probe's value is compared between them.

function hex(bytes: Uint8Array): string {
	let out = ""

	for (const byte of bytes) out += byte.toString(16).padStart(2, "0")

	return out
}

function bytes(...values: number[]): Uint8Array {
	return Uint8Array.from(values)
}

const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text)

function utf16(text: string, littleEndian: boolean): Uint8Array {
	const out = new Uint8Array(2 + text.length * 2)

	out[0] = littleEndian ? 0xff : 0xfe
	out[1] = littleEndian ? 0xfe : 0xff

	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)

		out[2 + i * 2] = littleEndian ? code & 0xff : code >> 8
		out[3 + i * 2] = littleEndian ? code >> 8 : code & 0xff
	}

	return out
}

const FORMULAS = [
	"SUM(A1:B2)+$C$3*Sheet2!D4",
	"'My Sheet'!A1+[1]Ext!B2+Jan:Mar!C3",
	'VLOOKUP(A:A,1:3,2,FALSE)&"A1"',
	"Table1[[#This Row],[Col]]*LOG10(A1)",
	"IF(A1>0,SUM(A$1:A1),-1",
	"A1;B1",
	"_xlfn.XLOOKUP(A1,B:B,C:C)",
	"1E+3+1E3+TRUE"
]

const CSV_FILES: readonly [string, Uint8Array, boolean][] = [
	["utf-8", utf8('id,name\r\n1,"Größe, ""x"""\r\n2,😀\r\n'), false],
	["utf-8 bom", bytes(0xef, 0xbb, 0xbf, ...utf8("a;b;c\n1;2,5;3\n")), false],
	["utf-8 double bom", bytes(0xef, 0xbb, 0xbf, 0xef, 0xbb, 0xbf, 0x61, 0x0a), false],
	["utf-16le", utf16("a\tb\r\nx€y\t😀\r\n", true), false],
	["utf-16be", utf16("a,b\n1,2", false), false],
	["utf-16le lone and odd", bytes(0xff, 0xfe, 0x61, 0x00, 0x2c, 0x00, 0x00, 0xd8, 0x0a, 0x00, 0x41), false],
	["windows-1252", bytes(0x63, 0x61, 0x66, 0xe9, 0x2c, 0x80, 0x0d, 0x0a, 0x93, 0x71, 0x94, 0x2c, 0x78, 0x0d, 0x0a), false],
	["windows-1252 undefined", bytes(0x63, 0x61, 0x81, 0x2c, 0x62), false],
	["gbk", bytes(0xc3, 0xfb, 0xb3, 0xc6, 0x2c, 0xb3, 0xc7, 0xca, 0xd0, 0x0a), false],
	["invalid utf-8", bytes(0x61, 0x2c, 0xed, 0xa0, 0x80, 0x0a), false],
	["quotes", utf8('ab"c,d"e"f,"g"\n"open\nrest'), false],
	["tsv", utf8('a\t"x\ty"\t"l1\nl2"\n'), true],
	["cr only", utf8("a,b\r1,2\r"), false],
	["empty", bytes(), false]
]

const INPUTS = ["42", "1,234.5", "12%", "-.5e3", "2024-02-29", "2023-02-29", "true", "'007", "=SUM(A1:A3)", "=SUM(A1", " text "]

function csvProbe(file: Uint8Array, tsv: boolean): string {
	const { rows, format } = parseCsvFile(file, tsv)
	const doc = csvDoc(rows, format.writable)
	const cells = [...(doc.sheets[0]?.cells ?? new Map())].map(([key, view]) => `${String(key)}:${view.text}${view.numeric ? "#" : ""}`)
	let saved: string

	try {
		saved = hex(serializeCsv(rows, format))
	} catch (error) {
		saved = `threw ${error instanceof Error ? error.message : String(error)}`
	}

	return JSON.stringify({ rows, format, cells, saved })
}

function csvEdits(): string {
	const { rows, format } = parseCsvFile(utf8("a,b\n1,2\n"), false)
	const document = new CsvDocument(rows, format)
	const results = [
		document.apply({ type: "setCells", sheet: 0, cells: [{ row: 3, col: 2, input: 'new "value"' }] }),
		document.apply({ sheet: 0, type: "insert", axis: "rows", at: 1, count: 1 }),
		document.apply({ sheet: 0, type: "delete", axis: "cols", at: 0, count: 1 }),
		document.undo(),
		document.redo()
	].map(result => result.type)

	return `${results.join()} ${hex(document.serialize().bytes)}`
}

function engineProbe(): string {
	const engine = new FormulaEngine(
		[
			{
				name: "Data",
				rows: [
					[1, 2, "=A1+B1"],
					["=DATE(2024,2,29)", '=TEXT(A2,"yyyy-mm-dd")', "=1/0"]
				],
				formulas: []
			},
			{
				name: "Sum",
				rows: [["=SUM(Data!A1:C1)*Rate", "=NOSUCH(1)", "=VLOOKUP(2,Data!A1:C1,1,FALSE)"]],
				formulas: [{ row: 0, col: 1, content: "=NOSUCH(1)", fallback: 7 }]
			}
		],
		[{ name: "Rate", expression: "2" }]
	)

	try {
		const changed = engine
			.set(0, [{ row: 0, col: 0, content: 10 }])
			.map(cell => `${String(cell.sheet)}:${String(cell.row)},${String(cell.col)}=${String(cell.value)}`)
			.sort()

		engine.insert(0, "rows", 0, 1)

		const after = [
			engine.result(0, 1, 2),
			engine.result(0, 2, 0),
			engine.result(0, 2, 1),
			engine.result(0, 2, 2),
			engine.result(1, 0, 0)
		]

		engine.renameSheet(0, "Renamed")

		return JSON.stringify({
			changed,
			after,
			renamed: engine.result(1, 0, 0),
			fallback: engine.result(1, 0, 1),
			lookup: engine.result(1, 0, 2)
		})
	} finally {
		engine.destroy()
	}
}

export function sheetProbes(probe: (name: string, run: () => unknown) => void): void {
	probe("sheet.refs", () =>
		JSON.stringify(
			FORMULAS.map(formula => [
				wellFormedFormula(formula),
				shownFormula(formula),
				storedFormula(formula),
				engineFormula(formula),
				formulaTranslator(formula)(2, 1),
				shiftFormula(formula, "Sheet1", "Sheet1", { type: "insert", axis: "rows", at: 0, count: 2 }),
				shiftFormula(formula, "Sheet2", "Sheet2", { type: "delete", axis: "cols", at: 2, count: 1 }),
				renameSheetInFormula(formula, "Sheet2", "New 'one'")
			])
		)
	)

	for (const [name, file, tsv] of CSV_FILES) probe(`sheet.csv ${name}`, () => csvProbe(file, tsv))

	probe("sheet.csv edits", csvEdits)
	probe("sheet.tsv clipboard", () => JSON.stringify(parseTsv('a\t"b\nc"\td\r\n"x""y"\t\t')))
	probe("sheet.input", () =>
		JSON.stringify(
			INPUTS.map(input => {
				const parsed = parseCellInput(input)

				return [
					parsed.type === "value" && parsed.value instanceof Date ? { ...parsed, value: parsed.value.toISOString() } : parsed,
					malformedFormula(input)
				]
			})
		)
	)
	probe("sheet.hyperformula", engineProbe)
}
