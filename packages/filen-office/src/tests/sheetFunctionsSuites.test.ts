import { existsSync, readFileSync } from "node:fs"
import { join, relative } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { pluginFunctionNames } from "../core/sheet/formula/functions"
import { RenderStyleTable } from "../core/sheet/session"
import { XlsxDocument } from "../core/xlsx/edit/document"
import { trackFilteredRows } from "../core/xlsx/edit/filter"
import { openXlsx } from "../core/xlsx/read/index"
import { FormulaEngine, type EngineSheet } from "../core/sheet/formula/index"
import { columnName } from "../core/sheet/refs/cellRef"
import { engineCellFormula } from "../core/sheet/refs/formulaRefs"
import { allFormulajsCases, libreOfficeCases, type CaseValue, type FunctionCase } from "./oracles/sheet/functionCases"
import { suiteDir, suiteFiles, suiteMissing } from "./oracles/sheet/functionSuites"
import { parseA1 } from "./oracles/sheet/xlsxCells"

// The app's own functions against the formulajs suite's literal cases and LibreOffice's function workbooks (fetched
// on request: pnpm exec tsx src/tests/oracles/sheet/functionSuites.ts --fetch formulajs libreoffice; LibreOffice's
// once converted to xlsx by functionSuites.test.ts with FILEN_OFFICE_SOFFICE). A formulajs array argument stands for
// a range: its values go into cells of a data sheet the call refers to. LibreOffice's cases are its converted
// workbooks recalculated here, each formula cell against its sheet's expected column. A case Excel's own
// documentation answers otherwise is listed with the page that decides it; so are LibreOffice's own readings where
// Microsoft is silent and Excel differs.
//
// Licences of the cases' sources: formulajs test/, MIT, Copyright (c) 2014 Sutoiku, Inc.
// (https://github.com/formulajs/formulajs/blob/master/LICENSE); LibreOffice sc/qa/unit/data/functions, MPL-2.0
// (https://www.mozilla.org/MPL/2.0/). Nothing derived from them is committed.

// Where formulajs answers otherwise than Excel's documentation, by case id: the reason.
const RANGES =
	"the array is laid out as a range here, and in a reference these functions leave text out (support.microsoft.com: " +
	'"If an array or reference argument contains text, logical values, or empty cells, those values are ignored")'
const SYNTAX = "Excel takes no such call: the formula does not enter"
const DEVIATIONS: Readonly<Record<string, string>> = {
	"formulajs/test/date-time.js:188": "a text date is read in the engine's day-first order (en-GB); Excel reads it in the system locale's",
	"formulajs/test/information.js:103": "a one-value array is laid out as one cell here, and TYPE of a cell is its value's type",
	"formulajs/test/lookup-reference.js:177": SYNTAX,
	"formulajs/test/lookup-reference.js:179": SYNTAX,
	"formulajs/test/lookup-reference.js:237": SYNTAX,
	"formulajs/test/lookup-reference.js:239": SYNTAX,
	"formulajs/test/lookup-reference.js:3202": SYNTAX,
	"formulajs/test/math-trig.js:64":
		'MODE.SNGL: "If the data set contains no duplicate data points, MODE.SNGL returns the #N/A error value"',
	"formulajs/test/statistical.js:81":
		'AVERAGEIFS: "*" matches text, and no cell here is text: "If no cells meet all the criteria, AVERAGEIFS returns the #DIV/0! error value"',
	"formulajs/test/statistical.js:83": 'AVERAGEIFS: "If no cells meet all the criteria, AVERAGEIFS returns the #DIV/0! error value"',
	"formulajs/test/statistical.js:84": 'AVERAGEIFS: "*" matches text only, so no cell meets both criteria (#DIV/0!)',
	"formulajs/test/statistical.js:241": SYNTAX,
	"formulajs/test/statistical.js:242": SYNTAX,
	"formulajs/test/statistical.js:244": 'COUNTBLANK takes one range (support.microsoft.com "COUNTBLANK function")',
	"formulajs/test/statistical.js:245": 'COUNTBLANK takes one range (support.microsoft.com "COUNTBLANK function")',
	"formulajs/test/statistical.js:247": 'COUNTBLANK takes one range (support.microsoft.com "COUNTBLANK function")',
	"formulajs/test/statistical.js:396": RANGES,
	"formulajs/test/statistical.js:401": RANGES,
	"formulajs/test/statistical.js:571": RANGES,
	"formulajs/test/statistical.js:727": RANGES,
	"formulajs/test/statistical.js:733": RANGES,
	"formulajs/test/text.js:53": 'DOLLAR: "The format used is $#,##0.00_);($#,##0.00)", a negative in parentheses around the sign',
	"formulajs/test/text.js:54": 'DOLLAR: "The format used is $#,##0.00_);($#,##0.00)", a negative in parentheses around the sign',
	"formulajs/test/text.js:55": 'DOLLAR: "The format used is $#,##0.00_);($#,##0.00)", a negative in parentheses around the sign',
	"formulajs/test/text.js:96": "an argument left empty in a formula is 0, not the default (formulajs's undefined)",
	"formulajs/test/text.js:209": "NUMBERVALUE of a logical value: text expected, #VALUE! (formulajs's own #N/A)",
	"formulajs/test/text.js:340": "an error argument is the result, as for every function (formulajs's own #N/A)",
	"formulajs/test/text.js:346": "TEXT reads a logical format_text as its text, TRUE, a format of literal letters"
}

type Cell = string | number | boolean | null

function literal(value: CaseValue): string | null {
	switch (value.type) {
		case "n":
			return String(value.value)
		case "s":
			return `"${value.value.replaceAll('"', '""')}"`
		case "b":
			return value.value ? "TRUE" : "FALSE"
		case "e":
			return value.value
		case "missing":
			return ""
		default:
			return null
	}
}

function cellOf(value: CaseValue): Cell {
	switch (value.type) {
		case "n":
		case "b":
			return value.value
		case "s":
			return `'${value.value}`
		case "e":
			return `=${value.value}`
		default:
			return null
	}
}

// The call as a formula over a data sheet "D": each array (or blank) argument laid out in its own block of cells.
function build(item: FunctionCase): { formula: string; data: Cell[][] } | null {
	if (item.inputs.kind !== "args") return null

	const data: Cell[][] = []
	const args: string[] = []
	let column = 0

	for (const arg of item.inputs.args) {
		if (arg.type === "array" || arg.type === "blank") {
			const rows = arg.type === "blank" ? [[{ type: "blank" } as CaseValue]] : arg.rows
			const width = Math.max(1, ...rows.map(row => row.length))

			rows.forEach((row, r) => {
				const line = (data[r] ??= [])

				row.forEach((value, c) => {
					if (value.type === "array") return

					while (line.length < column + c) line.push(null)

					line[column + c] = cellOf(value)
				})
			})

			const first = `${columnName(column)}1`
			const last = `${columnName(column + width - 1)}${String(rows.length)}`

			args.push(first === last ? `D!${first}` : `D!${first}:${last}`)
			column += width + 1

			continue
		}

		const text = literal(arg)

		if (text === null) return null

		args.push(text)
	}

	return { formula: `${item.function}(${args.join(",")})`, data }
}

function matches(expected: CaseValue, value: unknown, error: boolean, tolerance: number | undefined): boolean {
	if (expected.type === "e") return error && value === expected.value
	if (error) return false
	if (expected.type === "n")
		return (
			typeof value === "number" && Math.abs(value - expected.value) <= (tolerance ?? Math.max(1e-9, Math.abs(expected.value) * 1e-9))
		)
	if (expected.type === "s" || expected.type === "b") return value === expected.value
	if (expected.type === "blank") return value === null || value === 0

	return false
}

function evaluate(item: FunctionCase): { value: unknown; error: boolean } | null {
	const built = build(item)

	if (built === null) return null

	const formula = engineCellFormula(built.formula)

	if (formula === null) return null

	const sheets: EngineSheet[] = [
		{ name: "S", rows: [[formula]], formulas: [] },
		{
			name: "D",
			rows: built.data.map(row =>
				row.map(cell => (typeof cell === "string" && cell.startsWith("=") ? (engineCellFormula(cell.slice(1)) ?? "") : cell))
			),
			formulas: []
		}
	]
	const engine = new FormulaEngine(sheets, [])

	try {
		const result = engine.result(0, 0, 0)

		return { value: result.value, error: result.error === true }
	} finally {
		engine.destroy()
	}
}

describe.skipIf(suiteMissing("formulajs") !== undefined)("the app's functions against the formulajs suite", () => {
	it("gives the suite's results", () => {
		const names = new Set(pluginFunctionNames())
		const cases = allFormulajsCases().cases.filter(item => names.has(item.function) && item.expected.type !== "array")
		const failures: string[] = []
		let compared = 0

		for (const item of cases) {
			const result = evaluate(item)

			if (result === null) continue

			compared++

			if (!matches(item.expected, result.value, result.error, item.tolerance) && DEVIATIONS[item.id] === undefined) {
				const inputs = item.inputs.kind === "args" ? JSON.stringify(item.inputs.args) : ""

				failures.push(
					`${item.id} ${item.function}(${inputs}) expected ${JSON.stringify(item.expected)} got ${JSON.stringify(result)}`
				)
			}
		}

		console.log(`formulajs: ${String(compared)} cases of the app's functions, ${String(failures.length)} failing`)
		expect(failures).toEqual([])
	})
})

// LibreOffice's conversions of its own suite (functionSuites.test.ts writes them; the newest office's).
const LIBREOFFICE_XLSX = join(suiteDir("libreoffice"), "xlsx-lo-26.8.1.1")

const LO_DATES = "LibreOffice's day 0 is 1899-12-30; Excel's 1900 system has 0 as 1900-01-00 (ECMA-376 Part 1 §18.17.4.1)"
const LO_BOOLEANS = "LibreOffice has no logical type (TRUE is 1); Excel's TYPE(TRUE) is 4 and 1 is not TRUE in a lookup"
const LO_SYNTAX = "Excel takes no such call: the formula does not enter"

// Where LibreOffice answers otherwise than Excel, by case id: the reason.
const LIBREOFFICE_DEVIATIONS: Readonly<Record<string, string>> = {
	"array/fods/frequency.fods Sheet2!A61":
		"an ordinary formula reads ABS's array argument at the formula's row (implicit intersection); LibreOffice evaluates it as an array",
	"array/fods/frequency.fods Sheet2!A63":
		"an ordinary formula reads the arguments at the formula's row (implicit intersection); LibreOffice evaluates them as arrays",
	"date_time/fods/day.fods Sheet2!A5": "Excel reads no date before 1900 from text",
	"date_time/fods/month.fods Sheet2!A2": LO_DATES,
	"date_time/fods/month.fods Sheet2!A4": LO_DATES,
	"date_time/fods/month.fods Sheet2!A5": LO_DATES,
	"date_time/fods/month.fods Sheet2!A6": LO_DATES,
	"date_time/fods/month.fods Sheet2!A7": LO_DATES,
	"date_time/fods/year.fods Sheet2!A2": LO_DATES,
	"date_time/fods/year.fods Sheet2!A5": LO_DATES,
	"date_time/fods/year.fods Sheet2!A6": LO_DATES,
	"date_time/fods/year.fods Sheet2!A7": LO_DATES,
	"date_time/fods/year.fods Sheet2!A8": "Excel has no negative date serials: YEAR(-1) is #NUM!",
	"information/fods/type.fods Sheet2!A2": LO_BOOLEANS,
	"information/fods/type.fods Sheet2!A7": LO_SYNTAX,
	"information/fods/type.fods Sheet2!A15": 'TYPE has no 8 in Excel (support.microsoft.com "TYPE function": 1, 2, 4, 16, 64, 128)',
	"spreadsheet/fods/hyperlink.fods Sheet2!A30": LO_BOOLEANS,
	"spreadsheet/fods/xmatch.fods Sheet2!A31":
		"XMATCH with nothing to look up: Microsoft's page is silent; an empty cell matches nothing here",
	"spreadsheet/fods/xmatch.fods Sheet2!A32":
		"XMATCH's next larger item: an empty cell is no larger text here (Microsoft's page is silent)",
	"spreadsheet/fods/xmatch.fods Sheet2!A36":
		'XMATCH search_mode 1 searches from the first item (support.microsoft.com "XMATCH function"): April is item 4',
	"spreadsheet/fods/xmatch.fods Sheet2!A38": LO_BOOLEANS,
	"spreadsheet/fods/xmatch.fods Sheet2!A39": LO_BOOLEANS,
	"spreadsheet/fods/xmatch.fods Sheet2!A40": LO_BOOLEANS,
	"spreadsheet/fods/xmatch.fods Sheet2!A41": LO_BOOLEANS,
	"statistical/fods/averageifs.fods Sheet2!A5":
		"LibreOffice reads the criterion as a regular expression; Excel's criteria take wildcards only",
	"statistical/fods/countblank.fods Sheet2!A3": 'COUNTBLANK takes one range (support.microsoft.com "COUNTBLANK function")',
	"statistical/fods/mode.fods Sheet2!A3": 'MODE of no duplicates is #N/A (support.microsoft.com "MODE function")',
	"statistical/fods/mode.fods Sheet2!A9": 'MODE of no duplicates is #N/A (support.microsoft.com "MODE function")',
	"statistical/fods/mode.fods Sheet2!A10": LO_SYNTAX,
	"statistical/fods/mode.mult.fods Sheet2!A9": 'MODE.MULT of no duplicates is #N/A (support.microsoft.com "MODE.MULT function")',
	"statistical/fods/mode.sngl.fods Sheet2!A9": 'MODE.SNGL of no duplicates is #N/A (support.microsoft.com "MODE.SNGL function")',
	"statistical/fods/rank.fods Sheet2!A10":
		"the ranked cell is INT of a time difference times 86400, which LibreOffice's floating point gives one less",
	"text/fods/dollar.fods Sheet2!A6": 'DOLLAR: "The format used is $#,##0.00_);($#,##0.00)", a negative in parentheses',
	"text/fods/dollar.fods Sheet2!A7": 'DOLLAR: "The format used is $#,##0.00_);($#,##0.00)", a negative in parentheses',
	"text/fods/fixed.fods Sheet2!A4": 'the engine reads no grouping comma in text as a number ("1,234,567.890"); Excel en-US does',
	"text/fods/fixed.fods Sheet2!A6": 'the engine reads no grouping comma in text as a number ("1,234,567.890"); Excel en-US does',
	"text/fods/numbervalue.fods Sheet2!A6": 'NUMBERVALUE("3.5%") is 0.035 (support.microsoft.com "NUMBERVALUE function", its example)',
	"text/fods/text.fods Sheet2!A4": LO_DATES,
	"text/fods/text.fods Sheet2!A5": LO_DATES,
	"text/fods/text.fods Sheet2!A6": LO_DATES,
	"text/fods/text.fods Sheet2!A22": LO_SYNTAX
}

// LibreOffice's expected text for an error: an Excel error text, "error" for any, "Err:nnn" for its own codes.
// Null for a case left out: LibreOffice's own error codes, and arrays (array formulas keep their stored results here).
function matchesLibreOffice(expected: CaseValue, value: unknown, error: boolean): boolean | null {
	if (expected.type === "array") return null
	if (expected.type === "s" && expected.value.startsWith("Err:")) return null
	if (expected.type === "s" && expected.value === "error") return error
	if (expected.type === "s" && /^#[A-Z0-9/!?]+$/.test(expected.value)) return error && value === expected.value

	return matches(expected, value, error, undefined)
}

describe.skipIf(suiteMissing("libreoffice") !== undefined || !existsSync(LIBREOFFICE_XLSX))(
	"the app's functions against LibreOffice's function workbooks",
	() => {
		it("gives the workbooks' expected results", { timeout: 600_000 }, async () => {
			const names = new Set(pluginFunctionNames())
			const files = join(suiteDir("libreoffice"), "files")
			const failures: string[] = []
			let compared = 0

			for (const source of suiteFiles("libreoffice")) {
				const path = relative(files, source)
				const xlsx = join(LIBREOFFICE_XLSX, `${path}.xlsx`)
				const stem = (path.split("/").at(-1) ?? "").replace(/\.fods$/, "").toUpperCase()

				if (!names.has(stem) || !existsSync(xlsx)) continue

				const cases = (await libreOfficeCases(path, xlsx)).cases.filter(item => names.has(item.function))

				if (cases.length === 0) continue

				const opened = await openXlsx(bytesSource(new Uint8Array(readFileSync(xlsx))))

				if (!opened.ok) throw new Error(`${path}: ${opened.reason}`)

				const document = new XlsxDocument(opened.workbook, new RenderStyleTable(), { kind: "xlsx" })

				try {
					await document.readyAllSheets()
					trackFilteredRows(document)

					const engine = document.buildEngine()
					const sheets = new Map(document.engineSheets.map((id, index) => [document.entry(id)?.name ?? "", index]))

					for (const item of cases) {
						if (item.inputs.kind !== "cell") continue

						const at = parseA1(item.inputs.cell)
						const sheet = sheets.get(item.inputs.sheet)

						if (at === undefined || sheet === undefined) continue

						const result = engine.result(sheet, at.row, at.column)
						const ok = matchesLibreOffice(item.expected, result.value, result.error === true)

						if (ok === null) continue

						compared++

						if (!ok && LIBREOFFICE_DEVIATIONS[item.id] === undefined)
							failures.push(
								`${item.id} ${item.inputs.formula} expected ${JSON.stringify(item.expected)} got ${JSON.stringify(result)}`
							)
					}
				} finally {
					document.close()
				}
			}

			console.log(`LibreOffice: ${String(compared)} cases of the app's functions, ${String(failures.length)} failing`)
			expect(compared).toBeGreaterThan(1000)
			expect(failures).toEqual([])
		})
	}
)
