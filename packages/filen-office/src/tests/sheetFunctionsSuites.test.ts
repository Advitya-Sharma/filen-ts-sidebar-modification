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
	"formulajs/test/financial.js:25":
		'ACCRINT sums Aᵢ/NLᵢ, "NLi = normal length in days of the quasi-coupon period within odd period" (support.microsoft.com "ACCRINT function"); formulajs divides by YEARFRAC\'s year',
	"formulajs/test/financial.js:29":
		'ACCRINT sums Aᵢ/NLᵢ, "NLi = normal length in days of the quasi-coupon period within odd period" (support.microsoft.com "ACCRINT function"); formulajs divides by YEARFRAC\'s year',
	"formulajs/test/financial.js:33":
		'ACCRINT sums Aᵢ/NLᵢ, "NLi = normal length in days of the quasi-coupon period within odd period" (support.microsoft.com "ACCRINT function"); formulajs divides by YEARFRAC\'s year',
	"formulajs/test/financial.js:73":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:77":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:81":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:85":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:89":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:93":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:97":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:101":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:105":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:109":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:113":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:117":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:121":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:125":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:129":
		'ACCRINT: "A value of FALSE (0) returns the accrued interest from first_interest to settlement" (support.microsoft.com "ACCRINT function"); formulajs ignores calc_method',
	"formulajs/test/financial.js:153":
		'ACCRINT sums Aᵢ/NLᵢ, "NLi = normal length in days of the quasi-coupon period within odd period" (support.microsoft.com "ACCRINT function"); formulajs divides by YEARFRAC\'s year',
	"formulajs/test/financial.js:213":
		'ACCRINT sums Aᵢ/NLᵢ, "NLi = normal length in days of the quasi-coupon period within odd period" (support.microsoft.com "ACCRINT function"); formulajs divides by YEARFRAC\'s year',
	"formulajs/test/financial.js:412":
		'"If settlement ≥ maturity, DISC returns the #NUM! error value" (support.microsoft.com "DISC function")',
	"formulajs/test/financial.js:622":
		'"If settlement ≥ maturity, PRICEDISC returns the #NUM! error value" (support.microsoft.com "PRICEDISC function")',
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
	"formulajs/test/text.js:346": "TEXT reads a logical format_text as its text, TRUE, a format of literal letters",
	"formulajs/test/engineering.js:137": SYNTAX,
	"formulajs/test/information.js:8": "the engine has no #NULL! error value: a formula naming it does not parse here (ERROR.TYPE's 1)",
	"formulajs/test/information.js:15":
		"the engine has no #GETTING_DATA error value: a formula naming it does not parse here (ERROR.TYPE's 8)",
	"formulajs/test/math-trig.js:632": SYNTAX,
	"formulajs/test/math-trig.js:634": SYNTAX,
	"formulajs/test/statistical.js:566":
		"INTERCEPT: \"If known_y's and known_x's contain a different number of data points or contain no data points, INTERCEPT returns the #N/A error value\"",
	"formulajs/test/statistical.js:821": RANGES,
	"formulajs/test/statistical.js:1110": RANGES
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
	it("gives the suite's results", { timeout: 120_000 }, () => {
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
const LO_SHEET_POINT =
	"LibreOffice separates a sheet with its own point (Sheet1.A1); Excel's ADDRESS and INDIRECT use ! (support.microsoft.com \"ADDRESS function\": 'EXCEL SHEET'!R2C3)"

const LO_CONVERT_EXACT =
	'CONVERT: each unit is the one its page names ("Grain", "Light-year", "IT calorie", "BTU", "Horsepower", "Pferdestärke", "Admiralty knot", ...) at its exact definition (NIST SP 811 Appendix B); LibreOffice\'s table rounds or approximates the factor'
const LO_CONVERT_CASE =
	'CONVERT: "Unit names and prefixes are case-sensitive.", and the page lists "Picapt" and "Pica" (1/72 inch), "pica" (1/6 inch), and squares and cubes of "Picapt" and "Pica" only: "picapt", and "pica" squared or cubed, are no units'
const LO_CONVERT_FLB = 'CONVERT: "Foot-pound" "flb" is a foot times a pound-force (1.3558 J); LibreOffice\'s flb is the foot-poundal'
const LO_CONVERT_MTON = 'CONVERT: "Measurement ton (freight ton)" "MTON" is 40 cubic feet; LibreOffice\'s is 25'
const LO_CONVERT_HWEIGHT = 'CONVERT: "Imperial hundredweight" is "uk_cwt" or "lcwt" ("hweight"), one unit: 1 exactly'
const LO_PERCENTRANK =
	'PERCENTRANK: significance counts decimals and cuts, not rounds ("If omitted, PERCENTRANK uses three digits (0.xxx)"; the pages give 0.555 for 5/9 and PERCENTRANK.EXC(A2:A10,5.43,1) as 0.3)'
const LO_SIGNIFICANCE = '"If significance < 1, PERCENTRANK returns the #NUM! error value."'
const LO_TRIMMEAN = 'TRIMMEAN: "If percent < 0 or percent > 1, TRIMMEAN returns the #NUM! error value." (1 itself is not out of range)'
const LO_ENCODEURL =
	'ENCODEURL keeps "." (support.microsoft.com "ENCODEURL function": its example gives http%3A%2F%2Fcontoso.sharepoint.com%2FFinance%2F...)'
const LO_ARRAY_ARITHMETIC =
	"the engine evaluates an ordinary formula without array arithmetic, so LINEST's (LOGEST's) array times {9,1} reads its first value; Excel sums both (support.microsoft.com \"LINEST function\", Example 2: 11000)"
const LO_RANDOM = "RANDARRAY's numbers are random: the workbook's expected cell holds a check of them, not a value"

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
	"financial/fods/accrintm.fods Sheet2!A7": LO_SYNTAX,
	"financial/fods/accrintm.fods Sheet2!A12":
		'ACCRINT sums Aᵢ/NLᵢ, "NLi = normal length in days of the quasi-coupon period within odd period" (support.microsoft.com "ACCRINT function"); LibreOffice divides by the days of the issue\'s year',
	"financial/fods/amorlinc.fods Sheet2!A19":
		'AMORLINC takes the bases its page lists, 0, 1, 3 and 4 (support.microsoft.com "AMORLINC function"): actual/360 is #NUM!',
	"financial/fods/amorlinc.fods Sheet2!A20":
		'AMORLINC takes the bases its page lists, 0, 1, 3 and 4 (support.microsoft.com "AMORLINC function"): actual/360 is #NUM!',
	"financial/fods/amorlinc.fods Sheet2!A21":
		'AMORLINC takes the bases its page lists, 0, 1, 3 and 4 (support.microsoft.com "AMORLINC function"): actual/360 is #NUM!',
	"financial/fods/amorlinc.fods Sheet2!A22":
		'AMORLINC takes the bases its page lists, 0, 1, 3 and 4 (support.microsoft.com "AMORLINC function"): actual/360 is #NUM!',
	"financial/fods/amorlinc.fods Sheet2!A23":
		'AMORLINC takes the bases its page lists, 0, 1, 3 and 4 (support.microsoft.com "AMORLINC function"): actual/360 is #NUM!',
	"financial/fods/mduration.fods Sheet2!A2":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/mduration.fods Sheet2!A3":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/mduration.fods Sheet2!A7":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/mduration.fods Sheet2!A8":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/mduration.fods Sheet2!A11":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/mduration.fods Sheet2!A12":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/mduration.fods Sheet2!A16":
		'"If yld < 0 or if coupon < 0, MDURATION returns the #NUM! error value": a yield left empty is 0',
	"financial/fods/mduration.fods Sheet2!A18":
		'"If yld < 0 or if coupon < 0, MDURATION returns the #NUM! error value": a coupon left empty is 0',
	"financial/fods/mduration.fods Sheet2!A19":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/mduration.fods Sheet2!A20":
		'MDURATION discounts from the coupon dates, DSC/E periods to the next: support.microsoft.com "MDURATION function" gives its example as 5.736, LibreOffice\'s YEARFRAC offset 5.734',
	"financial/fods/oddlprice.fods Sheet2!A19":
		'"If rate < 0 or if yld < 0, ODDLPRICE returns the #NUM! error value": a rate of 0, or an argument left empty (0), prices',
	"financial/fods/oddlprice.fods Sheet2!A20":
		'"If rate < 0 or if yld < 0, ODDLPRICE returns the #NUM! error value": a rate of 0, or an argument left empty (0), prices',
	"financial/fods/oddlprice.fods Sheet2!A22":
		'"If rate < 0 or if yld < 0, ODDLPRICE returns the #NUM! error value": a rate of 0, or an argument left empty (0), prices',
	"financial/fods/oddlyield.fods Sheet2!A25":
		'"If rate < 0 or if pr ≤ 0, ODDLYIELD returns the #NUM! error value": a rate of 0, or one left empty (0), yields',
	"financial/fods/oddlyield.fods Sheet2!A26":
		'"If rate < 0 or if pr ≤ 0, ODDLYIELD returns the #NUM! error value": a rate of 0, or one left empty (0), yields',
	"financial/fods/price.fods Sheet2!A17":
		'"If yld < 0 or if rate < 0, PRICE returns the #NUM! error value": an argument left empty is 0, which prices',
	"financial/fods/price.fods Sheet2!A21":
		'"If yld < 0 or if rate < 0, PRICE returns the #NUM! error value": an argument left empty is 0, which prices',
	"financial/fods/price.fods Sheet2!A32":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A33":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A34":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A39":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A40":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A41":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A42":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A43":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A44":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A45":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A46":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A47":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A48":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A49":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A50":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A51":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A52":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A53":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/price.fods Sheet2!A54":
		'PRICE: "When N = 1 (N is the number of coupons payable between the settlement date and redemption date), PRICE is calculated" by its own equation (support.microsoft.com "PRICE function"); LibreOffice discounts as for N > 1',
	"financial/fods/pricemat.fods Sheet2!A10":
		'"If rate < 0 or if yld < 0, PRICEMAT returns the #NUM! error value": an argument left empty is 0, which prices',
	"financial/fods/pricemat.fods Sheet2!A12":
		'"If rate < 0 or if yld < 0, PRICEMAT returns the #NUM! error value": an argument left empty is 0, which prices',
	"financial/fods/received.fods Sheet2!A16":
		"a settlement left empty is 0, the 1900 system's day 0 (1900-01-00, ECMA-376 Part 1 §18.17.4.1): a date, so RECEIVED calculates",
	"financial/fods/yield.fods Sheet2!A14": '"If rate < 0, YIELD returns the #NUM! error value": a rate left empty is 0, which yields',
	"information/fods/type.fods Sheet2!A2": LO_BOOLEANS,
	"information/fods/type.fods Sheet2!A7": LO_SYNTAX,
	"information/fods/type.fods Sheet2!A15": 'TYPE has no 8 in Excel (support.microsoft.com "TYPE function": 1, 2, 4, 16, 64, 128)',
	"spreadsheet/fods/address.fods Sheet2!A2": LO_SHEET_POINT,
	"spreadsheet/fods/address.fods Sheet2!A10": LO_SHEET_POINT,
	"spreadsheet/fods/address.fods Sheet2!A14": LO_SHEET_POINT,
	"spreadsheet/fods/address.fods Sheet2!A25": LO_SHEET_POINT,
	"spreadsheet/fods/address.fods Sheet2!A26": LO_SHEET_POINT,
	"spreadsheet/fods/address.fods Sheet2!A27": LO_SHEET_POINT,
	"spreadsheet/fods/address.fods Sheet2!A28": LO_SHEET_POINT,
	"spreadsheet/fods/hyperlink.fods Sheet2!A30": LO_BOOLEANS,
	"spreadsheet/fods/indirect.fods Sheet2!A12": LO_SHEET_POINT,
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
	"text/fods/text.fods Sheet2!A22": LO_SYNTAX,
	"addin/fods/convert.fods Sheet2!A9": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A16": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A19": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A23": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A26": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A27": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A28": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A29": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A32": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A36": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A43": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A44": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A46": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A47": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A49": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A50": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A51": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A53": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A54": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A57": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A59": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A61": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A77": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A78": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A83": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A105": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A111": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A112": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A120": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A136": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A146": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A153": LO_CONVERT_EXACT,
	"addin/fods/convert.fods Sheet2!A31": LO_CONVERT_FLB,
	"addin/fods/convert.fods Sheet2!A113": LO_CONVERT_FLB,
	"addin/fods/convert.fods Sheet2!A85": LO_CONVERT_MTON,
	"addin/fods/convert.fods Sheet2!A86": LO_CONVERT_MTON,
	"addin/fods/convert.fods Sheet2!A140": LO_CONVERT_MTON,
	"addin/fods/convert.fods Sheet2!A118": LO_CONVERT_HWEIGHT,
	"addin/fods/convert.fods Sheet2!A37":
		'CONVERT: "If the unit does not exist, CONVERT returns the #N/A error value." (the page lists no "pond")',
	"addin/fods/convert.fods Sheet2!A141": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A142": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A143": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A144": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A145": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A147": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A148": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A149": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A150": LO_CONVERT_CASE,
	"addin/fods/convert.fods Sheet2!A151": LO_CONVERT_CASE,
	"array/dubious/fods/linest.fods Sheet2!A17": LO_ARRAY_ARITHMETIC,
	"array/fods/logest.fods Sheet2!A17": LO_ARRAY_ARITHMETIC,
	"spreadsheet/fods/error.type.fods Sheet2!A13": 'A9 is #VALUE! here as in LibreOffice, and "ERROR.TYPE function" numbers #VALUE! 3',
	"spreadsheet/fods/randarray.fods Sheet2!A12": LO_RANDOM,
	"spreadsheet/fods/randarray.fods Sheet2!A26": LO_RANDOM,
	"statistical/fods/intercept.fods Sheet2!A11":
		"INTERCEPT: \"If known_y's and known_x's ... contain no data points, INTERCEPT returns the #N/A error value\"",
	"statistical/fods/percentrank.exc.fods Sheet2!A4": LO_PERCENTRANK,
	"statistical/fods/percentrank.exc.fods Sheet2!A5": LO_PERCENTRANK,
	"statistical/fods/percentrank.exc.fods Sheet2!A12": LO_PERCENTRANK,
	"statistical/fods/percentrank.exc.fods Sheet2!A13": LO_PERCENTRANK,
	"statistical/fods/percentrank.exc.fods Sheet2!A25": LO_PERCENTRANK,
	"statistical/fods/percentrank.exc.fods Sheet2!A21": LO_SIGNIFICANCE,
	"statistical/fods/percentrank.exc.fods Sheet2!A22": LO_SIGNIFICANCE,
	"statistical/fods/percentrank.fods Sheet2!A2": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A3": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A4": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A7": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A9": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A10": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A12": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A14": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A15": LO_PERCENTRANK,
	"statistical/fods/percentrank.fods Sheet2!A21": LO_SIGNIFICANCE,
	"statistical/fods/percentrank.fods Sheet2!A22": LO_SIGNIFICANCE,
	"statistical/fods/percentrank.inc.fods Sheet2!A2": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A3": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A4": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A7": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A9": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A10": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A12": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A14": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A15": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A25": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A27": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A28": LO_PERCENTRANK,
	"statistical/fods/percentrank.inc.fods Sheet2!A21": LO_SIGNIFICANCE,
	"statistical/fods/percentrank.inc.fods Sheet2!A22": LO_SIGNIFICANCE,
	"statistical/fods/permut.fods Sheet2!A15": 'PERMUT: "If number ≤ 0 or if number_chosen < 0, PERMUT returns the #NUM! error value."',
	"statistical/fods/permutationa.fods Sheet2!A11":
		'PERMUTATIONA: "when the total number is zero (0) and the chosen number is larger than zero (0), PERMUTATIONA returns the #NUM! error value"',
	"statistical/fods/trimmean.fods Sheet2!A17": LO_TRIMMEAN,
	"statistical/fods/trimmean.fods Sheet2!A18": LO_TRIMMEAN,
	"text/fods/bahttext.fods Sheet2!A10": LO_SYNTAX,
	"text/fods/encodeurl.fods Sheet2!A7": LO_SYNTAX,
	"text/fods/encodeurl.fods Sheet2!A2": LO_ENCODEURL,
	"text/fods/encodeurl.fods Sheet2!A4": LO_ENCODEURL,
	"text/fods/encodeurl.fods Sheet2!A5": LO_ENCODEURL,
	"text/fods/encodeurl.fods Sheet2!A6": LO_ENCODEURL
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
