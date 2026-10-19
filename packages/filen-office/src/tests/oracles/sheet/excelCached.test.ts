import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { CORPUS_DIR } from "../../corpusDir"
import {
	CachedComparison,
	compareCached,
	corpusSource,
	excelCachedValues,
	excelCorpus,
	formulaFunctions,
	manifestSources,
	type ExcelWorkbook
} from "./excelCached"
import { buildXlsx, type SyntheticWorkbook } from "./synthetic"

// Excel's cached formula results as an oracle. The unit tests use synthetic packages; the corpus sweep
// (FILEN_OFFICE_SHEET_CACHED=1) reads every Excel-calculated corpus workbook and checks the harness against
// itself, reporting how many workbooks and formula cells qualify.

const EXCEL = { name: "Microsoft Excel", version: "16.0300" }
const SHEET =
	'<sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f><v>6</v></c><c r="C1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>' +
	'<c r="D1" t="b"><f>A1&gt;1</f><v>1</v></c><c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="F1"><f>NOW()</f><v>45000.5</v></c>' +
	'<c r="G1"><f>[1]Sheet1!A1</f><v>4</v></c></row>' +
	'<row r="2"><c r="B2"><f t="shared" ref="B2:B3" si="0">_xlfn.CONCAT(A2,"x(")</f><v>0</v></c></row>' +
	'<row r="3"><c r="B3"><f t="shared" si="0"/><v>1</v></c><c r="C3"><f>A3</f></c></row></sheetData>'

// calcPr "" leaves calcPr out.
function workbook(overrides: Partial<SyntheticWorkbook> = {}): Uint8Array {
	const { calcPr = 'calcId="191029"', ...rest } = overrides

	return buildXlsx({ sheets: [{ name: "One", body: SHEET }], application: EXCEL, ...(calcPr !== "" && { calcPr }), ...rest })
}

describe("Excel cached values", () => {
	it("reads every cached formula result with its type, functions and flags", async () => {
		const result = await excelCachedValues("synthetic.xlsx", workbook(), "govuk")

		if (typeof result === "string") throw new Error(result)

		const byRef = new Map(result.cells.map(cell => [cell.ref, cell]))

		expect([...byRef.keys()]).toEqual(["B1", "C1", "D1", "E1", "F1", "G1", "B2", "B3"])
		expect(byRef.get("B1")?.expected).toEqual({ type: "n", value: 6 })
		expect(byRef.get("C1")?.expected).toEqual({ type: "s", value: "ab" })
		expect(byRef.get("D1")?.expected).toEqual({ type: "b", value: true })
		expect(byRef.get("E1")?.expected).toEqual({ type: "e", value: "#DIV/0!" })
		expect(byRef.get("F1")?.volatile).toBe(true)
		expect(byRef.get("G1")?.external).toBe(true)
		// A shared formula's dependent takes its master's text.
		expect(byRef.get("B3")?.source).toBe('_xlfn.CONCAT(A2,"x(")')
		expect(byRef.get("B3")?.functions).toEqual(["CONCAT"])
	})

	it("counts only workbooks Excel calculated", async () => {
		expect(await excelCachedValues("x", workbook(), "xlsxwriter")).toBe("source")
		expect(await excelCachedValues("x", workbook({ application: { name: "LibreOffice", version: "1" } }), "govuk")).toBe("application")
		// What XlsxWriter and openpyxl write: results left to be computed on open.
		expect(await excelCachedValues("x", workbook({ calcPr: 'calcId="124519" fullCalcOnLoad="1"' }), "govuk")).toBe("no-calculation")
		expect(await excelCachedValues("x", workbook({ calcPr: "" }), "govuk")).toBe("no-calculation")
		expect(typeof (await excelCachedValues("x", workbook({ calcPr: "", calcChain: true }), "govuk"))).toBe("object")
		expect(await excelCachedValues("x", workbook({ calcPr: 'calcId="191029" calcMode="manual"' }), "govuk")).toBe("manual")
		expect(
			await excelCachedValues(
				"x",
				buildXlsx({
					sheets: [{ name: "S", body: '<sheetData><row r="1"><c r="A1"><f>1+1</f></c></row></sheetData>' }],
					application: EXCEL,
					calcPr: 'calcId="1"'
				}),
				"govuk"
			)
		).toBe("no-cached-values")
		expect(corpusSource("xlsx/lo__x.xlsx", new Map())).toBe("libreoffice")
		expect(corpusSource("xlsx/oxsdk__x.xlsx", new Map())).toBe("openxml-sdk")
		expect(corpusSource("xlsx/a.xlsx", new Map([["xlsx/a.xlsx", "pandas"]]))).toBe("pandas")
	})

	it("finds the functions a formula calls", () => {
		expect(formulaFunctions("IF(A1>0,_xlfn.XLOOKUP(B1,'Sh (1)'!A:A,C:C),\"SUM(\")+Sheet1!A1")).toEqual(["IF", "XLOOKUP"])
		expect(formulaFunctions("_xlfn._xlws.SORT(A1:A9)+t.dist.rt(1,2)")).toEqual(["SORT", "T.DIST.RT"])
	})

	it("compares numbers relatively and the rest exactly", () => {
		expect(compareCached({ type: "n", value: 1e10 }, { type: "n", value: 1e10 + 1 })).toBeUndefined()
		expect(compareCached({ type: "n", value: 1 }, { type: "n", value: 1 + 1e-8 })).toMatch(/Excel 1$/)
		expect(compareCached({ type: "n", value: 0 }, { type: "n", value: 1e-17 })).toBeDefined()
		expect(compareCached({ type: "s", value: "a" }, { type: "s", value: "A" })).toBeDefined()
		expect(compareCached({ type: "b", value: true }, { type: "n", value: 1 })).toBe("type n, Excel b")
		expect(compareCached({ type: "e", value: "#N/A" }, undefined)).toBe("no value")
	})

	it("tallies a recalculation, leaving out volatile, external and excluded cells", async () => {
		const result = (await excelCachedValues("synthetic.xlsx", workbook(), "govuk")) as ExcelWorkbook
		const comparison = new CachedComparison({ exclude: (_, cell) => cell.functions.includes("CONCAT") })

		comparison.add(result, cell => (cell.ref === "B1" ? { type: "n", value: 7 } : cell.expected))

		expect(comparison.compared).toBe(4)
		expect(comparison.matched).toBe(3)
		expect(comparison.skipped).toEqual({ volatile: 1, external: 1, excluded: 2 })
		expect(comparison.byFunction.get("(operators)")).toBe(1)
	})

	// A couple of corpus files when the corpus is there.
	it.skipIf(!existsSync(join(CORPUS_DIR, "xlsx")))("reads corpus workbooks", { timeout: 120_000 }, async () => {
		const sources = manifestSources()

		for (const relative of [
			"xlsx/govuk__25-26_CPID_list_.xlsx",
			"xlsx/govuk__PiPP_Monitoring_October_2026_template.xlsx",
			"xlsx/poi__50154.xlsx"
		]) {
			const path = join(CORPUS_DIR, relative)

			if (!existsSync(path)) continue

			const result = await excelCachedValues(relative, new Uint8Array(readFileSync(path)), corpusSource(relative, sources))

			if (relative.includes("poi__")) expect(result).toBe("source")
			else if (typeof result !== "string") expect(result.cells.length).toBeGreaterThan(0)

			console.log(`excel-cached ${relative}: ${typeof result === "string" ? result : `${result.cells.length} formula cells`}`)
		}
	})
})

describe.skipIf(process.env["FILEN_OFFICE_SHEET_CACHED"] !== "1" || !existsSync(join(CORPUS_DIR, "xlsx")))(
	"Excel cached values over the corpus",
	() => {
		it("selects the Excel-calculated workbooks and matches itself", { timeout: 900_000 }, async () => {
			const comparison = new CachedComparison()
			const rejected: Record<string, number> = {}
			let cells = 0

			for await (const item of excelCorpus()) {
				if ("rejection" in item) {
					rejected[item.rejection] = (rejected[item.rejection] ?? 0) + 1
					continue
				}

				cells += item.cells.length
				comparison.add(item, cell => cell.expected)
			}

			console.log(
				`excel-cached corpus: ${comparison.workbooks} workbooks, ${cells} formula cells with a cached value; compared ${comparison.compared}, skipped ${JSON.stringify(comparison.skipped)}; rejected ${JSON.stringify(rejected)}`
			)
			expect(comparison.workbooks).toBeGreaterThan(0)
			expect(comparison.rate()).toBe(1)
		})
	}
)
