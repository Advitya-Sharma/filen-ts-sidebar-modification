import { existsSync, readFileSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import type { SheetFileKind } from "../core/sheet/protocol"
import { CORPUS_DIR } from "./corpusDir"
import { RecalcTally, calculates, functionSupport, mismatchClass } from "./fuzz/sheet/recalc"
import { excelCachedValues, excelCorpus, type ExcelWorkbook } from "./oracles/sheet/excelCached"
import { buildXlsx } from "./oracles/sheet/synthetic"

// Our full recalculation against the results Excel cached in the workbooks it calculated: every sheet parsed, the
// formula engine built as an edit builds it, and each formula cell's result compared (relative 1e-9, text,
// booleans and errors exactly) where every function it calls is one the engine calculates, nothing in it is
// volatile, it refers to no other workbook and depends on no volatile cell. Formulas the engine could not take
// fall back to the file's result and are counted apart.
// The corpus run is opt-in: FILEN_OFFICE_SHEET_RECALC=1 (minutes). FILEN_OFFICE_SHEET_RECALC_MIN (default 0.995, the gate) is
// the rate it must reach; FILEN_OFFICE_SHEET_RECALC_OUT=<file.json> keeps the full report.

const EXCEL = { name: "Microsoft Excel", version: "16.0300" }
const RUN = process.env["FILEN_OFFICE_SHEET_RECALC"] === "1"

// Results as Excel stored them; F1 and G2 are deliberately stale (Excel wrote 99), H1 depends on NOW(), and D2
// calls a function no engine has.
const SHEET =
	'<sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f><v>6</v></c><c r="C1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>' +
	'<c r="D1" t="b"><f>A1&gt;1</f><v>1</v></c><c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="F1"><f>A1+1</f><v>99</v></c>' +
	'<c r="G1"><f>NOW()</f><v>45000.5</v></c><c r="H1"><f>G1+1</f><v>45001.5</v></c></row>' +
	'<row r="2"><c r="A2"><v>5</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">A2*2</f><v>10</v></c>' +
	'<c r="D2"><f>_xludf.MYFUNC(A2)</f><v>7</v></c><c r="E2"><f>D2+1</f><v>8</v></c><c r="G2"><f>F1*2</f><v>99</v></c></row>' +
	'<row r="3"><c r="A3"><v>6</v></c><c r="B3"><f t="shared" si="0"/><v>12</v></c></row></sheetData>'

async function synthetic(): Promise<{ workbook: ExcelWorkbook; bytes: Uint8Array }> {
	const bytes = buildXlsx({ sheets: [{ name: "One", body: SHEET }], application: EXCEL, calcPr: 'calcId="191029"' })
	const workbook = await excelCachedValues("synthetic.xlsx", bytes, "govuk")

	if (typeof workbook === "string") throw new Error(workbook)

	return { workbook, bytes }
}

describe("recalculation against Excel's cached results", () => {
	it("names how each function is supported", () => {
		expect(functionSupport("SUM")).toBe("engine")
		expect(functionSupport("CUBEVALUE")).toBe("cached")
		expect(functionSupport("LET")).toBe("missing")
		expect(functionSupport("MYFUNC")).toBe("unknown")
		expect(calculates(["SUM", "IF"])).toBe(true)
		expect(calculates(["SUM", "LET"])).toBe(false)
		expect(mismatchClass({ type: "n", value: 1 }, { type: "s", value: "" })).toBe("n→s")
		expect(mismatchClass({ type: "e", value: "#DIV/0!" }, { type: "e", value: "#REF!" })).toBe("error #DIV/0!→#REF!")
	})

	it("recalculates, compares and classifies a workbook", async () => {
		const { workbook, bytes } = await synthetic()
		const tally = new RecalcTally()
		const result = await tally.add(workbook, bytes, "xlsx")

		expect(result.error).toBeUndefined()
		// B1 C1 D1 E1 F1 B2 B3 G2 compared; G1 volatile (skipped), H1 a dependent of it; D2 unknown; E2 takes D2's
		// stored result and calculates.
		expect(result.unsupported).toBe(1)
		expect(result.volatileDependents).toBe(1)
		expect(tally.comparison.skipped.volatile).toBe(1)
		expect(result.compared).toBe(9)
		// F1 and G2 differ from their stale results; only F1 is a root cause (G2 inherits F1's).
		expect(result.matched).toBe(7)
		expect(tally.comparison.mismatches.map(mismatch => mismatch.cell.ref)).toEqual(["F1", "G2"])
		expect(tally.roots).toBe(1)
		expect(tally.unsupported.get("MYFUNC")).toEqual({ support: "unknown", cells: 1 })
	})
})

describe.skipIf(!RUN || !existsSync(join(CORPUS_DIR, "xlsx")))("recalculation over the Excel-calculated corpus", () => {
	it("matches Excel's cached results", { timeout: 3_600_000 }, async () => {
		const tally = new RecalcTally()
		const rejected: Record<string, number> = {}
		let cells = 0

		for await (const item of excelCorpus()) {
			if ("rejection" in item) {
				rejected[item.rejection] = (rejected[item.rejection] ?? 0) + 1
				continue
			}

			cells += item.cells.length
			await tally.add(item, new Uint8Array(readFileSync(join(CORPUS_DIR, item.name))), item.name.split("/")[0] as SheetFileKind)
		}

		const comparison = tally.comparison
		const failed = tally.workbooks.filter(workbook => workbook.error !== undefined)
		const top = <T>(map: Map<string, T>, value: (item: T) => number, count = 25) =>
			[...map].sort((a, z) => value(z[1]) - value(a[1])).slice(0, count)
		const report = {
			workbooks: comparison.workbooks,
			cells,
			compared: comparison.compared,
			matched: comparison.matched,
			rate: comparison.rate(),
			skipped: comparison.skipped,
			fallbacks: tally.fallbackKinds,
			roots: tally.roots,
			rejected,
			failed: failed.map(workbook => `${workbook.name}: ${workbook.error ?? ""}`),
			mismatchesByFunction: top(comparison.byFunction, count => count),
			rootsByFunction: top(tally.rootFunctions, count => count),
			classes: [...tally.classes].map(([name, kinds]) => [name, [...kinds]]),
			fellBackByFunction: top(tally.fellBack, count => count),
			unsupported: top(tally.unsupported, entry => entry.cells, 60),
			examples: [...tally.examples],
			fallbackExamples: [...tally.fallbackExamples]
		}
		const out = process.env["FILEN_OFFICE_SHEET_RECALC_OUT"]

		if (out !== undefined && out !== "") writeFileSync(out, JSON.stringify(report, null, 1))

		console.log(
			`recalc vs Excel: ${report.workbooks} workbooks, ${cells} formula cells; compared ${report.compared}, matched ${report.matched} ` +
				`(${(report.rate * 100).toFixed(3)} %), root mismatches ${report.roots}; skipped ${JSON.stringify(report.skipped)}; ` +
				`fell back ${JSON.stringify(report.fallbacks)}; failed ${failed.length}\n` +
				`mismatches by function ${JSON.stringify(report.mismatchesByFunction)}\nroots by function ${JSON.stringify(report.rootsByFunction)}`
		)

		expect(failed.map(workbook => workbook.name)).toEqual([])
		expect(comparison.rate()).toBeGreaterThanOrEqual(Number(process.env["FILEN_OFFICE_SHEET_RECALC_MIN"] ?? "0.995"))
	})
})
