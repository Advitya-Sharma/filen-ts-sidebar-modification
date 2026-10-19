import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { CORPUS_DIR } from "../../corpusDir"
import { formatA1, readXlsxCells, type CellFormula, type CellScalar, type XlsxWorkbook } from "./xlsxCells"

// Excel's cached formula results as an oracle: in a workbook Excel calculated and saved, every <f> cell's <v>
// (with its t) is Excel's own result. A recalculation of ours is compared against those values: relative
// 1e-9 for numbers, exact for text, booleans and errors. Volatile functions are left out, and so is
// anything whose cached value cannot be Excel's.
//
// Which workbooks count. app.xml alone is not evidence: libraries write "Microsoft Excel" too, and leave the
// results to be computed on open (XlsxWriter and openpyxl write calcPr fullCalcOnLoad="1"). A workbook counts
// when all hold:
//   - its corpus source is not a library's or another office suite's test files (EXCLUDED_SOURCES);
//   - app.xml names Microsoft Excel;
//   - Excel's calculation left a trace: a calcChain part, or calcPr with a calcId and without fullCalcOnLoad;
//   - calcMode is not "manual" (results may be stale);
//   - at least one <f> cell has a non-empty cached value.

// Corpus sources whose spreadsheets come from library or other-suite test suites, with cached values that need
// not be Excel's (manifest source ids, and the file name prefixes the corpus uses for them).
export const EXCLUDED_SOURCES: ReadonlySet<string> = new Set(["xlsxwriter", "poi", "closedxml", "libreoffice", "openxml-sdk"])
const PREFIX_SOURCES: Readonly<Record<string, string>> = { lo: "libreoffice", oxsdk: "openxml-sdk" }

// Functions whose result changes on every calculation or with the environment: Excel recalculates them each
// time ("Excel Recalculation", learn.microsoft.com/office/client-developer/excel/excel-recalculation: RAND,
// NOW, TODAY, OFFSET, CELL, INDIRECT, INFO), plus RANDBETWEEN and RANDARRAY, which return new random numbers on
// each calculation (their function pages on support.microsoft.com).
export const VOLATILE: ReadonlySet<string> = new Set([
	"RAND",
	"RANDBETWEEN",
	"RANDARRAY",
	"NOW",
	"TODAY",
	"OFFSET",
	"CELL",
	"INDIRECT",
	"INFO"
])

export interface CachedCell {
	readonly sheet: string
	// A1, and 0-based row and column.
	readonly ref: string
	readonly row: number
	readonly column: number
	readonly formula: CellFormula
	// The text the cell's formula has: its own, or its shared formula master's (unshifted) for a dependent.
	readonly source: string
	readonly expected: CellScalar
	// Upper-case function names called (_xlfn./_xlws. prefixes removed).
	readonly functions: readonly string[]
	readonly volatile: boolean
	// Refers to another workbook ([n] index into externalLinks) or calls a function Excel resolves outside.
	readonly external: boolean
}

export interface ExcelWorkbook {
	readonly name: string
	readonly date1904: boolean
	readonly appVersion: string | undefined
	readonly cells: readonly CachedCell[]
}

export type Rejection = "source" | "application" | "no-calculation" | "manual" | "no-cached-values" | "unreadable"

// Names of the functions a formula calls: string literals and quoted sheet names are skipped first.
export function formulaFunctions(formula: string): string[] {
	const stripped = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')*'/g, "''")
	const out = new Set<string>()

	for (const match of stripped.matchAll(/(?<![\w.$!\]])((?:_xlfn\.|_xlws\.|_xludf\.)*[A-Za-z][\w.]*)\s*\(/g)) {
		out.add((match[1] as string).replace(/^(?:_xlfn\.|_xlws\.|_xludf\.)+/i, "").toUpperCase())
	}

	return [...out]
}

const EXTERNAL = /\[\d+\]/

// The corpus source of a corpus-relative path: the manifest's, else the file name's prefix.
export function corpusSource(relative: string, sources: ReadonlyMap<string, string>): string {
	const listed = sources.get(relative)

	if (listed !== undefined) return listed

	const prefix = relative.slice(relative.lastIndexOf("/") + 1).split("__")[0] ?? ""

	return PREFIX_SOURCES[prefix] ?? prefix
}

// path -> source from the corpus directory's manifest (written by corpus:fetch) or the committed one.
export function manifestSources(): Map<string, string> {
	const candidates = [join(CORPUS_DIR, "manifest.json"), join(import.meta.dirname, "..", "..", "..", "..", "corpus", "manifest.json")]
	const out = new Map<string, string>()

	for (const path of candidates) {
		if (!existsSync(path)) continue

		for (const file of (JSON.parse(readFileSync(path, "utf8")) as { files: { path: string; source: string }[] }).files)
			if (!out.has(file.path)) out.set(file.path, file.source)
	}

	return out
}

// Why `workbook` is not an Excel-calculated workbook, or undefined when it is.
export function rejection(workbook: XlsxWorkbook, source: string): Rejection | undefined {
	if (EXCLUDED_SOURCES.has(source)) return "source"
	if (workbook.application !== "Microsoft Excel") return "application"
	if (!workbook.hasCalcChain && (workbook.calcId === undefined || workbook.fullCalcOnLoad)) return "no-calculation"
	if (workbook.calcMode === "manual") return "manual"

	return undefined
}

// Excel's cached results of every <f> cell of `bytes`, or why the workbook does not count.
export async function excelCachedValues(name: string, bytes: Uint8Array, source: string): Promise<ExcelWorkbook | Rejection> {
	let workbook: XlsxWorkbook

	try {
		workbook = await readXlsxCells(bytes, cell => cell.formula !== undefined)
	} catch {
		return "unreadable"
	}

	const rejected = rejection(workbook, source)

	if (rejected !== undefined) return rejected

	const cells: CachedCell[] = []

	for (const sheet of workbook.sheets) {
		const masters = new Map<string, string>()

		for (const cell of sheet.cells)
			if (cell.formula?.kind === "shared" && cell.formula.si !== undefined && cell.formula.text !== "")
				masters.set(cell.formula.si, cell.formula.text)

		for (const cell of sheet.cells) {
			const formula = cell.formula

			if (formula === undefined || cell.value === undefined) continue

			const source = formula.text !== "" || formula.si === undefined ? formula.text : (masters.get(formula.si) ?? "")
			const functions = formulaFunctions(source)

			cells.push({
				sheet: sheet.name,
				ref: formatA1(cell.row, cell.column),
				row: cell.row,
				column: cell.column,
				formula,
				source,
				expected: cell.value,
				functions,
				volatile: functions.some(name => VOLATILE.has(name)),
				external: EXTERNAL.test(source)
			})
		}
	}

	if (cells.length === 0) return "no-cached-values"

	return { name, date1904: workbook.date1904, appVersion: workbook.appVersion, cells }
}

// The corpus's spreadsheet packages (xlsx, xlsm, xltx, xltm), corpus-relative.
export function corpusSpreadsheets(): string[] {
	const out: string[] = []

	for (const family of ["xlsx", "xlsm", "xltx", "xltm"]) {
		const directory = join(CORPUS_DIR, family)

		if (existsSync(directory)) for (const name of readdirSync(directory).sort()) out.push(`${family}/${name}`)
	}

	return out
}

export interface Rejected {
	readonly name: string
	readonly rejection: Rejection
}

// Every corpus spreadsheet, one at a time (a workbook's cells are dropped once the caller moves on): the
// Excel-calculated ones with their cached results, the others with why they do not count.
export async function* excelCorpus(): AsyncGenerator<ExcelWorkbook | Rejected> {
	const sources = manifestSources()

	for (const relative of corpusSpreadsheets()) {
		const source = corpusSource(relative, sources)
		const result = EXCLUDED_SOURCES.has(source)
			? ("source" as const)
			: await excelCachedValues(relative, new Uint8Array(readFileSync(join(CORPUS_DIR, relative))), source)

		yield typeof result === "string" ? { name: relative, rejection: result } : result
	}
}

export const RELATIVE_TOLERANCE = 1e-9

// Why `actual` does not match Excel's `expected`, or undefined when it does: numbers within `relative` of each
// other (relative to the larger magnitude), text, booleans and errors exactly.
export function compareCached(expected: CellScalar, actual: CellScalar | undefined, relative = RELATIVE_TOLERANCE): string | undefined {
	if (actual === undefined) return "no value"
	if (actual.type !== expected.type) return `type ${actual.type}, Excel ${expected.type}`

	if (expected.type === "n" && actual.type === "n") {
		const difference = Math.abs(actual.value - expected.value)

		return difference <= relative * Math.max(Math.abs(actual.value), Math.abs(expected.value))
			? undefined
			: `${actual.value}, Excel ${expected.value}`
	}

	return actual.value === expected.value ? undefined : `${JSON.stringify(actual.value)}, Excel ${JSON.stringify(expected.value)}`
}

export interface Mismatch {
	readonly workbook: string
	readonly cell: CachedCell
	readonly reason: string
}

export interface CompareOptions {
	// Cells the caller leaves out besides volatile and external ones: unsupported functions, dependents of
	// volatile cells (the caller has the dependency graph), and the like.
	readonly exclude?: (workbook: ExcelWorkbook, cell: CachedCell) => boolean
	readonly relative?: number
}

// Tallies a recalculation against Excel's cached values, workbook by workbook.
export class CachedComparison {
	workbooks = 0
	compared = 0
	matched = 0
	readonly skipped = { volatile: 0, external: 0, excluded: 0 }
	readonly mismatches: Mismatch[] = []
	// Mismatches by function the formula calls (a formula with several counts for each; "(operators)" for none).
	readonly byFunction = new Map<string, number>()
	private readonly options: CompareOptions

	constructor(options: CompareOptions = {}) {
		this.options = options
	}

	// `actual(cell)`: the recalculated value of the cell, undefined when there is none.
	add(workbook: ExcelWorkbook, actual: (cell: CachedCell) => CellScalar | undefined): void {
		this.workbooks++

		for (const cell of workbook.cells) {
			if (cell.volatile) {
				this.skipped.volatile++
				continue
			}

			if (cell.external) {
				this.skipped.external++
				continue
			}

			if (this.options.exclude?.(workbook, cell) === true) {
				this.skipped.excluded++
				continue
			}

			this.compared++

			const reason = compareCached(cell.expected, actual(cell), this.options.relative)

			if (reason === undefined) {
				this.matched++
				continue
			}

			this.mismatches.push({ workbook: workbook.name, cell, reason })

			for (const name of cell.functions.length === 0 ? ["(operators)"] : cell.functions)
				this.byFunction.set(name, (this.byFunction.get(name) ?? 0) + 1)
		}
	}

	// matched / compared.
	rate(): number {
		return this.compared === 0 ? 1 : this.matched / this.compared
	}
}
