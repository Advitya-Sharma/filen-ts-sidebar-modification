import { HyperFormula, type SimpleCellAddress, type SimpleCellRange } from "hyperformula"
import { bytesSource } from "../../../core/opc/index"
import type { FormulaEngine } from "../../../core/sheet/formula/formulaEngine"
import { registerSheetFunctions } from "../../../core/sheet/formula/functions"
import { sheetFunction } from "../../../core/sheet/functions/catalogue"
import type { SheetFileKind } from "../../../core/sheet/protocol"
import { RenderStyleTable } from "../../../core/sheet/session"
import { XlsxDocument } from "../../../core/xlsx/edit/document"
import { openXlsx } from "../../../core/xlsx/read/index"
import { CachedComparison, type CachedCell, type ExcelWorkbook } from "../../oracles/sheet/excelCached"
import type { CellScalar } from "../../oracles/sheet/xlsxCells"

// Our full recalculation against Excel's cached results: a workbook is opened by the engine's own reader, every
// sheet parsed, and the formula engine built as an edit builds it (excelCached.ts selects the workbooks and
// cells). A cell is compared when every function it calls is one the engine calculates, it calls nothing
// volatile, refers to no other workbook, and depends on no volatile cell (Excel's result there is of its own
// moment). Cells the engine could not calculate at all (it fell back to the file's result) are counted apart:
// they would otherwise match by construction.

registerSheetFunctions()

// What the engine registers (the app's own functions included), upper case: what "calculates" means here, whatever
// the catalogue says.
const REGISTERED: ReadonlySet<string> = new Set(HyperFormula.getRegisteredFunctionNames("enGB").map(name => name.toUpperCase()))

export type Support = "engine" | "cached" | "missing" | "unknown"

// How a function name is supported: by the engine, cached by design (network, add-ins, OLAP; the catalogue's
// "cached"), missing from the engine (to be added), or not one of Excel's functions at all (a user's or an
// add-in's name).
export function functionSupport(name: string): Support {
	if (REGISTERED.has(name)) return "engine"

	const listed = sheetFunction(name)

	if (listed === undefined) return "unknown"

	return listed.support === "cached" ? "cached" : "missing"
}

export function calculates(functions: readonly string[]): boolean {
	return functions.every(name => functionSupport(name) === "engine")
}

// HyperFormula's instance inside our engine: only the oracle reads it (formula presence and the dependency graph).
function hyperFormula(engine: FormulaEngine): HyperFormula {
	return (engine as unknown as { readonly engine: HyperFormula }).engine
}

function isRange(node: SimpleCellAddress | SimpleCellRange): node is SimpleCellRange {
	return "start" in node
}

function nodeKey(node: SimpleCellAddress | SimpleCellRange): string {
	return isRange(node)
		? `${node.start.sheet}:${node.start.row},${node.start.col}:${node.end.row},${node.end.col}`
		: `${node.sheet}:${node.row},${node.col}`
}

// Every cell that depends on one of `roots`, through any chain (bounded: a cell graph larger than `limit` stops
// the walk and reports it).
function dependentsOf(hf: HyperFormula, roots: readonly SimpleCellAddress[], limit: number): { cells: Set<string>; truncated: boolean } {
	const seen = new Set<string>()
	const cells = new Set<string>()
	const queue: (SimpleCellAddress | SimpleCellRange)[] = [...roots]

	for (const root of roots) seen.add(nodeKey(root))

	while (queue.length > 0) {
		if (seen.size > limit) return { cells, truncated: true }

		const node = queue.pop() as SimpleCellAddress | SimpleCellRange

		for (const dependent of hf.getCellDependents(node)) {
			const key = nodeKey(dependent)

			if (seen.has(key)) continue

			seen.add(key)
			if (!isRange(dependent)) cells.add(key)
			queue.push(dependent)
		}
	}

	return { cells, truncated: false }
}

// The engine's result as the scalar a file would store.
function scalarOf(result: { readonly value: number | string | boolean | null; readonly error?: true }): CellScalar | undefined {
	const value = result.value

	if (value === null) return undefined
	if (result.error === true) return { type: "e", value: String(value) }
	if (typeof value === "number") return { type: "n", value }
	if (typeof value === "boolean") return { type: "b", value }

	return { type: "s", value }
}

// Why a mismatch happened, for the report: what kind of value differs.
export function mismatchClass(expected: CellScalar, actual: CellScalar | undefined): string {
	if (actual === undefined) return "empty result"
	if (actual.type !== expected.type) return `${expected.type}→${actual.type}`
	if (expected.type === "n" && actual.type === "n") {
		const relative = Math.abs(actual.value - expected.value) / Math.max(Math.abs(expected.value), Math.abs(actual.value))

		return relative < 1e-6 ? "number (rounding)" : "number"
	}
	if (expected.type === "e") return `error ${String(expected.value)}→${String(actual.value)}`

	return expected.type === "s" ? "text" : "boolean"
}

export interface RecalcExample {
	readonly workbook: string
	readonly sheet: string
	readonly ref: string
	readonly source: string
	readonly expected: CellScalar
	readonly actual: CellScalar | undefined
}

export interface RecalcWorkbook {
	readonly name: string
	readonly cells: number
	readonly compared: number
	readonly matched: number
	readonly fellBack: number
	readonly volatileDependents: number
	readonly unsupported: number
	readonly error?: string
	readonly buildMs: number
}

// The tally over many workbooks: the plain comparison, plus why cells were left out.
// Whether the engine fell back to the file's result for a cell: it was given a value, not a formula, or the formula
// only holds the file's result (FILEN.HELD...: the session keeps a result it cannot calculate in place). Either
// matches Excel by construction, so neither counts as calculated.
export function fellBack(hf: HyperFormula, address: SimpleCellAddress): boolean {
	return !hf.doesCellHaveFormula(address) || /^=\s*FILEN\.HELD/i.test(hf.getCellFormula(address) ?? "")
}

export class RecalcTally {
	readonly comparison: CachedComparison
	readonly workbooks: RecalcWorkbook[] = []
	// Left out because a function is not calculated, by the function and why.
	readonly unsupported = new Map<string, { support: Support; cells: number }>()
	// Cells whose functions all calculate, yet the engine fell back to the file's result: by function.
	readonly fellBack = new Map<string, number>()
	// Mismatches by function and by kind of difference.
	readonly classes = new Map<string, Map<string, number>>()
	// A few mismatches of each function and kind, with both values.
	readonly examples = new Map<string, RecalcExample[]>()
	// Cells the engine was given a value for, not a formula: legacy array formulas and data tables keep their
	// results by design ("kept"), the rest failed to read or calculate in the engine ("failed").
	readonly fallbackKinds = { kept: 0, failed: 0 }
	// A few formulas the engine fell back on, by workbook.
	readonly fallbackExamples = new Map<string, string[]>()
	// Mismatches none of whose precedents mismatch (root causes), and by function.
	roots = 0
	readonly rootFunctions = new Map<string, number>()
	private current = new Set<string>()
	private fallbacks = new Set<string>()
	private counts = { unsupported: 0, fellBack: 0, dependents: 0 }
	// Root-cause examples kept per function and kind.
	private readonly exampleLimit: number

	constructor(exampleLimit = 4) {
		this.exampleLimit = exampleLimit

		this.comparison = new CachedComparison({
			exclude: (workbook, cell) => this.excluded(workbook, cell)
		})
	}

	private key(workbook: ExcelWorkbook, cell: CachedCell): string {
		return `${workbook.name}|${cell.sheet}|${cell.row},${cell.column}`
	}

	private excluded(workbook: ExcelWorkbook, cell: CachedCell): boolean {
		if (!calculates(cell.functions)) {
			for (const name of cell.functions) {
				const support = functionSupport(name)

				if (support === "engine") continue

				const entry = this.unsupported.get(name) ?? { support, cells: 0 }

				entry.cells++
				this.unsupported.set(name, entry)
			}

			this.counts.unsupported++

			return true
		}

		const key = this.key(workbook, cell)

		if (this.fallbacks.has(key)) {
			if (cell.formula.kind === "array" || cell.formula.kind === "dataTable") this.fallbackKinds.kept++
			else {
				this.fallbackKinds.failed++

				const examples = this.fallbackExamples.get(workbook.name) ?? []

				if (examples.length < 3) examples.push(`${cell.sheet}!${cell.ref} =${cell.source.slice(0, 160)}`)

				this.fallbackExamples.set(workbook.name, examples)
			}

			for (const name of cell.functions.length === 0 ? ["(operators)"] : cell.functions)
				this.fellBack.set(name, (this.fellBack.get(name) ?? 0) + 1)

			this.counts.fellBack++

			return true
		}

		if (!this.current.has(key)) return false

		this.counts.dependents++

		return true
	}

	// Recalculates one Excel-calculated workbook and compares its cells.
	async add(workbook: ExcelWorkbook, bytes: Uint8Array, kind: SheetFileKind): Promise<RecalcWorkbook> {
		const started = performance.now()
		const before = {
			compared: this.comparison.compared,
			matched: this.comparison.matched,
			mismatches: this.comparison.mismatches.length
		}

		this.counts = { unsupported: 0, fellBack: 0, dependents: 0 }

		const opened = await openXlsx(bytesSource(bytes))

		if (!opened.ok) return this.failed(workbook, `open ${opened.reason}`, started)

		const document = new XlsxDocument(opened.workbook, new RenderStyleTable(), { kind })

		try {
			await document.readyAllSheets()

			const engine = document.buildEngine()
			const hf = hyperFormula(engine)
			const indexByName = new Map<string, number>()

			document.engineSheets.forEach((sheet, index) => {
				const name = document.entry(sheet)?.name

				if (name !== undefined) indexByName.set(name, index)
			})

			const hfSheet = (index: number): number | undefined =>
				hf.getSheetId(document.entry(document.engineSheets[index] ?? -1)?.name ?? "")
			const volatile: SimpleCellAddress[] = []

			this.fallbacks = new Set()

			for (const cell of workbook.cells) {
				const index = indexByName.get(cell.sheet)
				const sheet = index === undefined ? undefined : hfSheet(index)

				if (sheet === undefined) continue

				const address = { sheet, row: cell.row, col: cell.column }

				if (cell.volatile) volatile.push(address)
				else if (fellBack(hf, address)) this.fallbacks.add(this.key(workbook, cell))
			}

			// Volatile functions the engine has (NOW, RAND, OFFSET...) taint what depends on them.
			const tainted = dependentsOf(hf, volatile, 2_000_000)
			const names = new Map<number, string>()

			for (const [name, index] of indexByName) {
				const sheet = hfSheet(index)

				if (sheet !== undefined) names.set(sheet, name)
			}

			this.current = new Set()

			for (const key of tainted.cells) {
				const [sheetText = "", at = ""] = key.split(":")
				const name = names.get(Number(sheetText))

				if (name !== undefined) this.current.add(`${workbook.name}|${name}|${at}`)
			}

			this.comparison.add(workbook, cell => {
				const index = indexByName.get(cell.sheet)

				return index === undefined ? undefined : scalarOf(engine.result(index, cell.row, cell.column))
			})

			const mismatches = this.comparison.mismatches.slice(before.mismatches)
			// Mismatched cells by engine sheet: a mismatch none of whose precedents mismatch is a root cause; the
			// rest inherit theirs.
			const wrong = new Map<number, { row: number; col: number }[]>()
			const wrongKeys = new Set<string>()

			for (const mismatch of mismatches) {
				const sheet = hfSheet(indexByName.get(mismatch.cell.sheet) ?? -1)

				if (sheet === undefined) continue

				const list = wrong.get(sheet) ?? []

				list.push({ row: mismatch.cell.row, col: mismatch.cell.column })
				wrong.set(sheet, list)
				wrongKeys.add(`${sheet}:${mismatch.cell.row},${mismatch.cell.column}`)
			}

			const inherits = (sheet: number, row: number, col: number): boolean =>
				hf
					.getCellPrecedents({ sheet, row, col })
					.some(node =>
						isRange(node)
							? (wrong.get(node.start.sheet) ?? []).some(
									at =>
										at.row >= node.start.row &&
										at.row <= node.end.row &&
										at.col >= node.start.col &&
										at.col <= node.end.col &&
										(at.row !== row || at.col !== col)
								)
							: wrongKeys.has(nodeKey(node))
					)

			for (const mismatch of mismatches) {
				const cell = mismatch.cell
				const actual = scalarOf(engine.result(indexByName.get(cell.sheet) ?? -1, cell.row, cell.column))
				const sheet = hfSheet(indexByName.get(cell.sheet) ?? -1)
				const root = sheet !== undefined && !inherits(sheet, cell.row, cell.column)
				const kind = mismatchClass(cell.expected, actual)

				if (root) this.roots++

				for (const name of cell.functions.length === 0 ? ["(operators)"] : cell.functions) {
					if (root) this.rootFunctions.set(name, (this.rootFunctions.get(name) ?? 0) + 1)

					const byKind = this.classes.get(name) ?? new Map<string, number>()

					byKind.set(kind, (byKind.get(kind) ?? 0) + 1)
					this.classes.set(name, byKind)

					const examples = this.examples.get(`${name} ${kind}`) ?? []

					if (root && examples.length < this.exampleLimit)
						examples.push({
							workbook: workbook.name,
							sheet: cell.sheet,
							ref: cell.ref,
							source: cell.source,
							expected: cell.expected,
							actual
						})

					this.examples.set(`${name} ${kind}`, examples)
				}
			}

			const result: RecalcWorkbook = {
				name: workbook.name,
				cells: workbook.cells.length,
				compared: this.comparison.compared - before.compared,
				matched: this.comparison.matched - before.matched,
				fellBack: this.counts.fellBack,
				volatileDependents: this.counts.dependents,
				unsupported: this.counts.unsupported,
				...(tainted.truncated && { error: "volatile dependents truncated" }),
				buildMs: Math.round(performance.now() - started)
			}

			this.workbooks.push(result)

			return result
		} catch (error) {
			return this.failed(workbook, error instanceof Error ? error.message : String(error), started)
		} finally {
			document.close()
			this.current = new Set()
			this.fallbacks = new Set()
		}
	}

	private failed(workbook: ExcelWorkbook, error: string, started: number): RecalcWorkbook {
		const result: RecalcWorkbook = {
			name: workbook.name,
			cells: workbook.cells.length,
			compared: 0,
			matched: 0,
			fellBack: 0,
			volatileDependents: 0,
			unsupported: 0,
			error,
			buildMs: Math.round(performance.now() - started)
		}

		this.workbooks.push(result)

		return result
	}
}
