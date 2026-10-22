import {
	CellError,
	ErrorType,
	FunctionArgumentType,
	HyperFormula,
	SimpleRangeValue,
	type ImplementedFunctions,
	type SimpleCellAddress
} from "hyperformula"
import { MAX_COLUMNS, MAX_ROWS } from "../../model/model"
import { singleReference } from "../../refs/formulaRefs"
import { refError, SheetFunctionPlugin, type PluginCall, type PluginRange, type PluginState } from "./plugin"

// INDIRECT(ref_text, [a1]) (support.microsoft.com "INDIRECT function"): the reference a text names, in A1 or R1C1
// style, a defined name's included. What it reads is known only once it has calculated, so the engine does not link
// it to its cells: each call records the cells it read (IndirectHost.read), and the engine running it calculates it
// again when one of them changes (formulaEngine.ts), as Excel recalculates it.

// A block of an engine sheet read by INDIRECT: rows and columns inclusive, MAX_ROWS (MAX_COLUMNS) as the end of a whole
// column (row).
export interface IndirectRead {
	readonly sheet: number
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

export interface IndirectHost {
	// The reference text a defined name stands for, the scope of the engine sheet `own` first, and the sheet it is read
	// on; undefined for no such name.
	name(name: string, own: number): { readonly text: string; readonly sheet: number } | undefined
	// The cell `formula` reads `read` (each calculation reports all it reads, the first read replacing the last
	// calculation's); null as it starts calculating, so one that reads nothing (a sheet or name not there) is known
	// too, and calculated again when the sheets change.
	read(formula: SimpleCellAddress, read: IndirectRead | null): void
}

let host: IndirectHost | null = null

export function withIndirect<T>(next: IndirectHost | null, work: () => T): T {
	const previous = host

	host = next

	try {
		return work()
	} finally {
		host = previous
	}
}

// Records what the formula at `formula` read where the engine did not link it, for the engine running now.
export function recordRead(formula: SimpleCellAddress, read: IndirectRead): void {
	host?.read(formula, read)
}

// The engine's range of a block, or null before installIndirect.
export function cellRange(read: IndirectRead): PluginRange | null {
	return classes === null ? null : classes.cells(read.sheet, read.startRow, read.startCol, read.endRow, read.endCol)
}

// The engine's range classes, which it does not export: found from a throwaway engine's ranges once.
interface RangeClasses {
	readonly cells: (sheet: number, startRow: number, startCol: number, endRow: number, endCol: number) => PluginRange
	readonly cols: (sheet: number, start: number, end: number) => PluginRange
	readonly rows: (sheet: number, start: number, end: number) => PluginRange
}

let classes: RangeClasses | null = null

// What a formula cell not calculated yet gives where INDIRECT led to it: the engine, which never linked the two, would
// fail the whole calculation. The INDIRECT is calculated again once that cell has (formulaEngine.ts followIndirect).
export const PENDING = "#FILEN.PENDING"

// Finds the engine's range classes, and makes a formula cell not yet calculated read as PENDING. Once, before the first
// engine is built.
export function installIndirect(licenseKey: string): void {
	if (classes !== null) return

	const probe = HyperFormula.buildFromArray([["=SUM(A3:B4)", "=SUM(C:C)", "=SUM(5:5)", "=SEQUENCE(2)"]], { licenseKey })

	try {
		const graph = probe.dependencyGraph

		for (const col of [0, 3]) {
			const vertex = graph.addressMapping.getCell({ sheet: 0, row: 0, col }) as
				{ getCellValue?: unknown; isComputed?: () => boolean } | undefined
			const prototype = vertex === undefined ? undefined : (Object.getPrototypeOf(vertex) as { getCellValue?: unknown })
			const read = prototype?.getCellValue

			if (prototype === undefined || typeof read !== "function") throw new Error("spreadsheet: no formula vertex")

			const original = read as (this: { isComputed(): boolean }) => unknown

			prototype.getCellValue = function (this: { isComputed(): boolean }) {
				return this.isComputed() ? original.call(this) : new CellError(ErrorType.ERROR, PENDING)
			}
		}

		const finite = graph.getRange({ sheet: 0, row: 2, col: 0 }, { sheet: 0, row: 3, col: 1 })?.range
		let column: PluginRange | undefined
		let row: PluginRange | undefined

		for (const vertex of graph.graph.getInfiniteRanges()) {
			const range = (vertex as { range?: PluginRange }).range

			if (range === undefined) continue
			if (range.end.row === Number.POSITIVE_INFINITY) column = range
			else row = range
		}

		if (finite === undefined || column === undefined || row === undefined) throw new Error("spreadsheet: no range classes")

		type Cells = new (start: SimpleCellAddress, end: SimpleCellAddress) => PluginRange
		type Axis = new (sheet: number, start: number, end: number) => PluginRange

		const Cells = finite.constructor as Cells
		const Cols = column.constructor as Axis
		const Rows = row.constructor as Axis

		classes = {
			cells: (sheet, startRow, startCol, endRow, endCol) =>
				new Cells({ sheet, row: startRow, col: startCol }, { sheet, row: endRow, col: endCol }),
			cols: (sheet, start, end) => new Cols(sheet, start, end),
			rows: (sheet, start, end) => new Rows(sheet, start, end)
		}
	} finally {
		probe.destroy()
	}
}

// A block of a sheet as text names it, 0-based, or null.
interface Block {
	readonly sheet: string | null
	readonly kind: "cell" | "cols" | "rows"
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

const R1C1_PART = /^(?:R(\[-?\d+\]|\d+)?)?(?:C(\[-?\d+\]|\d+)?)?$/i

// One R1C1 coordinate: absolute (R3), relative (R[-1]) or the formula's own (R). Null past the grid.
function coordinate(text: string | undefined, own: number, limit: number): number | null {
	if (text === undefined || text === "") return own

	const value = text.startsWith("[") ? own + Number(text.slice(1, -1)) : Number(text) - 1

	return Number.isInteger(value) && value >= 0 && value < limit ? value : null
}

// An R1C1 reference (R2C3, R[-1]C, R1:R3, C[2], R1C1:R2C2) relative to the formula's cell.
function r1c1Block(text: string, row: number, col: number): Block | null {
	let sheet: string | null = null
	let rest = text
	const bang = text.lastIndexOf("!")

	if (bang >= 0) {
		const prefix = text.slice(0, bang)

		sheet = prefix.startsWith("'") && prefix.endsWith("'") && prefix.length >= 2 ? prefix.slice(1, -1).replaceAll("''", "'") : prefix
		rest = text.slice(bang + 1)

		if (sheet === "") return null
	}

	const pieces = rest.split(":")

	if (pieces.length > 2 || pieces.some(piece => piece === "")) return null

	const read = pieces.map(piece => {
		const match = R1C1_PART.exec(piece)

		if (match === null) return null

		const hasRow = /^R/i.test(piece)
		const hasCol = /C/i.test(piece.replace(/^R(\[-?\d+\]|\d+)?/i, ""))

		return {
			hasRow,
			hasCol,
			row: hasRow ? coordinate(match[1], row, MAX_ROWS) : null,
			col: hasCol ? coordinate(match[2], col, MAX_COLUMNS) : null
		}
	})

	const first = read[0]
	const last = read[1] ?? first

	if (
		first === null ||
		first === undefined ||
		last === null ||
		last === undefined ||
		first.hasRow !== last.hasRow ||
		first.hasCol !== last.hasCol
	)
		return null

	if (first.hasRow && first.hasCol) {
		if (first.row === null || first.col === null || last.row === null || last.col === null) return null

		return {
			sheet,
			kind: "cell",
			startRow: Math.min(first.row, last.row),
			startCol: Math.min(first.col, last.col),
			endRow: Math.max(first.row, last.row),
			endCol: Math.max(first.col, last.col)
		}
	}

	if (first.hasRow) {
		if (first.row === null || last.row === null) return null

		return {
			sheet,
			kind: "rows",
			startRow: Math.min(first.row, last.row),
			endRow: Math.max(first.row, last.row),
			startCol: 0,
			endCol: 0
		}
	}

	if (first.col === null || last.col === null) return null

	return { sheet, kind: "cols", startCol: Math.min(first.col, last.col), endCol: Math.max(first.col, last.col), startRow: 0, endRow: 0 }
}

function a1Block(text: string): Block | null {
	const ref = singleReference(text)

	if (ref === null || ref.external || ref.lastSheet !== null || ref.area === null) return null

	const area = ref.area

	return {
		sheet: ref.sheet,
		kind: area.kind,
		startRow: Math.min(area.startRow, area.endRow),
		startCol: Math.min(area.startCol, area.endCol),
		endRow: Math.max(area.startRow, area.endRow),
		endCol: Math.max(area.startCol, area.endCol)
	}
}

// A defined name as INDIRECT reads one: letters, digits, "_", "\" and ".", starting with a letter, "_" or "\".
const NAME = /^(?:(?:'(?:[^']|'')+'|[^!'"]+)!)?[A-Za-z_\\À-￿][\w.\\À-￿]*$/

// The most times a defined name may lead to another name.
const MAX_NAMES = 16

export class IndirectFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		INDIRECT: {
			method: "indirect",
			parameters: [{ argumentType: FunctionArgumentType.STRING }, { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true }],
			isDependentOnSheetStructureChange: true
		}
	}

	// An engine sheet id by name (case-insensitive), or undefined.
	private sheetId(name: string): number | undefined {
		return this.dependencyGraph.sheetMapping.getSheetId(name)
	}

	indirect(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("INDIRECT"), (text: string, a1: boolean) => {
			const reading = host

			if (reading === null) return refError()

			const formula = state.formulaAddress

			reading.read(formula, null)

			let block: Block | null = null
			let own = formula.sheet
			let pending = text.trim()

			for (let hops = 0; hops <= MAX_NAMES && block === null; hops++) {
				block = a1 ? a1Block(pending) : r1c1Block(pending, formula.row, formula.col)

				if (block !== null || !NAME.test(pending)) break

				const bang = pending.lastIndexOf("!")
				const scope = bang < 0 ? own : this.sheetId(pending.slice(0, bang).replace(/^'|'$/g, "").replaceAll("''", "'"))

				if (scope === undefined) return refError()

				const named = reading.name(bang < 0 ? pending : pending.slice(bang + 1), scope)

				if (named === undefined) return refError()

				// A name's own text is A1 style, whatever INDIRECT was asked for.
				a1 = true
				own = named.sheet
				pending = named.text.trim()
			}

			if (block === null) return refError()

			const sheet = block.sheet === null ? own : this.sheetId(block.sheet)

			if (sheet === undefined) return refError()

			const made = classes

			if (made === null) return refError()
			const read: IndirectRead = {
				sheet,
				startRow: block.kind === "cols" ? 0 : block.startRow,
				startCol: block.kind === "rows" ? 0 : block.startCol,
				endRow: block.kind === "cols" ? MAX_ROWS : block.endRow,
				endCol: block.kind === "rows" ? MAX_COLUMNS : block.endCol
			}

			// The formula's own cell among what it reads: a circle, as Excel finds it.
			if (
				sheet === formula.sheet &&
				formula.row >= read.startRow &&
				formula.row <= read.endRow &&
				formula.col >= read.startCol &&
				formula.col <= read.endCol
			)
				return new CellError(ErrorType.CYCLE)

			reading.read(formula, read)

			const range =
				block.kind === "cols"
					? made.cols(sheet, block.startCol, block.endCol)
					: block.kind === "rows"
						? made.rows(sheet, block.startRow, block.endRow)
						: made.cells(sheet, block.startRow, block.startCol, block.endRow, block.endCol)

			return SimpleRangeValue.onlyRange(range, this.dependencyGraph)
		})
	}
}
