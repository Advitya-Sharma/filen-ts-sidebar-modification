import { EmptyValue, HyperFormula, SimpleRangeValue } from "hyperformula"
import { MAX_COLUMNS, MAX_ROWS } from "../../model/model"
import { rawNumber, unmodeledError, wholeTail, type PluginAst, type PluginCall, type PluginState } from "./plugin"

// A whole column's (row's) cells past where the engine ends reading it (formula/wholeRanges.ts) are all empty, and
// Excel counts them: COUNTIF(A:A,"") is 1,048,576 less the column's filled cells. The engine's COUNTIF, COUNTIFS
// and SUMPRODUCT read only to that end; here they add what the empty rest gives, without reading it: once per empty
// cell for a criterion an empty cell meets, and for SUMPRODUCT the value of one empty row times their number.

// The parts of the engine's plugin instance used here (protected on the class).
export interface Plugin {
	evaluateAst(ast: PluginAst, state: PluginState): unknown
	readonly interpreter: {
		readonly criterionBuilder: { fromCellValue(raw: unknown, helper: unknown): { lambda(value: unknown): boolean } | undefined }
	}
	readonly arithmeticHelper: unknown
}

type Method = (this: Plugin, ast: PluginCall, state: PluginState) => unknown

// The empty cells past the end of the whole range `ast` reads, or 0 (not a whole column or row).
function tailOf(plugin: Plugin, ast: PluginAst | undefined, state: PluginState): number {
	if (ast === undefined) return 0

	const value = plugin.evaluateAst(ast, state)

	return value instanceof SimpleRangeValue ? wholeTail(value) : 0
}

// Whether an empty cell meets a criterion argument (an array of criteria answers no: the engine gives an array then).
function meetsEmpty(plugin: Plugin, ast: PluginAst | undefined, state: PluginState): boolean {
	if (ast === undefined) return false

	const value = plugin.evaluateAst(ast, state)
	const raw =
		typeof value === "string" || typeof value === "boolean" ? value : typeof value === "object" ? rawNumber(value as never) : value

	if (typeof raw !== "string" && typeof raw !== "boolean" && typeof raw !== "number") return false

	return plugin.interpreter.criterionBuilder.fromCellValue(raw, plugin.arithmeticHelper)?.lambda(EmptyValue) === true
}

// COUNTIF(range, criterion) and COUNTIFS(range1, criterion1, …): the empty rest, counted where every criterion meets an
// empty cell (the ranges have one size, else the engine gave #VALUE!).
function countingBlanks(run: Method): Method {
	return function (ast, state) {
		const counted = run.call(this, ast, state)

		if (typeof counted !== "number") return counted

		const tail = tailOf(this, ast.args[0], state)

		if (tail === 0) return counted

		for (let index = 0; index < ast.args.length; index += 2) {
			if (index > 0 && tailOf(this, ast.args[index], state) !== tail) return counted
			if (!meetsEmpty(this, ast.args[index + 1], state)) return counted
		}

		return counted + tail
	}
}

// Functions an empty row's value is read through cell by cell, the same in every row (no position, no aggregate).
const ELEMENTWISE = new Set([
	"ABS",
	"EXACT",
	"FALSE",
	"IF",
	"IFERROR",
	"IFNA",
	"INT",
	"ISBLANK",
	"ISERR",
	"ISERROR",
	"ISEVEN",
	"ISLOGICAL",
	"ISNA",
	"ISNONTEXT",
	"ISNUMBER",
	"ISODD",
	"ISTEXT",
	"LEFT",
	"LEN",
	"LOWER",
	"MID",
	"N",
	"NOT",
	"PROPER",
	"RIGHT",
	"SIGN",
	"TRIM",
	"TRUE",
	"UPPER",
	"VALUE"
])

const OPERATORS = new Set([
	"MINUS_UNARY_OP",
	"PLUS_UNARY_OP",
	"PERCENT_OP",
	"CONCATENATE_OP",
	"EQUALS_OP",
	"NOT_EQUAL_OP",
	"GREATER_THAN_OP",
	"LESS_THAN_OP",
	"GREATER_THAN_OR_EQUAL_OP",
	"LESS_THAN_OR_EQUAL_OP",
	"PLUS_OP",
	"MINUS_OP",
	"TIMES_OP",
	"DIV_OP",
	"POWER_OP"
])

type Node = PluginAst & { readonly type: string; readonly [key: string]: unknown }

// Functions read the same in every row an empty one is (no position, no aggregate over the column).
const UNIFORM = new Set([...ELEMENTWISE])
// And, read at one row, those that give its position.
const POSITIONAL = new Set([...ELEMENTWISE, "ROW", "COLUMN"])

// Functions giving one value whatever the row, of whole columns read as ranges: right over the engine's end when
// every whole column they read is an argument as it is (blanks add nothing to them).
const RANGE_AGGREGATES = new Set([
	"AGGREGATE",
	"AVERAGE",
	"AVERAGEIF",
	"AVERAGEIFS",
	"COUNT",
	"COUNTA",
	"HLOOKUP",
	"INDEX",
	"LARGE",
	"MATCH",
	"MAX",
	"MAXIFS",
	"MEDIAN",
	"MIN",
	"MINIFS",
	"PRODUCT",
	"SMALL",
	"SUBTOTAL",
	"SUM",
	"SUMIF",
	"SUMIFS",
	"VLOOKUP"
])

// Functions that read a whole column's empty rest themselves (here, or in lookup.ts), or refuse.
const SELF_GUARDED = new Set(["COUNTBLANK", "COUNTIF", "COUNTIFS", "LOOKUP", "SUMPRODUCT", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "XMATCH"])

const WHOLE = new Set(["COLUMN_RANGE", "ROW_RANGE", "NAMED_EXPRESSION"])

function unwrapped(ast: Node): Node {
	let node = ast

	while (node.type === "PARENTHESES") node = node["expression"] as Node

	return node
}

function children(ast: Node): Node[] {
	if (ast.type === "FUNCTION_CALL") return ast["args"] as Node[]
	if (ast.type === "PARENTHESES") return [ast["expression"] as Node]

	return ["value", "left", "right"].flatMap(key => {
		const part = ast[key]

		return typeof part === "object" && part !== null && "type" in part ? [part as Node] : []
	})
}

// Whether a whole column or row (or a name, which may be one) is read anywhere in the argument.
function holdsWhole(ast: Node): boolean {
	return WHOLE.has(ast.type) || (ast.type !== "ARRAY" && children(ast).some(holdsWhole))
}

// The most empty cells past the engine's end of any whole column (row) read in the argument.
function tailIn(plugin: Plugin, ast: Node, state: PluginState): number {
	if (WHOLE.has(ast.type)) {
		const value = plugin.evaluateAst(ast, state)

		return value instanceof SimpleRangeValue ? wholeTail(value) : 0
	}

	return ast.type === "ARRAY" ? 0 : Math.max(0, ...children(ast).map(child => tailIn(plugin, child, state)))
}

interface FixedRange {
	readonly sheet: number
	readonly start: { readonly row: number; readonly col: number }
	readonly end: { readonly row: number; readonly col: number }
}

// A whole column (row) as its cells in the grid's last row (column): a reference the engine reads as any other.
function lastLine(range: FixedRange, columns: boolean): Node {
	const at = (col: number, row: number) => ({ sheet: range.sheet, toSimpleCellAddress: () => ({ sheet: range.sheet, col, row }) })

	return (columns
		? { type: "CELL_RANGE", start: at(range.start.col, MAX_ROWS - 1), end: at(range.end.col, MAX_ROWS - 1) }
		: { type: "CELL_RANGE", start: at(MAX_COLUMNS - 1, range.start.row), end: at(MAX_COLUMNS - 1, range.end.row) }) as unknown as Node
}

function exactMatch(ast: Node | undefined): boolean {
	const node = ast === undefined ? undefined : unwrapped(ast)

	return node?.type === "NUMBER" && node["value"] === 0
}

// The argument read at the grid's last row (column), each whole column (row) in it as its cells there, with the empty
// cells past where the engine ends reading them added to `tails` (per row, or per column); null where the engine
// cannot read it so (a whole column through a function not in `allowed` that reads it by row, or is unknown here).
function atLastLine(plugin: Plugin, ast: Node, state: PluginState, tails: Set<number>, allowed: ReadonlySet<string>): Node | null {
	if (WHOLE.has(ast.type)) {
		const value = plugin.evaluateAst(ast, state)
		const range = value instanceof SimpleRangeValue ? (value.range as FixedRange | undefined) : undefined

		if (value instanceof SimpleRangeValue && range !== undefined) {
			const columns = range.end.row === Number.POSITIVE_INFINITY

			if (columns || range.end.col === Number.POSITIVE_INFINITY) {
				tails.add(wholeTail(value) / (columns ? value.width() : value.height()))

				return lastLine(range, columns)
			}
		}

		return ast.type === "NAMED_EXPRESSION" ? ast : null
	}

	switch (ast.type) {
		case "NUMBER":
		case "STRING":
		case "CELL_REFERENCE":
		case "CELL_RANGE":
		case "EMPTY":
		case "ERROR":
		case "ARRAY":
			return ast
		case "PARENTHESES": {
			const expression = atLastLine(plugin, ast["expression"] as Node, state, tails, allowed)

			return expression === null ? null : { ...ast, expression }
		}
		case "FUNCTION_CALL": {
			const name = String(ast["procedureName"])
			const args = ast["args"] as Node[]

			if (SELF_GUARDED.has(name) || !holdsWhole(ast)) return ast
			if (RANGE_AGGREGATES.has(name) && args.every(arg => !holdsWhole(arg) || WHOLE.has(unwrapped(arg).type))) return ast
			// An exact MATCH finds the first: the engine's rows hold an empty one, the same as all past its end.
			if (name === "MATCH" && exactMatch(args[2]) && args.every(arg => atLastLine(plugin, arg, state, new Set(), UNIFORM) !== null))
				return ast
			if (!allowed.has(name)) return null

			const next: Node[] = []

			for (const arg of args) {
				const read = atLastLine(plugin, arg, state, tails, allowed)

				if (read === null) return null

				next.push(read)
			}

			return { ...ast, args: next }
		}
		default: {
			if (!OPERATORS.has(ast.type)) return null

			const parts: Record<string, Node> = {}

			for (const key of ["value", "left", "right"]) {
				const part = ast[key] as Node | undefined

				if (part === undefined || typeof part !== "object") continue

				const next = atLastLine(plugin, part, state, tails, allowed)

				if (next === null) return null

				parts[key] = next
			}

			return { ...ast, ...parts }
		}
	}
}

// The empty rows (columns) past the end of the whole columns (rows) an argument reads cell by cell, the same in each
// of them (0 when it reads none), or null where what they give cannot be told (read by position, through an
// aggregate of an expression).
export function uniformTail(plugin: Plugin, ast: PluginAst | undefined, state: PluginState): number | null {
	if (ast === undefined) return 0

	const tails = new Set<number>()

	if (atLastLine(plugin, ast as Node, state, tails, UNIFORM) === null) return tailIn(plugin, ast as Node, state) > 0 ? null : 0

	const [tail] = tails

	if (tail === undefined) return 0

	return tails.size === 1 ? tail : null
}

// What an argument gives in the grid's last row (column), its whole columns (rows) read there; undefined where it
// cannot be read so.
export function lastLineValue(plugin: Plugin, ast: PluginAst | undefined, state: PluginState): unknown {
	if (ast === undefined) return undefined

	const node = atLastLine(plugin, ast as Node, state, new Set(), POSITIONAL)

	return node === null ? undefined : plugin.evaluateAst(node, state)
}

// SUMPRODUCT(array1, …) and the SUMX* sums: plus the value of one empty row of their whole columns times the empty
// rows past their end (an error there is the result, as Excel meets it in those rows). Where that row's value cannot
// be told, not calculated (unmodeled: the file's result is kept).
function summingBlanks(run: Method): Method {
	return function (ast, state) {
		const total = run.call(this, ast, state)

		if (typeof total !== "number") return total

		const tails = new Set<number>()
		const args: Node[] = []

		for (const arg of ast.args as Node[]) {
			const next = atLastLine(this, arg, state, tails, UNIFORM)

			if (next === null) return tailIn(this, arg, state) > 0 ? unmodeledError() : total

			args.push(next)
		}

		const [tail] = tails

		if (tail === undefined || tail === 0) return total
		if (tails.size !== 1) return unmodeledError()

		const row = run.call(this, { ...ast, args }, state)

		return typeof row === "number" ? total + row * tail : row
	}
}

// ROW(range) and COLUMN(range) where the call around them evaluates arrays: the rows' (columns') numbers, a whole
// column's as far as the engine reads it (what reads past is told apart above). The engine gave the first alone.
function positions(run: Method, rows: boolean): Method {
	return function (ast, state) {
		const [arg] = ast.args as Node[]

		if (!state.arraysFlag || arg === undefined || ast.args.length !== 1) return run.call(this, ast, state)

		const node = unwrapped(arg)

		if (node.type !== "CELL_RANGE" && !WHOLE.has(node.type)) return run.call(this, ast, state)

		const value = this.evaluateAst(node, state)
		const range = value instanceof SimpleRangeValue ? value.range : undefined

		if (!(value instanceof SimpleRangeValue) || range === undefined) return run.call(this, ast, state)

		const count = rows ? value.height() : value.width()

		if (count <= 1) return run.call(this, ast, state)

		const first = (rows ? range.start.row : range.start.col) + 1
		const numbers = Array.from({ length: count }, (_, index) => first + index)

		return SimpleRangeValue.onlyValues(rows ? numbers.map(number => [number]) : [numbers])
	}
}

// Functions giving an array (the engine's first value of it in an ordinary formula): not calculated over a whole
// column with empty rows past the engine's end, whose reading of those rows here is not Excel's (TAKE(A:A,-2), a
// sort). Those of numbers alone (blanks add nothing) and XLOOKUP's first match are read as they are.
const ARRAYS_OF_NUMBERS = new Set(["FREQUENCY", "MODE.MULT", "XLOOKUP"])

// TRIMRANGE trimming the trailing empty rows and columns (2 or 3, the default), which takes away those past the engine's
// end whatever they would be: read as it is.
function trimsTails(args: readonly Node[]): boolean {
	return [args[1], args[2]].every(mode => mode === undefined || (mode.type === "NUMBER" && (mode["value"] === 2 || mode["value"] === 3)))
}

function refusingTails(run: Method, name: string): Method {
	return function (ast, state) {
		const args = ast.args as Node[]
		const searchMode = name === "XLOOKUP" ? args[5] : undefined
		const plain =
			(ARRAYS_OF_NUMBERS.has(name) && (searchMode === undefined || (searchMode.type === "NUMBER" && searchMode["value"] === 1))) ||
			(name === "TRIMRANGE" && trimsTails(args))

		if (!plain && args.some(arg => holdsWhole(arg) && tailIn(this, arg, state) > 0)) return unmodeledError()

		return run.call(this, ast, state)
	}
}

let installed = false

// Wraps the engine's own functions that read whole columns past its end, once (before every function is guarded,
// plugin.ts). `arrayResults`: the functions giving arrays (refs/formulaRefs.ts).
export function installBlankTails(arrayResults: ReadonlySet<string>): void {
	if (installed) return

	installed = true

	const wrap = (name: string, wrapper: (run: Method) => Method): void => {
		const plugin = HyperFormula.getFunctionPlugin(name)
		const method = plugin?.implementedFunctions[name]?.method

		if (plugin === undefined || method === undefined) return

		const prototype = plugin.prototype as unknown as Record<string, Method>
		const run = prototype[method]

		if (run !== undefined) prototype[method] = wrapper(run)
	}

	wrap("COUNTIF", countingBlanks)
	wrap("COUNTIFS", countingBlanks)

	for (const name of ["SUMPRODUCT", "SUMX2MY2", "SUMX2PY2", "SUMXMY2"]) wrap(name, summingBlanks)

	wrap("ROW", run => positions(run, true))
	wrap("COLUMN", run => positions(run, false))

	for (const name of arrayResults) wrap(name, run => refusingTails(run, name))
}
