import { ArraySize, CellError, ErrorType, FunctionPlugin, HyperFormula, SimpleRangeValue, type SimpleCellAddress } from "hyperformula"
import { MAX_COLUMNS, MAX_ROWS } from "../../model/model"

// The engine's plugin types, read from the plugin class (HyperFormula does not export them by name).
type RunFunction = FunctionPlugin["runFunction"]
export type PluginAst = Parameters<RunFunction>[0][number]
export type PluginState = Parameters<RunFunction>[1]
export type PluginValue = ReturnType<FunctionPlugin["evaluateAst"]>
export type PluginScalar = Exclude<PluginValue, SimpleRangeValue>
export type PluginRange = NonNullable<SimpleRangeValue["range"]>
type Helper = FunctionPlugin["arithmeticHelper"]
type Interpreter = ConstructorParameters<typeof FunctionPlugin>[0]

// A function call as the engine hands it to a plugin method.
export interface PluginCall {
	readonly args: PluginAst[]
	readonly procedureName: string
}

// Whether a row of a sheet (the engine's sheet id) is shown, hidden, or hidden by a filter, as the engine running now
// says (formulaEngine.ts sets it around every calculation).
export const ROW_SHOWN = 0
export const ROW_HIDDEN = 1
export const ROW_FILTERED = 2

export type RowState = (sheet: number, row: number) => number

let rowState: RowState | null = null

export function withRowState<T>(state: RowState | null, work: () => T): T {
	const previous = rowState

	rowState = state

	try {
		return work()
	} finally {
		rowState = previous
	}
}

export function rowStateOf(sheet: number, row: number): number {
	return rowState === null ? ROW_SHOWN : rowState(sheet, row)
}

// The engine's helpers adjusted to Excel, once per engine (they are shared by every plugin of one engine).
const adjusted = new WeakSet<object>()

// Where the engine's own coercion differs from Excel's, for every function and operator of one engine: text that is
// empty ("" from a formula) is not a number, so ""*2 and ""+1 are #VALUE! (the engine read it as 0). An empty cell
// still is 0.
function adjustToExcel(helper: Helper): void {
	if (adjusted.has(helper)) return

	adjusted.add(helper)

	const coerce = helper.coerceNonDateScalarToMaybeNumber.bind(helper)

	helper.coerceNonDateScalarToMaybeNumber = arg => (arg === "" ? undefined : coerce(arg))
}

// Every plugin of the app's own: the engine it runs in coerces as Excel does (adjustToExcel).
export abstract class SheetFunctionPlugin extends FunctionPlugin {
	constructor(interpreter: Interpreter) {
		super(interpreter)
		adjustToExcel(this.arithmeticHelper)
	}

	// A number from a value as the engine coerces one (#VALUE! for text that reads as none).
	protected numberOf(value: PluginScalar): number | CellError {
		const number = this.coerceScalarToNumberOrError(value)

		return number instanceof CellError ? number : typeof number === "number" ? number : number.val
	}

	// The formula text of a cell, or undefined for a cell without one.
	protected formulaAt(address: SimpleCellAddress): string | undefined {
		return this.serialization.getCellFormula(address)
	}
}

// The error a function's JavaScript failure gives its cell: a number out of what JavaScript can work with (a
// RangeError: digits, lengths, a call stack too deep) is #NUM!, anything else #VALUE!. The engine would let it escape
// the calculation and fail every later edit of the workbook.
function thrown(error: unknown): CellError {
	if (error instanceof CellError) return error

	return new CellError(error instanceof RangeError ? ErrorType.NUM : ErrorType.VALUE)
}

const guarded = new WeakSet<object>()

type PluginMethod = (this: unknown, ...args: unknown[]) => unknown

// Every function the engine knows (its own and the app's, registered by now) gives a failure as its cell's error.
// The engine calls them by method name on each plugin class's instance, so each class's methods are wrapped, once.
export function guardFunctionPlugins(): void {
	for (const plugin of HyperFormula.getAllFunctionPlugins()) {
		const prototype = plugin.prototype as unknown as Record<string, unknown>

		for (const definition of Object.values(plugin.implementedFunctions)) {
			for (const [name, failed] of [
				[definition.method, thrown],
				[definition.sizeOfResultArrayMethod, () => ArraySize.scalar()]
			] as const) {
				const method = name === undefined ? undefined : prototype[name]

				if (name === undefined || typeof method !== "function" || guarded.has(method)) continue

				const run = method as PluginMethod
				const wrapped: PluginMethod = function (this: unknown, ...args: unknown[]) {
					try {
						return run.apply(this, args)
					} catch (error) {
						return failed(error)
					}
				}

				guarded.add(wrapped)
				prototype[name] = wrapped
			}
		}
	}
}

// The cells of a whole column (row) past where the engine ends reading it (formula/wholeRanges.ts), every one of
// them empty: Excel's columns have all 1,048,576 rows and its rows 16,384 columns.
export function wholeTail(value: SimpleRangeValue): number {
	const range = value.range

	if (range === undefined) return 0
	if (range.end.row === Number.POSITIVE_INFINITY) return Math.max(0, MAX_ROWS - range.start.row - value.height()) * value.width()
	if (range.end.col === Number.POSITIVE_INFINITY) return Math.max(0, MAX_COLUMNS - range.start.col - value.width()) * value.height()

	return 0
}

export function valueError(): CellError {
	return new CellError(ErrorType.VALUE)
}

export function naError(): CellError {
	return new CellError(ErrorType.NA)
}

export function numError(): CellError {
	return new CellError(ErrorType.NUM)
}

export function divError(): CellError {
	return new CellError(ErrorType.DIV_BY_ZERO)
}

// Excel's #CALC! (an array function given nothing to return), which the engine has no type of its own for: an
// error of its generic type with this message, given out as "#CALC!" (formulaEngine.ts).
export const CALC_ERROR = "#CALC!"

export function calcError(): CellError {
	return new CellError(ErrorType.ERROR, CALC_ERROR)
}

// What a function gives where it reads a whole column past the engine's end and cannot tell what Excel's empty rows
// there give (formula/wholeRanges.ts): the engine's generic error with this message. The formula keeps the file's
// result (formulaEngine.ts), or shows the engine failing on it.
export const UNMODELED = "#FILEN.UNMODELED"

export function unmodeledError(): CellError {
	return new CellError(ErrorType.ERROR, UNMODELED)
}

export function refError(): CellError {
	return new CellError(ErrorType.REF)
}

// The plain number of an engine number (dates, percentages and currencies carry their format).
export function rawNumber(value: PluginScalar): number | undefined {
	if (typeof value === "number") return value
	if (typeof value === "object" && !(value instanceof CellError)) return value.val

	return undefined
}
