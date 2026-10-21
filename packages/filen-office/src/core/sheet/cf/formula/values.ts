import type { SheetId } from "../../protocol"

// The values a rule's formula works with, as Excel's calculation has them: numbers (dates are serials), text,
// logicals, errors and blank. Conditional formats and data validation evaluate their formulas as array formulas
// (support.microsoft.com, "Use formulas with conditional formatting": a formula is written as for the range's
// top-left cell and applied relative to each cell), so a value may also be an array or a reference.

export type FormulaErrorCode = "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#SPILL!" | "#CALC!"

export class FormulaError {
	readonly code: FormulaErrorCode

	private constructor(code: FormulaErrorCode) {
		this.code = code
	}

	static readonly NULL = new FormulaError("#NULL!")
	static readonly DIV0 = new FormulaError("#DIV/0!")
	static readonly VALUE = new FormulaError("#VALUE!")
	static readonly REF = new FormulaError("#REF!")
	static readonly NAME = new FormulaError("#NAME?")
	static readonly NUM = new FormulaError("#NUM!")
	static readonly NA = new FormulaError("#N/A")
	static readonly SPILL = new FormulaError("#SPILL!")
	static readonly CALC = new FormulaError("#CALC!")

	// The error a cell's text names ("#DIV/0!"), #VALUE! for one Excel has no name for here.
	static of(text: string): FormulaError {
		switch (text.trim().toUpperCase()) {
			case "#NULL!":
				return FormulaError.NULL
			case "#DIV/0!":
				return FormulaError.DIV0
			case "#REF!":
				return FormulaError.REF
			case "#NAME?":
				return FormulaError.NAME
			case "#NUM!":
				return FormulaError.NUM
			case "#N/A":
				return FormulaError.NA
			case "#SPILL!":
				return FormulaError.SPILL
			case "#CALC!":
				return FormulaError.CALC
			default:
				return FormulaError.VALUE
		}
	}
}

// One cell's value; null is blank.
export type Scalar = number | string | boolean | FormulaError | null

// A rectangle of cells on a sheet (inclusive).
export interface FormulaRef {
	readonly kind: "ref"
	readonly sheet: SheetId
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

export interface FormulaArray {
	readonly kind: "array"
	readonly rows: number
	readonly cols: number
	// Row-major: the cells of the top-left `stored` rows and columns (all of them without `stored`).
	readonly values: readonly Scalar[]
	// A reference's array holds what the sheet holds; every cell past it is `fill` (blank, or what an operation made
	// of a blank), so a whole column costs what it holds.
	readonly stored?: { readonly rows: number; readonly cols: number }
	readonly fill?: Scalar
}

export type FormulaValue = Scalar | FormulaRef | FormulaArray

// A formula the evaluator cannot calculate as Excel would (a function it lacks, a structured reference, a name it
// cannot resolve to something it reads): the rule is not applied, never guessed.
export class Unsupported extends Error {
	constructor(what: string) {
		super(`Not calculated here: ${what}`)
		this.name = "Unsupported"
	}
}

export function isRef(value: FormulaValue): value is FormulaRef {
	return typeof value === "object" && value !== null && !(value instanceof FormulaError) && value.kind === "ref"
}

export function isArray(value: FormulaValue): value is FormulaArray {
	return typeof value === "object" && value !== null && !(value instanceof FormulaError) && value.kind === "array"
}

export function refRows(ref: FormulaRef): number {
	return ref.endRow - ref.startRow + 1
}

export function refCols(ref: FormulaRef): number {
	return ref.endCol - ref.startCol + 1
}

export function storedRows(array: FormulaArray): number {
	return array.stored?.rows ?? array.rows
}

export function storedCols(array: FormulaArray): number {
	return array.stored?.cols ?? array.cols
}

// The cells of an array past what it stores.
export function tailCells(array: FormulaArray): number {
	return array.stored === undefined ? 0 : array.rows * array.cols - array.stored.rows * array.stored.cols
}

export function arrayAt(array: FormulaArray, row: number, col: number): Scalar {
	if (array.stored === undefined) return array.values[row * array.cols + col] ?? null

	return row < array.stored.rows && col < array.stored.cols ? (array.values[row * array.stored.cols + col] ?? null) : (array.fill ?? null)
}

// Every cell of an array, row-major; Unsupported past `max` cells.
export function denseValues(array: FormulaArray, max: number): readonly Scalar[] {
	if (array.stored === undefined) return array.values
	if (array.rows * array.cols > max) throw new Unsupported("an array too large")

	const values: Scalar[] = new Array<Scalar>(array.rows * array.cols)

	for (let r = 0; r < array.rows; r++) for (let c = 0; c < array.cols; c++) values[r * array.cols + c] = arrayAt(array, r, c)

	return values
}

// Excel's own text of a number: up to 15 significant digits, no exponent between 1e-9 and 1e15 (what "&" and
// the text functions see; support.microsoft.com "Excel specifications and limits": 15 digits of precision).
export function numberText(value: number): string {
	if (value === 0) return "0"
	if (!Number.isFinite(value)) return "#NUM!"

	const precise = Number(value.toPrecision(15))
	const magnitude = Math.abs(precise)

	if (magnitude >= 1e-9 && magnitude < 1e15) {
		let text = precise.toFixed(Math.max(0, Math.min(20, 14 - Math.floor(Math.log10(magnitude)))))

		if (text.includes(".")) text = text.replace(/0+$/u, "").replace(/\.$/u, "")

		return text
	}

	const [mantissa = "0", exponent = "0"] = precise.toExponential(14).split("e")
	const trimmed = mantissa.includes(".") ? mantissa.replace(/0+$/u, "").replace(/\.$/u, "") : mantissa
	const power = Number(exponent)

	return `${trimmed}E${power < 0 ? "-" : "+"}${String(Math.abs(power)).padStart(2, "0")}`
}

// A value as text (the text functions' and "&"'s view). Errors pass through.
export function toText(value: Scalar): string | FormulaError {
	if (value instanceof FormulaError) return value
	if (value === null) return ""
	if (typeof value === "boolean") return value ? "TRUE" : "FALSE"
	if (typeof value === "number") return numberText(value)

	return value
}

const NUMERIC = /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*%?\s*$/u

// Text read as a number where arithmetic needs one ("5" + 1); dates and times typed as text are not read here.
export function textNumber(text: string): number | null {
	if (!NUMERIC.test(text)) return null

	const trimmed = text.trim()
	const percent = trimmed.endsWith("%")
	const value = Number(percent ? trimmed.slice(0, -1).trim() : trimmed)

	if (!Number.isFinite(value)) return null

	return percent ? value / 100 : value
}

// A value as a number for arithmetic: blank is 0, logicals 1 and 0, numeric text its number.
export function toNumber(value: Scalar): number | FormulaError {
	if (value instanceof FormulaError) return value
	if (value === null) return 0
	if (typeof value === "boolean") return value ? 1 : 0
	if (typeof value === "number") return value

	return textNumber(value) ?? FormulaError.VALUE
}

// A value as a logical (IF, AND, a rule's result): numbers are TRUE unless 0, "TRUE"/"FALSE" text reads as such.
export function toBoolean(value: Scalar): boolean | FormulaError {
	if (value instanceof FormulaError) return value
	if (value === null) return false
	if (typeof value === "boolean") return value
	if (typeof value === "number") return value !== 0

	const upper = value.toUpperCase()

	if (upper === "TRUE") return true
	if (upper === "FALSE") return false

	return FormulaError.VALUE
}

function typeRank(value: number | string | boolean): number {
	return typeof value === "number" ? 0 : typeof value === "string" ? 1 : 2
}

// Excel's comparison of two values (the = < > operators): numbers before text before logicals, text without
// regard to case, blank as the other side's empty value (0, "" or FALSE). Negative, zero or positive.
export function compareValues(a: Exclude<Scalar, FormulaError>, b: Exclude<Scalar, FormulaError>): number {
	let left = a
	let right = b

	if (left === null && right === null) return 0
	if (left === null) left = typeof right === "number" ? 0 : typeof right === "boolean" ? false : ""
	if (right === null) right = typeof left === "number" ? 0 : typeof left === "boolean" ? false : ""

	const rankA = typeRank(left as number | string | boolean)
	const rankB = typeRank(right as number | string | boolean)

	if (rankA !== rankB) return rankA - rankB
	if (typeof left === "number" && typeof right === "number") return compareNumbers(left, right)
	if (typeof left === "boolean" && typeof right === "boolean") return left === right ? 0 : left ? 1 : -1

	const x = (left as string).toUpperCase()
	const y = (right as string).toUpperCase()

	return x === y ? 0 : x < y ? -1 : 1
}

// Two numbers compared to Excel's precision, 15 significant digits ("Excel specifications and limits"), so 0.1+0.2
// equals 0.3 as in Excel, LibreOffice and the session's engine. Negative, zero or positive.
export function compareNumbers(a: number, b: number): number {
	if (a === b) return 0

	// Numbers that agree to 15 digits differ by at most one unit of the 15th: further apart, they are as they are.
	if (Math.abs(a - b) > Math.max(Math.abs(a), Math.abs(b)) * 1e-14) return a < b ? -1 : 1

	const x = Number(a.toPrecision(15))
	const y = Number(b.toPrecision(15))

	return x === y ? 0 : x < y ? -1 : 1
}
