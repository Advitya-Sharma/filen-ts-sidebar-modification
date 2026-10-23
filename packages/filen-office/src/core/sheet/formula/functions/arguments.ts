import { CellError, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, type FunctionArgument } from "hyperformula"

// Arguments Excel truncates to an integer where the engine either takes the fraction (FIND("M",…,3.9) was 8.9) or
// refuses it (EDATE(d,1.9) was #NUM!): truncated before the engine's own bounds apply. Each from its function's page
// on support.microsoft.com ("… is truncated"): CHOOSE's index_num; EDATE's and EOMONTH's months; WEEKDAY's return_type
// (ECMA-376 Part 1 §18.17.7.344); YEARFRAC's basis ("All arguments are truncated to integers"); T.TEST's tails and type; TDIST's
// deg_freedom and tails; the places of DEC2BIN and the other base conversions; BASE's arguments. Where a page is
// silent LibreOffice 7.5 truncates too: MID, REPLACE, FIND, SEARCH and SUBSTITUTE's positions and counts, WEEKNUM's
// return_type, the number DEC2BIN, DEC2OCT and DEC2HEX convert, DECIMAL's radix, DB's period and
// month, CUMIPMT's and CUMPRINC's periods and type, and ROMAN's number. DB's life takes a fraction there.
const TRUNCATED: Readonly<Record<string, readonly number[]>> = {
	CHOOSE: [0],
	MID: [1, 2],
	REPLACE: [1, 2],
	FIND: [2],
	SEARCH: [2],
	SUBSTITUTE: [3],
	EDATE: [1],
	EOMONTH: [1],
	WEEKDAY: [1],
	WEEKNUM: [1],
	YEARFRAC: [2],
	"T.TEST": [2, 3],
	TDIST: [1, 2],
	DEC2BIN: [0, 1],
	DEC2OCT: [0, 1],
	DEC2HEX: [0, 1],
	BIN2OCT: [1],
	BIN2HEX: [1],
	OCT2BIN: [1],
	OCT2HEX: [1],
	HEX2BIN: [1],
	HEX2OCT: [1],
	BASE: [0, 1, 2],
	DECIMAL: [1],
	DB: [3, 4],
	CUMIPMT: [3, 4, 5],
	CUMPRINC: [3, 4, 5],
	ROMAN: [0]
}

const ADJUST = Symbol("excelArgument")

interface Adjustment {
	// The argument as a number, before truncating it.
	readonly loose: FunctionArgument
	// The error an argument out of the function's bounds gives, where Excel's differs from the engine's #NUM!.
	readonly outOfRange?: ErrorType
}

type Adjustable = FunctionArgument & { [ADJUST]?: Adjustment }
type Coerce = (this: FunctionPlugin, arg: unknown, type: Adjustable, state: unknown) => unknown

const LOOSE_NUMBER: FunctionArgument = { argumentType: FunctionArgumentType.NUMBER }

function adjust(name: string, index: number, outOfRange?: ErrorType): void {
	const plugin = HyperFormula.getFunctionPlugin(name)
	const metadata = plugin?.implementedFunctions[name] ?? plugin?.implementedFunctions[plugin.aliases?.[name] ?? ""]
	const parameter = metadata?.parameters?.[index] as Adjustable | undefined

	if (parameter === undefined) throw new Error(`No argument ${String(index)} of ${name}`)

	parameter[ADJUST] = outOfRange === undefined ? { loose: LOOSE_NUMBER } : { loose: LOOSE_NUMBER, outOfRange }
}

let installed = false

// Marks the arguments (the engine reads every call's from the same metadata objects, its own functions' and the
// app's) and wraps the engine's coercion, once, after every function is registered.
export function installExcelArguments(): void {
	if (installed) return

	installed = true

	for (const [name, indices] of Object.entries(TRUNCATED)) for (const index of indices) adjust(name, index)

	// ROMAN: 0 is "" and a number out of 0-3999 is #VALUE! ("ROMAN function": "If number is negative, the #VALUE!
	// error value is returned", and the same past 3999; LibreOffice gives "" for 0).
	const roman = HyperFormula.getFunctionPlugin("ROMAN")?.implementedFunctions["ROMAN"]?.parameters?.[0] as Adjustable | undefined

	if (roman === undefined) throw new Error("No argument 0 of ROMAN")

	roman.minValue = 0
	adjust("ROMAN", 0, ErrorType.VALUE)

	const life = HyperFormula.getFunctionPlugin("DB")?.implementedFunctions["DB"]?.parameters?.[2]

	if (life === undefined) throw new Error("No argument 2 of DB")

	life.argumentType = FunctionArgumentType.NUMBER

	const prototype = FunctionPlugin.prototype as unknown as { coerceToType: Coerce }
	const coerce = prototype.coerceToType

	prototype.coerceToType = function (arg, type, state) {
		const adjustment = type[ADJUST]

		if (adjustment === undefined) return coerce.call(this, arg, type, state)

		// Most calls pass a whole number: coerced once, as the engine would.
		const direct = coerce.call(this, arg, type, state)

		if (typeof direct === "number" && Number.isInteger(direct)) return direct

		const number = coerce.call(this, arg, adjustment.loose, state)

		if (number === undefined || number instanceof CellError) return number

		const raw = typeof number === "number" ? number : (number as { val: number }).val
		const coerced = coerce.call(this, Math.trunc(raw), type, state)

		return adjustment.outOfRange !== undefined && coerced instanceof CellError && coerced.type === ErrorType.NUM
			? new CellError(adjustment.outOfRange)
			: coerced
	}
}
