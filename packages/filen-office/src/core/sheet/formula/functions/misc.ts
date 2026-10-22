import { CellError, ErrorType, FunctionArgumentType, type ImplementedFunctions } from "hyperformula"
import {
	CALC_ERROR,
	naError,
	SheetFunctionPlugin,
	valueError,
	type PluginAst,
	type PluginCall,
	type PluginScalar,
	type PluginState
} from "./plugin"

// ERROR.TYPE and AREAS as support.microsoft.com defines them.

// ERROR.TYPE's number for each error ("ERROR.TYPE function"; #SPILL! 9 and #CALC! 14 are Excel's numbers for the
// dynamic-array errors, which the page's table predates). #NULL! and #GETTING_DATA never reach a formula here.
const ERROR_NUMBERS: Readonly<Partial<Record<ErrorType, number>>> = {
	[ErrorType.DIV_BY_ZERO]: 2,
	[ErrorType.VALUE]: 3,
	[ErrorType.REF]: 4,
	[ErrorType.NAME]: 5,
	[ErrorType.NUM]: 6,
	[ErrorType.NA]: 7,
	[ErrorType.SPILL]: 9
}

// Functions that give a reference (each one area here).
const REFERENCE_FUNCTIONS: ReadonlySet<string> = new Set(["INDEX", "OFFSET", "INDIRECT"])

function unwrap(ast: PluginAst): PluginAst {
	let node = ast

	while (node.type === ("PARENTHESES" as PluginAst["type"]) && "expression" in node) node = node.expression

	return node
}

export class MiscFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		"ERROR.TYPE": { method: "errorType", parameters: [{ argumentType: FunctionArgumentType.SCALAR }] },
		AREAS: { method: "areas", parameters: [{ argumentType: FunctionArgumentType.ANY }] }
	}

	// ERROR.TYPE(error_val): the error's number; anything but an error is #N/A.
	errorType(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("ERROR.TYPE"), (value: PluginScalar) => {
			if (!(value instanceof CellError)) return naError()
			if (value.type === ErrorType.ERROR && value.message === CALC_ERROR) return 14

			return ERROR_NUMBERS[value.type] ?? naError()
		})
	}

	// AREAS(reference): how many areas the reference has. The engine reads no union of references ((A1:B2,C3)), so
	// every reference is one area; a value that is no reference is #VALUE!.
	areas(ast: PluginCall, state: PluginState): unknown {
		const [arg, ...rest] = ast.args

		if (arg === undefined || rest.length > 0) return valueError()

		const node = unwrap(arg)

		switch (node.type) {
			case "CELL_REFERENCE":
			case "CELL_RANGE":
			case "COLUMN_RANGE":
			case "ROW_RANGE":
				return 1
			case "NAMED_EXPRESSION":
			case "FUNCTION_CALL": {
				if (node.type === "FUNCTION_CALL" && !REFERENCE_FUNCTIONS.has(node.procedureName.toUpperCase())) return valueError()

				const value = this.evaluateAst(node, state)

				if (value instanceof CellError) return value

				return node.type === "FUNCTION_CALL" || (typeof value === "object" && "range" in value && value.range !== undefined)
					? 1
					: valueError()
			}
			case "ERROR":
				return this.evaluateAst(node, state)
			default:
				return valueError()
		}
	}
}
