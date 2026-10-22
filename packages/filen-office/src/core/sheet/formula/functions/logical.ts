import { CellError, EmptyValue, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from "hyperformula"
import {
	naError,
	rawNumber,
	SheetFunctionPlugin,
	valueError,
	type PluginCall,
	type PluginScalar,
	type PluginState,
	type PluginValue
} from "./plugin"

// IFS as Excel calculates it (support.microsoft.com "IFS function"): each test in turn, and the value of the first
// that is TRUE, the tests and values after it never calculated (the engine's own calculated them all, so a later
// error was the result). #N/A when no test is TRUE.

// A test's value as a logical value, or the error it gives: text is #VALUE!, an empty cell FALSE.
function truth(value: PluginScalar): boolean | CellError {
	if (value instanceof CellError) return value
	if (typeof value === "boolean") return value
	if (value === EmptyValue) return false
	if (typeof value === "string") return valueError()

	return (rawNumber(value) ?? 0) !== 0
}

function at(value: PluginValue, row: number, col: number): PluginScalar {
	if (!(value instanceof SimpleRangeValue)) return value

	const height = value.height()
	const width = value.width()

	if ((height > 1 && row >= height) || (width > 1 && col >= width)) return naError()

	return value.data[height === 1 ? 0 : row]?.[width === 1 ? 0 : col] ?? EmptyValue
}

function isArray(value: PluginValue): value is SimpleRangeValue {
	return value instanceof SimpleRangeValue && (value.height() > 1 || value.width() > 1)
}

export class LogicalFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		IFS: { method: "ifs", parameters: [{ argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 1, vectorizationForbidden: true }
	}

	ifs(ast: PluginCall, state: PluginState): unknown {
		const args = ast.args

		if (args.length < 2 || args.length % 2 === 1) return naError()

		for (let index = 0; index < args.length; index += 2) {
			const test = this.evaluateAst(args[index] as PluginCall["args"][number], state)

			// Tests over an array give an array: every test and value calculated, position by position.
			if (isArray(test)) return this.elementwise(args, index, test, state)

			const scalar = test instanceof SimpleRangeValue ? (test.data[0]?.[0] ?? EmptyValue) : test
			const met = truth(scalar)

			if (met instanceof CellError) return met
			if (met) return this.evaluateAst(args[index + 1] as PluginCall["args"][number], state)
		}

		return naError()
	}

	private elementwise(args: PluginCall["args"], from: number, first: SimpleRangeValue, state: PluginState): unknown {
		const values: PluginValue[] = []

		for (let index = from; index < args.length; index++)
			values.push(index === from ? first : this.evaluateAst(args[index] as PluginCall["args"][number], state))

		let height = 1
		let width = 1

		for (const value of values) {
			if (!(value instanceof SimpleRangeValue)) continue

			height = Math.max(height, value.height())
			width = Math.max(width, value.width())
		}

		const rows: PluginScalar[][] = []

		for (let row = 0; row < height; row++) {
			const line: PluginScalar[] = []

			for (let col = 0; col < width; col++) {
				let out: PluginScalar = naError()

				for (let index = 0; index < values.length; index += 2) {
					const met = truth(at(values[index] as PluginValue, row, col))

					if (met instanceof CellError) {
						out = met
						break
					}

					if (met) {
						out = at(values[index + 1] as PluginValue, row, col)
						break
					}
				}

				line.push(out)
			}

			rows.push(line)
		}

		return SimpleRangeValue.onlyValues(rows)
	}
}
