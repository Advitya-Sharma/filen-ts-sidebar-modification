import { CellError, EmptyValue, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from "hyperformula"
import { CRITERION_FUNCTION } from "../../functions/engineNames"
import { divError, rawNumber, SheetFunctionPlugin, valueError, type PluginCall, type PluginScalar, type PluginState } from "./plugin"

// Excel's error texts, as a criterion names an error.
const ERROR_TEXT: Readonly<Partial<Record<ErrorType, string>>> = {
	[ErrorType.DIV_BY_ZERO]: "#DIV/0!",
	[ErrorType.NAME]: "#NAME?",
	[ErrorType.VALUE]: "#VALUE!",
	[ErrorType.NUM]: "#NUM!",
	[ErrorType.NA]: "#N/A",
	[ErrorType.REF]: "#REF!",
	[ErrorType.SPILL]: "#SPILL!"
}

// The criteria of COUNTIF and its kind, as Excel reads them.
export class CriteriaFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		[CRITERION_FUNCTION]: {
			method: "criterion",
			parameters: [{ argumentType: FunctionArgumentType.ANY, optionalArg: true }]
		},
		AVERAGEIFS: {
			method: "averageifs",
			parameters: [
				{ argumentType: FunctionArgumentType.RANGE },
				{ argumentType: FunctionArgumentType.RANGE },
				{ argumentType: FunctionArgumentType.NOERROR }
			],
			repeatLastArgs: 2
		}
	}

	// AVERAGEIFS(average_range, criteria_range1, criteria1, ...): the mean of average_range's numbers where every
	// criteria range meets its criterion (support.microsoft.com "AVERAGEIFS function"): ranges of one size (else
	// #VALUE!), empty and non-numeric cells of average_range left out, #DIV/0! when none is left. Criteria are read as
	// the engine reads them for SUMIFS.
	averageifs(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(
			ast.args,
			state,
			this.metadata("AVERAGEIFS"),
			(values: SimpleRangeValue, ...conditions: (SimpleRangeValue | PluginScalar)[]) => {
				const height = values.height()
				const width = values.width()
				const tests: { readonly rows: PluginScalar[][]; readonly test: (value: PluginScalar) => boolean }[] = []

				for (let index = 0; index < conditions.length; index += 2) {
					const range = conditions[index]
					const criterion = conditions[index + 1]

					if (!(range instanceof SimpleRangeValue) || criterion === undefined || criterion instanceof SimpleRangeValue)
						return valueError()
					if (range.height() !== height || range.width() !== width) return valueError()
					if (criterion instanceof CellError) return criterion

					const raw = typeof criterion === "string" || typeof criterion === "boolean" ? criterion : (rawNumber(criterion) ?? 0)
					const built = this.interpreter.criterionBuilder.fromCellValue(raw, this.arithmeticHelper)

					if (built === undefined) return valueError()

					tests.push({
						rows: range.data,
						test: value => built.lambda(typeof value === "object" && !(value instanceof CellError) ? value.val : value)
					})
				}

				let total = 0
				let count = 0
				const rows = values.data

				for (let row = 0; row < height; row++) {
					for (let col = 0; col < width; col++) {
						if (!tests.every(({ rows: tested, test }) => test(tested[row]?.[col] ?? EmptyValue))) continue

						const value = rows[row]?.[col] ?? EmptyValue

						if (value instanceof CellError) return value

						const number = rawNumber(value)

						if (number === undefined) continue

						total += number
						count++
					}
				}

				return count === 0 ? divError() : total / count
			}
		)
	}

	// A criteria argument as written, an empty cell (or none given) as 0: "If the criteria argument is a reference to
	// an empty cell, the COUNTIFS function treats the empty cell as a 0 value" (support.microsoft.com, and the same
	// for SUMIF, SUMIFS, AVERAGEIF and AVERAGEIFS).
	criterion(ast: PluginCall, state: PluginState): unknown {
		const [arg] = ast.args

		if (arg === undefined) return 0

		const value = this.evaluateAst(arg, state)

		if (value instanceof SimpleRangeValue) return value
		// An error matches cells holding that error (COUNTIF(A:A,#N/A) counts them): its text as the criterion.
		if (value instanceof CellError) return ERROR_TEXT[value.type] ?? value

		return value === EmptyValue ? 0 : value
	}
}
