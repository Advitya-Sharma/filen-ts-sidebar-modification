import { CellError, EmptyValue, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from "hyperformula"
import {
	divError,
	naError,
	numError,
	rawNumber,
	SheetFunctionPlugin,
	valueError,
	type PluginAst,
	type PluginCall,
	type PluginScalar,
	type PluginState
} from "./plugin"
import { kurtosis, mean, mode, modes } from "./statistics"
import { flatten } from "./values"

// Statistical functions the engine lacks, as support.microsoft.com defines them.

const ARRAY_ARGUMENTS = { enableArrayArithmeticForArguments: true, vectorizationForbidden: true } as const
const NUMBERS = { parameters: [{ argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 1, ...ARRAY_ARGUMENTS }
const RANKS = {
	parameters: [
		{ argumentType: FunctionArgumentType.NUMBER },
		{ argumentType: FunctionArgumentType.RANGE },
		{ argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }
	]
}
const FORECASTS = {
	parameters: [
		{ argumentType: FunctionArgumentType.NUMBER },
		{ argumentType: FunctionArgumentType.RANGE },
		{ argumentType: FunctionArgumentType.RANGE }
	],
	...ARRAY_ARGUMENTS
}

function unwrap(ast: PluginAst): PluginAst {
	let node = ast

	while (node.type === ("PARENTHESES" as PluginAst["type"]) && "expression" in node) node = node.expression

	return node
}

// The numbers of a range or array: text, logical values and empty cells left out; the first error raised.
function rangeNumbers(rows: readonly (readonly PluginScalar[])[], out: number[]): CellError | null {
	for (const row of rows)
		for (const value of row) {
			if (value instanceof CellError) return value

			const number = rawNumber(value)

			if (number !== undefined) out.push(number)
		}

	return null
}

// Paired numbers of two arrays of one size, by position: a pair is left out where either is not a number. #N/A when
// the sizes differ.
function pairs(ys: SimpleRangeValue, xs: SimpleRangeValue): { ys: number[]; xs: number[] } | CellError {
	const yValues = flatten(ys.data)
	const xValues = flatten(xs.data)

	if (yValues.length !== xValues.length) return naError()

	const out = { ys: [] as number[], xs: [] as number[] }

	for (let index = 0; index < yValues.length; index++) {
		const y = yValues[index]
		const x = xValues[index]

		if (y instanceof CellError) return y
		if (x instanceof CellError) return x

		const yNumber = y === undefined ? undefined : rawNumber(y)
		const xNumber = x === undefined ? undefined : rawNumber(x)

		if (yNumber === undefined || xNumber === undefined) continue

		out.ys.push(yNumber)
		out.xs.push(xNumber)
	}

	return out
}

// Least squares over columns of predictors (each row one observation), with or without an intercept: the
// coefficients [b, m1, m2, ...] (b 0 without one), or #NUM! when the predictors are not independent.
function leastSquares(ys: readonly number[], xs: readonly (readonly number[])[], intercept: boolean): number[] | CellError {
	const k = xs[0]?.length ?? 0
	const size = k + (intercept ? 1 : 0)
	// The normal equations, augmented: [XᵀX | Xᵀy].
	const matrix = Array.from({ length: size }, () => new Array<number>(size + 1).fill(0))

	ys.forEach((y, row) => {
		const line = intercept ? [1, ...(xs[row] ?? [])] : [...(xs[row] ?? [])]

		for (let i = 0; i < size; i++) {
			const target = matrix[i] as number[]

			for (let j = 0; j < size; j++) target[j] = (target[j] ?? 0) + (line[i] ?? 0) * (line[j] ?? 0)

			target[size] = (target[size] ?? 0) + (line[i] ?? 0) * y
		}
	})

	// Gaussian elimination with partial pivoting.
	for (let column = 0; column < size; column++) {
		let pivot = column

		for (let row = column + 1; row < size; row++)
			if (Math.abs(matrix[row]?.[column] ?? 0) > Math.abs(matrix[pivot]?.[column] ?? 0)) pivot = row

		const top = matrix[pivot] as number[]

		if (Math.abs(top[column] ?? 0) < 1e-12) return numError()

		matrix[pivot] = matrix[column] as number[]
		matrix[column] = top

		for (let row = 0; row < size; row++) {
			if (row === column) continue

			const line = matrix[row] as number[]
			const factor = (line[column] ?? 0) / (top[column] ?? 1)

			for (let j = column; j <= size; j++) line[j] = (line[j] ?? 0) - factor * (top[j] ?? 0)
		}
	}

	const solved = matrix.map((line, index) => (line[size] ?? 0) / (line[index] ?? 1))

	return intercept ? solved : [0, ...solved]
}

export class StatisticalFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		RANK: { method: "rankEq", ...RANKS },
		"RANK.EQ": { method: "rankEq", ...RANKS },
		"RANK.AVG": { method: "rankAvg", ...RANKS },
		FORECAST: { method: "forecast", ...FORECASTS },
		"FORECAST.LINEAR": { method: "forecast", ...FORECASTS },
		TREND: {
			method: "trend",
			parameters: [
				{ argumentType: FunctionArgumentType.RANGE },
				{ argumentType: FunctionArgumentType.RANGE, optionalArg: true },
				{ argumentType: FunctionArgumentType.RANGE, optionalArg: true },
				{ argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true }
			],
			...ARRAY_ARGUMENTS
		},
		FREQUENCY: {
			method: "frequency",
			parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }],
			...ARRAY_ARGUMENTS
		},
		MODE: { method: "modeSingle", ...NUMBERS },
		"MODE.SNGL": { method: "modeSingle", ...NUMBERS },
		"MODE.MULT": { method: "modeMultiple", ...NUMBERS },
		KURT: { method: "kurt", ...NUMBERS }
	}

	// The numbers of the arguments of MODE and KURT (support.microsoft.com): in a range, array or reference only
	// numbers count; typed directly, logical values and text that reads as a number count too, other text is
	// #VALUE!. Errors are raised.
	private numbers(args: readonly PluginAst[], state: PluginState): number[] | CellError {
		const out: number[] = []

		for (const arg of args) {
			const node = unwrap(arg)
			const value = this.evaluateAst(node, state)

			if (value instanceof SimpleRangeValue) {
				const error = rangeNumbers(value.data, out)

				if (error !== null) return error

				continue
			}

			if (value instanceof CellError) return value
			if (value === EmptyValue) continue

			const number = rawNumber(value)

			if (number !== undefined) {
				out.push(number)
				continue
			}

			if (node.type === ("CELL_REFERENCE" as PluginAst["type"])) continue

			if (typeof value === "boolean") out.push(value ? 1 : 0)
			else {
				const read = this.numberOf(value)

				if (read instanceof CellError) return read

				out.push(read)
			}
		}

		return out
	}

	modeSingle(ast: PluginCall, state: PluginState): unknown {
		const numbers = this.numbers(ast.args, state)

		return numbers instanceof CellError ? numbers : mode(numbers)
	}

	// MODE.MULT: every most frequent number, as a vertical array.
	modeMultiple(ast: PluginCall, state: PluginState): unknown {
		const numbers = this.numbers(ast.args, state)

		if (numbers instanceof CellError) return numbers

		const found = modes(numbers)

		return found instanceof CellError ? found : SimpleRangeValue.onlyNumbers(found.map(value => [value]))
	}

	kurt(ast: PluginCall, state: PluginState): unknown {
		const numbers = this.numbers(ast.args, state)

		return numbers instanceof CellError ? numbers : kurtosis(numbers)
	}

	private rank(ast: PluginCall, state: PluginState, name: string, average: boolean): unknown {
		return this.runFunction(ast.args, state, this.metadata(name), (number: number, ref: SimpleRangeValue, order: number) => {
			const numbers: number[] = []
			const error = rangeNumbers(ref.data, numbers)

			if (error !== null) return error

			let before = 0
			let ties = 0

			for (const value of numbers) {
				if (value === number) ties++
				else if (order === 0 ? value > number : value < number) before++
			}

			if (ties === 0) return naError()

			return average ? before + (ties + 1) / 2 : before + 1
		})
	}

	// RANK and RANK.EQ: the number's place among the numbers of ref, ties sharing the best; order 0 (or omitted)
	// descending.
	rankEq(ast: PluginCall, state: PluginState): unknown {
		return this.rank(ast, state, ast.procedureName.toUpperCase() === "RANK" ? "RANK" : "RANK.EQ", false)
	}

	// RANK.AVG: ties share the average of their places.
	rankAvg(ast: PluginCall, state: PluginState): unknown {
		return this.rank(ast, state, "RANK.AVG", true)
	}

	// FORECAST and FORECAST.LINEAR(x, known_ys, known_xs): the linear regression's value at x.
	forecast(ast: PluginCall, state: PluginState): unknown {
		const name = ast.procedureName.toUpperCase().endsWith("LINEAR") ? "FORECAST.LINEAR" : "FORECAST"

		return this.runFunction(ast.args, state, this.metadata(name), (x: number, ys: SimpleRangeValue, xs: SimpleRangeValue) => {
			const paired = pairs(ys, xs)

			if (paired instanceof CellError) return paired
			if (paired.ys.length === 0) return naError()

			const xMean = mean(paired.xs)
			const yMean = mean(paired.ys)
			let covariance = 0
			let variance = 0

			paired.xs.forEach((value, index) => {
				covariance += (value - xMean) * ((paired.ys[index] ?? 0) - yMean)
				variance += (value - xMean) ** 2
			})

			if (variance === 0) return divError()

			const slope = covariance / variance

			return yMean + slope * (x - xMean)
		})
	}

	// TREND(known_ys, [known_xs], [new_xs], [const]): the least-squares line (or plane, with several predictors) at
	// new_xs, as an array of new_xs's shape. known_xs defaults to 1, 2, 3, ...; new_xs to known_xs.
	trend(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(
			ast.args,
			state,
			this.metadata("TREND"),
			(knownYs: SimpleRangeValue, knownXs: SimpleRangeValue | undefined, newXs: SimpleRangeValue | undefined, intercept: boolean) => {
				const yRows = knownYs.data
				const ys: number[] = []
				const yError = rangeNumbers(yRows, ys)

				if (yError !== null) return yError
				if (ys.length !== yRows.length * (yRows[0]?.length ?? 0)) return valueError()

				const column = knownYs.width() === 1
				const n = ys.length
				const xRows = knownXs?.data ?? (column ? ys.map((_, index) => [index + 1]) : [ys.map((_, index) => index + 1)])
				// Observations by predictor: one predictor over the same shape as known_ys, else one per column (known_ys
				// a column) or per row (known_ys a row).
				const observations = (rows: readonly (readonly PluginScalar[])[], one: boolean): number[][] | CellError => {
					const numbers: number[] = []
					const error = rangeNumbers(rows, numbers)

					if (error !== null) return error
					if (numbers.length !== rows.length * (rows[0]?.length ?? 0)) return valueError()
					if (one) return numbers.map(value => [value])

					const byColumn = column
					const count = byColumn ? rows.length : (rows[0]?.length ?? 0)
					const width = byColumn ? (rows[0]?.length ?? 0) : rows.length

					return Array.from({ length: count }, (_, at) =>
						Array.from({ length: width }, (__, k) => rawNumber((byColumn ? rows[at]?.[k] : rows[k]?.[at]) ?? EmptyValue) ?? 0)
					)
				}
				const xCount = xRows.length * (xRows[0]?.length ?? 0)
				const single = xCount === n
				const known = observations(xRows, single)

				if (known instanceof CellError) return known
				if (known.length !== n) return naError()

				const coefficients = leastSquares(ys, known, intercept)

				if (coefficients instanceof CellError) return coefficients

				const targetRows = newXs?.data ?? xRows
				const k = known[0]?.length ?? 1
				const predict = (values: readonly number[]): number =>
					values.reduce((total, value, index) => total + value * (coefficients[index + 1] ?? 0), coefficients[0] ?? 0)

				if (k === 1) {
					const out: number[][] = []

					for (const row of targetRows) {
						const line: number[] = []

						for (const value of row) {
							const number = rawNumber(value)

							if (number === undefined) return valueError()

							line.push(predict([number]))
						}

						out.push(line)
					}

					return SimpleRangeValue.onlyNumbers(out)
				}

				const targets = observations(targetRows, false)

				if (targets instanceof CellError) return targets

				const results = targets.map(predict)

				return SimpleRangeValue.onlyNumbers(column ? results.map(value => [value]) : [results])
			}
		)
	}

	// FREQUENCY(data_array, bins_array): how many values fall at or below each bin and above the one before (bins in
	// their order of value, counts in the bins' own order), then how many are above every bin; a vertical array.
	frequency(ast: PluginCall, state: PluginState): unknown {
		return this.runFunction(ast.args, state, this.metadata("FREQUENCY"), (data: SimpleRangeValue, bins: SimpleRangeValue) => {
			const values: number[] = []
			const limits: number[] = []
			const error = rangeNumbers(data.data, values) ?? rangeNumbers(bins.data, limits)

			if (error !== null) return error

			const order = limits.map((_, index) => index).sort((a, z) => (limits[a] ?? 0) - (limits[z] ?? 0))
			const counts = new Array<number>(limits.length + 1).fill(0)

			for (const value of values) {
				let low = 0
				let high = order.length

				// The first bin, in value order, at or above the value.
				while (low < high) {
					const mid = (low + high) >> 1

					if ((limits[order[mid] ?? 0] ?? 0) >= value) high = mid
					else low = mid + 1
				}

				const target = low < order.length ? (order[low] ?? limits.length) : limits.length

				counts[target] = (counts[target] ?? 0) + 1
			}

			return SimpleRangeValue.onlyNumbers(counts.map(count => [count]))
		})
	}
}
