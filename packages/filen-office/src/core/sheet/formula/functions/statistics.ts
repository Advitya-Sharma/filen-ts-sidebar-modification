import type { CellError } from "hyperformula"
import { divError, naError, numError } from "./plugin"

// Statistics over plain numbers, as the Excel functions of the same names define them (support.microsoft.com).

export function mean(numbers: readonly number[]): number {
	let total = 0

	for (const number of numbers) total += number

	return total / numbers.length
}

// The sum of squared deviations from the mean, in two passes (exact for constant data).
function squares(numbers: readonly number[]): number {
	const average = mean(numbers)
	let total = 0

	for (const number of numbers) total += (number - average) ** 2

	return total
}

// VAR.S: n - 1 in the denominator.
export function sampleVariance(numbers: readonly number[]): number | CellError {
	return numbers.length < 2 ? divError() : squares(numbers) / (numbers.length - 1)
}

// VAR.P.
export function populationVariance(numbers: readonly number[]): number | CellError {
	return numbers.length === 0 ? divError() : squares(numbers) / numbers.length
}

// MODE.SNGL: the value that repeats most, the first of them in order on a tie; #N/A when none repeats.
export function mode(numbers: readonly number[]): number | CellError {
	const counts = new Map<number, number>()
	let best = 0
	let value = 0

	for (const number of numbers) {
		const count = (counts.get(number) ?? 0) + 1

		counts.set(number, count)

		if (count > best) {
			best = count
			value = number
		}
	}

	if (best < 2) return naError()

	// The first value in order among those that reach the highest count.
	for (const number of numbers) if (counts.get(number) === best) return number

	return value
}

// MODE.MULT: every value that repeats most, in the order each first appears.
export function modes(numbers: readonly number[]): number[] | CellError {
	const counts = new Map<number, number>()
	let best = 0

	for (const number of numbers) {
		const count = (counts.get(number) ?? 0) + 1

		counts.set(number, count)
		best = Math.max(best, count)
	}

	if (best < 2) return naError()

	const seen = new Set<number>()
	const out: number[] = []

	for (const number of numbers) {
		if (counts.get(number) !== best || seen.has(number)) continue

		seen.add(number)
		out.push(number)
	}

	return out
}

// LARGE (largest true) and SMALL: the k-th value, k from 1.
export function kth(numbers: readonly number[], k: number | CellError, largest: boolean): number | CellError {
	if (typeof k !== "number") return k

	const index = Math.ceil(k)

	if (numbers.length === 0 || index < 1 || index > numbers.length) return numError()

	const sorted = [...numbers].sort((a, z) => (largest ? z - a : a - z))

	return sorted[index - 1] ?? numError()
}

// PERCENTILE.INC over sorted numbers, k in [0, 1].
export function percentileInclusive(sorted: readonly number[], k: number): number {
	const position = k * (sorted.length - 1)
	const low = Math.floor(position)
	const lower = sorted[low] ?? 0
	const upper = sorted[Math.min(low + 1, sorted.length - 1)] ?? lower

	return lower + (position - low) * (upper - lower)
}

// PERCENTILE.EXC over sorted numbers, k in (0, 1): #NUM! when k falls outside 1/(n+1) .. n/(n+1).
export function percentileExclusive(sorted: readonly number[], k: number): number | CellError {
	const position = k * (sorted.length + 1)

	if (position < 1 || position > sorted.length) return numError()

	const low = Math.floor(position)
	const lower = sorted[low - 1] ?? 0
	const upper = sorted[Math.min(low, sorted.length - 1)] ?? lower

	return lower + (position - low) * (upper - lower)
}

// KURT: the sample excess kurtosis; at least four values and a deviation that is not 0.
export function kurtosis(numbers: readonly number[]): number | CellError {
	const n = numbers.length

	if (n < 4) return divError()

	const average = mean(numbers)
	let second = 0
	let fourth = 0

	for (const number of numbers) {
		const d = (number - average) ** 2

		second += d
		fourth += d * d
	}

	const variance = second / (n - 1)

	if (variance === 0) return divError()

	return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * (fourth / variance ** 2) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
}
