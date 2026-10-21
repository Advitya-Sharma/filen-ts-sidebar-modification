import type { SheetId, SheetRange } from "../protocol"
import type { FormulaHost } from "./formula/evaluate"
import { FormulaError, type Scalar } from "./formula/values"
import type { RangeIndex } from "./ranges"

// What the rules over one set of ranges compute from its cells (top and bottom ranks, the average and standard
// deviation, how often each value occurs, the scales' ends), read a band of rows at a time and kept per band: an
// edit marks the bands it touches stale and only those are read again, so a keystroke in a million-row range reads a
// few thousand cells, and an edit made while the bands are read ahead is never lost (its band is read again). A
// stale band keeps counting until it is read again: what the bands give meanwhile is the last known statistics.

// A value as the unique and duplicate rules tell values apart: text without case, blank never counted.
export function valueKey(value: Scalar): string | null {
	if (value === null || value === "") return null
	if (typeof value === "number") return `n${String(value)}`
	if (typeof value === "boolean") return `b${String(value)}`
	if (value instanceof FormulaError) return `e${value.code}`

	return `s${value.toUpperCase()}`
}

interface Band {
	// The band's numbers, ascending.
	readonly numbers: Float64Array
	readonly sum: number
	// The sum of squared deviations from the band's mean.
	readonly m2: number
	// How often each other value occurs (unique and duplicate rules only).
	readonly texts: Map<string, number> | null
}

// About this many cells of the ranges' width per band, and at most this many bands over its rows.
const BAND_CELLS = 16_384
const MAX_BAND_ROWS = 4_096
const MIN_BAND_ROWS = 64
const MAX_BANDS = 512
// A band meeting more parts than this is read through a mask of them, up to this many cells.
const MASK_PARTS = 32
const MAX_MASK_CELLS = 1_048_576

function bandRows(bounds: SheetRange): number {
	const width = bounds.endCol - bounds.startCol + 1
	const span = bounds.endRow - bounds.startRow + 1
	let rows = MAX_BAND_ROWS

	while (rows > MIN_BAND_ROWS && rows * width > BAND_CELLS) rows >>= 1
	while (Math.ceil(span / rows) > MAX_BANDS) rows *= 2

	return rows
}

function lowerBound(numbers: Float64Array, value: number, low: number, high: number): number {
	let lo = low
	let hi = high

	while (lo < hi) {
		const middle = (lo + hi) >> 1

		if ((numbers[middle] ?? 0) < value) lo = middle + 1
		else hi = middle
	}

	return lo
}

function upperBound(numbers: Float64Array, value: number, low: number, high: number): number {
	let lo = low
	let hi = high

	while (lo < hi) {
		const middle = (lo + hi) >> 1

		if ((numbers[middle] ?? 0) <= value) lo = middle + 1
		else hi = middle
	}

	return lo
}

// The k-th smallest (from 0) of ascending arrays, without merging them: each turn splits at the weighted median of
// the remaining parts' middles, which leaves out at least a quarter of what remains.
function kth(arrays: readonly Float64Array[], k: number): number {
	if (arrays.length === 1) return (arrays[0] as Float64Array)[k] ?? Number.NaN

	const low = new Int32Array(arrays.length)
	const high = Int32Array.from(arrays, array => array.length)
	const lows = new Int32Array(arrays.length)
	const highs = new Int32Array(arrays.length)
	const middles: { value: number; weight: number }[] = []

	for (;;) {
		middles.length = 0

		let remaining = 0

		arrays.forEach((array, i) => {
			const size = (high[i] ?? 0) - (low[i] ?? 0)

			if (size > 0) {
				middles.push({ value: array[((low[i] ?? 0) + (high[i] ?? 0)) >> 1] ?? 0, weight: size })
				remaining += size
			}
		})

		if (remaining === 0) return Number.NaN

		middles.sort((a, b) => a.value - b.value)

		let pivot = middles[middles.length - 1]?.value ?? 0
		let seen = 0

		for (const middle of middles) {
			seen += middle.weight

			if (seen * 2 >= remaining) {
				pivot = middle.value
				break
			}
		}

		let less = 0
		let lessOrEqual = 0

		arrays.forEach((array, i) => {
			lows[i] = lowerBound(array, pivot, low[i] ?? 0, high[i] ?? 0)
			highs[i] = upperBound(array, pivot, low[i] ?? 0, high[i] ?? 0)
			less += lows[i] ?? 0
			lessOrEqual += highs[i] ?? 0
		})

		if (k < less) high.set(lows)
		else if (k >= lessOrEqual) low.set(highs)
		else return pivot
	}
}

// What the bands give together.
export class RangeTotals {
	readonly count: number
	readonly average: number
	// STDEV.P, as LibreOffice computes "above average by n standard deviations".
	readonly deviation: number
	readonly min: number
	readonly max: number
	private readonly arrays: readonly Float64Array[]
	private readonly order = new Map<number, number>()

	constructor(bands: Iterable<Band>) {
		const arrays: Float64Array[] = []
		let count = 0
		let sum = 0
		let mean = 0
		let m2 = 0
		let min = Number.NaN
		let max = Number.NaN

		for (const band of bands) {
			const size = band.numbers.length

			if (size === 0) continue

			arrays.push(band.numbers)

			// Chan et al.'s pairwise update of the mean and squared deviations.
			const bandMean = band.sum / size
			const total = count + size
			const delta = bandMean - mean

			m2 += band.m2 + (delta * delta * count * size) / total
			mean += (delta * size) / total
			count = total
			sum += band.sum

			const first = band.numbers[0] ?? 0
			const last = band.numbers[size - 1] ?? 0

			if (Number.isNaN(min) || first < min) min = first
			if (Number.isNaN(max) || last > max) max = last
		}

		this.arrays = arrays
		this.count = count
		this.average = count === 0 ? Number.NaN : sum / count
		this.deviation = count === 0 ? Number.NaN : Math.sqrt(m2 / count)
		this.min = min
		this.max = max
	}

	// The k-th smallest number (from 0).
	nth(k: number): number {
		if (k < 0 || k >= this.count) return Number.NaN

		let value = this.order.get(k)

		if (value === undefined) {
			value = kth(this.arrays, k)
			this.order.set(k, value)
		}

		return value
	}

	// PERCENTILE.INC.
	percentile(p: number): number {
		if (this.count === 0) return Number.NaN

		const rank = Math.max(0, Math.min(1, p)) * (this.count - 1)
		const low = Math.floor(rank)
		const high = Math.min(this.count - 1, low + 1)
		const lowValue = this.nth(low)

		return lowValue + (rank - low) * (this.nth(high) - lowValue)
	}
}

export class RangeStatistics {
	private readonly sheet: SheetId
	private readonly host: FormulaHost
	private readonly index: RangeIndex
	// Unique and duplicate rules read how often each value occurs.
	private readonly counting: boolean
	private readonly height: number
	private readonly firstBand: number
	private readonly overlaps: boolean
	private readonly bands = new Map<number, Band>()
	// Bands read before the cells they hold changed.
	private readonly stale = new Set<number>()
	// Bands read in turns (steps), taken in together once all are read: what the bands give meanwhile stays as it
	// was, never half old and half new.
	private readonly staged = new Map<number, Band>()
	// How the staged bands change the counts (unique and duplicate rules), taken in with them: a structure edit's
	// bands, read again as they were, change them by little.
	private readonly numberDelta = new Map<number, number>()
	private readonly textDelta = new Map<string, number>()
	// No band before this one is missing.
	private cursor: number
	private totals: RangeTotals | null = null
	private readonly numberCounts = new Map<number, number>()
	private readonly textCounts = new Map<string, number>()
	// What these ranges gave before a structure edit rewrote them (statistics.ts of the formats replaced), until
	// these are read.
	private fallback: RangeStatistics | null

	constructor(sheet: SheetId, host: FormulaHost, index: RangeIndex, counting: boolean, fallback: RangeStatistics | null = null) {
		this.sheet = sheet
		this.host = host
		this.index = index
		this.counting = counting
		this.fallback = fallback?.lastKnown() ?? null
		this.height = bandRows(index.bounds)
		this.firstBand = Math.floor(index.bounds.startRow / this.height)
		this.cursor = this.firstBand
		this.overlaps = index.ranges.some((range, i) => {
			let earlier = false

			index.each(range, (_part, j) => {
				if (j < i) earlier = true
			})

			return earlier
		})
	}

	// The last band holding cells (the sheet's used rows).
	private lastBand(): number {
		const last = Math.min(this.index.bounds.endRow, this.host.extent(this.sheet).rows - 1)

		return last < this.index.bounds.startRow ? this.firstBand - 1 : Math.floor(last / this.height)
	}

	private next(): number {
		const last = this.lastBand()

		let lowest = -1

		for (const band of [...this.stale]) {
			// Past the used rows now: nothing to read there.
			if (band > last) this.drop(band)
			else if (!this.staged.has(band) && (lowest < 0 || band < lowest)) lowest = band
		}

		if (lowest >= 0) return lowest

		while (this.cursor <= last) {
			if (!this.bands.has(this.cursor) && !this.staged.has(this.cursor)) return this.cursor

			this.cursor++
		}

		return -1
	}

	// Every band read and taken in.
	get complete(): boolean {
		return this.next() < 0 && this.staged.size === 0
	}

	// How many bands wait to be read, counted up to `limit` + 1.
	pending(limit: number): number {
		const last = this.lastBand()
		let count = 0

		for (const band of this.stale) if (band <= last && !this.staged.has(band) && ++count > limit) return count

		for (let band = this.cursor; band <= last; band++)
			if (!this.bands.has(band) && !this.staged.has(band) && ++count > limit) return count

		return count
	}

	// The statistics as last known: the bands held (stale ones as they were read), else what the ranges gave before
	// a structure edit; null when nothing is known yet.
	known(): RangeTotals | null {
		const source = this.lastKnown()

		if (source === null) return null

		source.totals ??= new RangeTotals(source.bands.values())

		return source.totals
	}

	private lastKnown(): RangeStatistics | null {
		return this.bands.size > 0 ? this : (this.fallback?.lastKnown() ?? null)
	}

	// Reads the bands left to read, one per step, and takes them in at the end.
	*steps(): Generator<void, void> {
		for (let band = this.next(); band >= 0; band = this.next()) {
			this.stage(band, this.scan(band))

			yield
		}

		this.takeStaged()
	}

	ensure(): RangeTotals {
		this.takeStaged()

		for (let band = this.next(); band >= 0; band = this.next()) this.take(band, this.scan(band))

		if (this.fallback !== null) this.fallback = null

		this.totals ??= new RangeTotals(this.bands.values())

		return this.totals
	}

	// How often a value occurs in the ranges (blank never counts), as last known with `known`, else read now.
	occurrences(value: Scalar, known = false): number {
		const source = known ? (this.lastKnown() ?? this) : this

		if (!known) this.ensure()

		if (typeof value === "number") return source.numberCounts.get(value) ?? 0

		const key = valueKey(value)

		return key === null ? 0 : (source.textCounts.get(key) ?? 0)
	}

	// Marks the bands `range` changed stale; false when it misses the ranges.
	invalidate(range: SheetRange): boolean {
		if (!this.index.any(range)) return false

		const first = Math.max(this.firstBand, Math.floor(range.startRow / this.height))
		const last = Math.floor(Math.min(range.endRow, this.index.bounds.endRow) / this.height)

		if (last - first + 1 > this.bands.size) {
			for (const band of this.bands.keys()) if (band >= first && band <= last) this.stale.add(band)
		} else {
			for (let band = first; band <= last; band++) if (this.bands.has(band)) this.stale.add(band)
		}

		// A band read in turns before this change is read again.
		for (const band of [...this.staged.keys()]) if (band >= first && band <= last) this.unstage(band)

		this.cursor = Math.min(this.cursor, first)
		this.totals = null

		return true
	}

	private takeStaged(): void {
		if (this.staged.size === 0) return

		for (const [value, delta] of this.numberDelta) adjust(this.numberCounts, value, delta, false)
		for (const [key, delta] of this.textDelta) adjust(this.textCounts, key, delta, false)

		this.numberDelta.clear()
		this.textDelta.clear()

		for (const [band, data] of this.staged) {
			this.stale.delete(band)
			this.bands.set(band, data)
		}

		this.staged.clear()
		this.totals = null

		if (this.fallback !== null && this.next() < 0) this.fallback = null
	}

	private scan(band: number): Band {
		const startRow = band * this.height
		const endRow = startRow + this.height - 1
		const bounds = this.index.bounds
		const parts: [SheetRange, number][] = []

		this.index.each({ startRow, startCol: bounds.startCol, endRow, endCol: bounds.endCol }, (part, index) => parts.push([part, index]))

		const numbers: number[] = []
		const texts = this.counting ? new Map<string, number>() : null
		const visit = (value: Scalar): void => {
			if (typeof value === "number") {
				numbers.push(value)
			} else if (texts !== null) {
				const key = valueKey(value)

				if (key !== null) texts.set(key, (texts.get(key) ?? 0) + 1)
			}
		}
		const top = Math.max(startRow, bounds.startRow)
		const bottom = Math.min(endRow, bounds.endRow)
		const width = bounds.endCol - bounds.startCol + 1

		if (parts.length > MASK_PARTS && (bottom - top + 1) * width <= MAX_MASK_CELLS) {
			// Many small parts (a fragmented sqref): the band's cells read at once, through a mask of what they cover
			// (a cell in two parts once).
			const mask = new Uint8Array((bottom - top + 1) * width)

			for (const [part] of parts) {
				for (let row = Math.max(part.startRow, top); row <= Math.min(part.endRow, bottom); row++) {
					const offset = (row - top) * width - bounds.startCol

					mask.fill(1, offset + part.startCol, offset + part.endCol + 1)
				}
			}

			this.host.forEach(
				{ kind: "ref", sheet: this.sheet, startRow: top, startCol: bounds.startCol, endRow: bottom, endCol: bounds.endCol },
				(row, col, value) => {
					if (mask[(row - top) * width + col - bounds.startCol] === 1) visit(value)
				}
			)
		} else {
			for (const [part, index] of parts) {
				this.host.forEach(
					{
						kind: "ref",
						sheet: this.sheet,
						startRow: Math.max(part.startRow, startRow),
						startCol: part.startCol,
						endRow: Math.min(part.endRow, endRow),
						endCol: part.endCol
					},
					(row, col, value) => {
						// A cell in two of the parts counts once.
						if (!this.overlaps || this.index.first(row, col) === index) visit(value)
					}
				)
			}
		}

		const sorted = Float64Array.from(numbers).sort()
		let sum = 0

		for (const value of sorted) sum += value

		const mean = sorted.length === 0 ? 0 : sum / sorted.length
		let m2 = 0

		for (const value of sorted) m2 += (value - mean) ** 2

		return { numbers: sorted, sum, m2, texts }
	}

	// A band as read, in place of what it counted before.
	private take(band: number, data: Band): void {
		this.drop(band)
		this.bands.set(band, data)
		this.totals = null

		if (this.counting) countBand(this.numberCounts, this.textCounts, data, 1, false)
	}

	private stage(band: number, data: Band): void {
		this.staged.set(band, data)

		if (!this.counting) return

		const held = this.bands.get(band)

		if (held !== undefined) countBand(this.numberDelta, this.textDelta, held, -1, true)

		countBand(this.numberDelta, this.textDelta, data, 1, true)
	}

	private unstage(band: number): void {
		const data = this.staged.get(band)

		if (data === undefined) return

		this.staged.delete(band)

		if (!this.counting) return

		const held = this.bands.get(band)

		countBand(this.numberDelta, this.textDelta, data, -1, true)

		if (held !== undefined) countBand(this.numberDelta, this.textDelta, held, 1, true)
	}

	private drop(band: number): void {
		this.unstage(band)
		this.stale.delete(band)

		const held = this.bands.get(band)

		if (held === undefined) return

		this.bands.delete(band)
		this.totals = null

		if (this.counting) countBand(this.numberCounts, this.textCounts, held, -1, false)
	}
}

// Adds `delta` to a count; a count that falls to zero (below it too, unless `signed`: a change of counts) goes.
function adjust<K>(counts: Map<K, number>, key: K, delta: number, signed: boolean): void {
	const next = (counts.get(key) ?? 0) + delta

	if (next === 0 || (!signed && next < 0)) counts.delete(key)
	else counts.set(key, next)
}

function countBand(numbers: Map<number, number>, texts: Map<string, number>, band: Band, sign: number, signed: boolean): void {
	for (let i = 0; i < band.numbers.length; i++) adjust(numbers, band.numbers[i] ?? 0, sign, signed)

	if (band.texts !== null) for (const [key, count] of band.texts) adjust(texts, key, sign * count, signed)
}
