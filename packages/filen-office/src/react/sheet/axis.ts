import { AxisFlag, type SheetAxisLayout } from "../../core/sheet/protocol"

// One axis of the grid as drawn: where each row or column starts and how large it is, in screen pixels at the
// sheet's zoom. Only the indices that differ from the default are stored (sorted, with the running difference
// from the default before each), so an axis of a million rows costs memory for its exceptions alone and every
// lookup is a binary search over them.

export interface GridAxis {
	readonly count: number
	readonly total: number
	readonly defaultSize: number
	size(index: number): number
	// Where `index` starts (clamped to 0..count).
	offset(index: number): number
	// The index whose span holds `position`, clamped to the axis.
	indexAt(position: number): number
	hidden(index: number): boolean
	// The last index with a size (-1: none), so a run of hidden indices at the end is never walked.
	readonly lastShown: number
}

// Pixels at 100 % (a resize being dragged); NaN takes an index back to the default.
export type AxisOverrides = ReadonlyMap<number, number>

export function createGridAxis(layout: SheetAxisLayout, count: number, scale: number, overrides?: AxisOverrides): GridAxis {
	const defaultSize = Math.max(1, Math.round(layout.defaultSize * scale))
	let indices: Uint32Array
	let sizes: Float64Array

	if (overrides === undefined || overrides.size === 0) {
		let n = 0

		while (n < layout.indices.length && (layout.indices[n] ?? 0) < count) n++

		indices = layout.indices.slice(0, n)
		sizes = new Float64Array(n)

		for (let k = 0; k < n; k++) sizes[k] = scaled(layout, k, scale)
	} else {
		const merged = new Map<number, number>()

		for (let k = 0; k < layout.indices.length; k++) {
			const index = layout.indices[k] ?? 0

			if (index < count) merged.set(index, scaled(layout, k, scale))
		}

		for (const [index, px] of overrides) {
			if (index >= count) continue
			if (Number.isNaN(px)) merged.delete(index)
			else merged.set(index, Math.round(px * scale))
		}

		indices = Uint32Array.from(merged.keys()).sort()
		sizes = Float64Array.from(indices, index => merged.get(index) ?? defaultSize)
	}

	const n = indices.length
	// before[k]: the summed difference from the default of the exceptions before k.
	const before = new Float64Array(n + 1)

	for (let k = 0; k < n; k++) before[k + 1] = (before[k] ?? 0) + (sizes[k] ?? 0) - defaultSize

	// The first exception at or after `index`.
	function lowerBound(index: number): number {
		let low = 0
		let high = n

		while (low < high) {
			const middle = (low + high) >>> 1

			if ((indices[middle] ?? 0) < index) low = middle + 1
			else high = middle
		}

		return low
	}

	function offset(index: number): number {
		const clamped = Math.max(0, Math.min(index, count))

		return clamped * defaultSize + (before[lowerBound(clamped)] ?? 0)
	}

	function size(index: number): number {
		const k = lowerBound(index)

		return k < n && indices[k] === index ? (sizes[k] ?? 0) : defaultSize
	}

	const total = offset(count)

	function indexAt(position: number): number {
		if (count === 0) return 0
		if (position >= total) return count - 1
		if (position < 0) return 0

		// The last exception starting at or before `position`.
		let low = 0
		let high = n

		while (low < high) {
			const middle = (low + high) >>> 1
			const start = (indices[middle] ?? 0) * defaultSize + (before[middle] ?? 0)

			if (start <= position) low = middle + 1
			else high = middle
		}

		const k = low - 1

		if (k < 0) return Math.min(count - 1, Math.floor(position / defaultSize))

		const index = indices[k] ?? 0
		const start = index * defaultSize + (before[k] ?? 0)
		const end = start + (sizes[k] ?? 0)

		if (position < end) return index

		const next = k + 1 < n ? (indices[k + 1] ?? count) : count

		return Math.min(count - 1, next - 1, index + 1 + Math.floor((position - end) / defaultSize))
	}

	// Walks back over the hidden indices at the end, which are all exceptions.
	let lastShown = count - 1
	let k = lowerBound(count) - 1

	while (lastShown >= 0 && k >= 0 && indices[k] === lastShown && sizes[k] === 0) {
		lastShown--
		k--
	}

	return { count, total, defaultSize, size, offset, indexAt, hidden: index => size(index) === 0, lastShown }
}

function scaled(layout: SheetAxisLayout, k: number, scale: number): number {
	return ((layout.flags[k] ?? 0) & AxisFlag.HIDDEN) !== 0 ? 0 : Math.round((layout.sizes[k] ?? layout.defaultSize) * scale)
}

// The axis's style for the cells of `index` that have no tile entry, and its flags.
export function axisEntry(layout: SheetAxisLayout, index: number): { readonly style: number; readonly flags: number } | null {
	let low = 0
	let high = layout.indices.length

	while (low < high) {
		const middle = (low + high) >>> 1

		if ((layout.indices[middle] ?? 0) < index) low = middle + 1
		else high = middle
	}

	return layout.indices[low] === index ? { style: layout.styles[low] ?? 0, flags: layout.flags[low] ?? 0 } : null
}

// The next shown index from `from` (inclusive) in `direction`, or -1.
export function shownIndex(axis: GridAxis, from: number, direction: 1 | -1): number {
	const end = direction === 1 ? axis.lastShown : 0

	for (let index = Math.min(from, axis.count - 1); direction === 1 ? index <= end : index >= end; index += direction)
		if (!axis.hidden(index)) return index

	return -1
}
