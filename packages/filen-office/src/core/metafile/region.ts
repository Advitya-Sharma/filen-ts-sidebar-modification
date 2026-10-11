import { MAX_REGION_RECTS, MetafileLimit } from "./limits"

// Exact boolean algebra on regions made of axis-aligned rectangles (GDI regions: RGNDATA, clip rectangles
// under an axis-aligned mapping). A region is a Float64Array of rectangles (x0, y0, x1, y1) in y-x bands:
// disjoint, sorted by top then left, rectangles of one band sharing top and bottom.

export type RectOp = "and" | "or" | "xor" | "diff"

// Work units a caller is charged for (it throws when its budget runs out).
export type Charge = (units: number) => void

interface Band {
	y0: number
	y1: number
	xs: number[]
}

function interval(op: RectOp, a: boolean, b: boolean): boolean {
	switch (op) {
		case "and":
			return a && b
		case "or":
			return a || b
		case "xor":
			return a !== b
		case "diff":
			return a && !b
	}
}

// Sorted, merged x intervals (x0, x1 pairs) of the rectangles in `active` that cover [y0, y1).
function covering(rects: ArrayLike<number>, active: readonly number[], y0: number, y1: number, charge: Charge): number[] {
	const spans: [number, number][] = []

	for (const index of active) {
		const r = index * 4

		if ((rects[r + 1] ?? 0) <= y0 && (rects[r + 3] ?? 0) >= y1) spans.push([rects[r] ?? 0, rects[r + 2] ?? 0])
	}

	charge(active.length + spans.length)
	spans.sort((p, q) => p[0] - q[0])

	const out: number[] = []

	for (const [x0, x1] of spans) {
		if (x1 <= x0) continue

		const last = out.length - 1

		if (last > 0 && x0 <= (out[last] ?? 0)) out[last] = Math.max(out[last] ?? 0, x1)
		else out.push(x0, x1)
	}

	return out
}

// Boolean combination of two sorted interval lists (alternating starts and ends), in one merge pass.
function combineIntervals(a: readonly number[], b: readonly number[], op: RectOp): number[] {
	const out: number[] = []
	let ia = 0
	let ib = 0
	let inA = false
	let inB = false
	let previous = 0

	while (ia < a.length || ib < b.length) {
		const x = Math.min(ia < a.length ? (a[ia] ?? 0) : Infinity, ib < b.length ? (b[ib] ?? 0) : Infinity)

		if (x > previous && interval(op, inA, inB)) {
			const last = out.length - 1

			if (last > 0 && out[last] === previous) out[last] = x
			else out.push(previous, x)
		}

		while (ia < a.length && a[ia] === x) {
			inA = !inA
			ia++
		}

		while (ib < b.length && b[ib] === x) {
			inB = !inB
			ib++
		}

		previous = x
	}

	return out
}

// Sweeps the rectangles' horizontal edges, combining what covers each band.
function sweep(a: ArrayLike<number>, b: ArrayLike<number>, op: RectOp, charge: Charge): Float64Array {
	const countA = Math.floor(a.length / 4)
	const countB = Math.floor(b.length / 4)

	charge(countA + countB)

	const ys: number[] = []

	for (let k = 0; k < countA; k++) ys.push(a[k * 4 + 1] ?? 0, a[k * 4 + 3] ?? 0)
	for (let k = 0; k < countB; k++) ys.push(b[k * 4 + 1] ?? 0, b[k * 4 + 3] ?? 0)

	ys.sort((p, q) => p - q)

	const orderA = Array.from({ length: countA }, (_, k) => k).sort((p, q) => (a[p * 4 + 1] ?? 0) - (a[q * 4 + 1] ?? 0))
	const orderB = Array.from({ length: countB }, (_, k) => k).sort((p, q) => (b[p * 4 + 1] ?? 0) - (b[q * 4 + 1] ?? 0))
	let nextA = 0
	let nextB = 0
	let activeA: number[] = []
	let activeB: number[] = []
	const bands: Band[] = []
	let rectCount = 0

	for (let k = 0; k + 1 < ys.length; k++) {
		const y0 = ys[k] ?? 0
		const y1 = ys[k + 1] ?? 0

		if (y1 <= y0) continue

		while (nextA < countA && (a[(orderA[nextA] ?? 0) * 4 + 1] ?? 0) <= y0) activeA.push(orderA[nextA++] ?? 0)
		while (nextB < countB && (b[(orderB[nextB] ?? 0) * 4 + 1] ?? 0) <= y0) activeB.push(orderB[nextB++] ?? 0)

		activeA = activeA.filter(index => (a[index * 4 + 3] ?? 0) > y0)
		activeB = activeB.filter(index => (b[index * 4 + 3] ?? 0) > y0)

		const xs = combineIntervals(covering(a, activeA, y0, y1, charge), covering(b, activeB, y0, y1, charge), op)

		if (xs.length === 0) continue

		const previous = bands[bands.length - 1]

		// A band equal to the one above it extends that one.
		if (previous !== undefined && previous.y1 === y0 && previous.xs.length === xs.length && previous.xs.every((x, i) => x === xs[i])) {
			previous.y1 = y1
			continue
		}

		rectCount += xs.length / 2

		if (rectCount > MAX_REGION_RECTS) throw new MetafileLimit("region rectangles")

		bands.push({ y0, y1, xs })
	}

	const out = new Float64Array(rectCount * 4)
	let at = 0

	for (const band of bands) {
		for (let k = 0; k + 1 < band.xs.length; k += 2) {
			out[at++] = band.xs[k] ?? 0
			out[at++] = band.y0
			out[at++] = band.xs[k + 1] ?? 0
			out[at++] = band.y1
		}
	}

	return out
}

// Any rectangles (corners in any order) as a banded region.
export function normalizeRects(rects: ArrayLike<number>, charge: Charge): Float64Array {
	if (rects.length === 4) {
		// One rectangle (the common clip rectangle) needs no sweep.
		charge(1)

		const x0 = Math.min(rects[0] ?? 0, rects[2] ?? 0)
		const x1 = Math.max(rects[0] ?? 0, rects[2] ?? 0)
		const y0 = Math.min(rects[1] ?? 0, rects[3] ?? 0)
		const y1 = Math.max(rects[1] ?? 0, rects[3] ?? 0)

		return x1 > x0 && y1 > y0 && Number.isFinite(x0 + x1 + y0 + y1) ? Float64Array.of(x0, y0, x1, y1) : new Float64Array(0)
	}

	const fixed: number[] = []

	for (let k = 0; k + 3 < rects.length && fixed.length < MAX_REGION_RECTS * 4; k += 4) {
		const x0 = Math.min(rects[k] ?? 0, rects[k + 2] ?? 0)
		const x1 = Math.max(rects[k] ?? 0, rects[k + 2] ?? 0)
		const y0 = Math.min(rects[k + 1] ?? 0, rects[k + 3] ?? 0)
		const y1 = Math.max(rects[k + 1] ?? 0, rects[k + 3] ?? 0)

		if (x1 > x0 && y1 > y0 && Number.isFinite(x0 + x1 + y0 + y1)) fixed.push(x0, y0, x1, y1)
	}

	return sweep(fixed, [], "or", charge)
}

export function combineRects(a: Float64Array, b: Float64Array, op: RectOp, charge: Charge): Float64Array {
	if (op === "and" && a.length === 4 && b.length === 4) {
		charge(1)

		const x0 = Math.max(a[0] ?? 0, b[0] ?? 0)
		const y0 = Math.max(a[1] ?? 0, b[1] ?? 0)
		const x1 = Math.min(a[2] ?? 0, b[2] ?? 0)
		const y1 = Math.min(a[3] ?? 0, b[3] ?? 0)

		return x1 > x0 && y1 > y0 ? Float64Array.of(x0, y0, x1, y1) : new Float64Array(0)
	}

	if (op === "and" && (a.length === 0 || b.length === 0)) return new Float64Array(0)

	return sweep(a, b, op, charge)
}

export function translateRects(rects: Float64Array, dx: number, dy: number, charge: Charge): Float64Array {
	charge(rects.length)

	const out = new Float64Array(rects.length)

	for (let k = 0; k < rects.length; k += 2) {
		out[k] = (rects[k] ?? 0) + dx
		out[k + 1] = (rects[k + 1] ?? 0) + dy
	}

	return out
}
