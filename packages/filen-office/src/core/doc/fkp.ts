import type { Bytes } from "./bytes"
import { corrupt } from "./errors"

// [MS-DOC] §2.8.5 PlcBteChpx / §2.8.6 PlcBtePapx → §2.9.33 ChpxFkp / §2.9.174 PapxFkp (512-byte pages in the
// WordDocument stream), flattened into one sorted list of fc ranges per kind.

const PAGE = 512

// fc ranges [fcStart, fcEnd) in WordDocument bytes, each with the location of its grpprl (Chpx) or GrpPrlAndIstd
// (PapxInFkp) in the WordDocument stream; length 0 means no properties (Chpx offset 0).
export interface FkpRuns {
	readonly fcStart: Uint32Array
	readonly fcEnd: Uint32Array
	readonly at: Uint32Array
	readonly length: Uint16Array
	readonly count: number
}

interface Bte {
	readonly pns: number[]
}

function readBte(table: Bytes, fc: number, lcb: number, word: Bytes): Bte {
	if (lcb === 0) return { pns: [] }

	if (lcb < 4 || (lcb - 4) % 8 !== 0 || !table.has(fc, lcb)) throw corrupt("PlcBte size is invalid")

	const count = (lcb - 4) / 8
	const pnAt = fc + 4 * (count + 1)
	const seen = new Set<number>()
	const pns: number[] = []
	const pages = Math.floor(word.length / PAGE)

	for (let k = 0; k < count; k++) {
		const pn = table.u32(pnAt + 4 * k) & 0x3fffff

		// A page used twice adds nothing; a page past the stream is skipped (its text keeps default properties).
		if (pn >= pages || seen.has(pn)) continue

		seen.add(pn)
		pns.push(pn)
	}

	return { pns }
}

class RunBuilder {
	fcStart: number[] = []
	fcEnd: number[] = []
	at: number[] = []
	length: number[] = []

	push(start: number, end: number, at: number, length: number): void {
		if (end <= start) return

		this.fcStart.push(start)
		this.fcEnd.push(end)
		this.at.push(at)
		this.length.push(length)
	}

	build(): FkpRuns {
		const count = this.fcStart.length
		const order = Array.from({ length: count }, (_, k) => k)
		let sorted = true

		for (let k = 1; k < count && sorted; k++) sorted = (this.fcStart[k] ?? 0) >= (this.fcStart[k - 1] ?? 0)

		if (!sorted) order.sort((a, b) => (this.fcStart[a] ?? 0) - (this.fcStart[b] ?? 0))

		const fcStart = new Uint32Array(count)
		const fcEnd = new Uint32Array(count)
		const at = new Uint32Array(count)
		const length = new Uint16Array(count)
		let kept = 0
		let reached = 0

		// Overlapping ranges (only in damaged files) are clipped so the runs stay disjoint and every lookup linear.
		for (const source of order) {
			const start = Math.max(this.fcStart[source] ?? 0, reached)
			const end = this.fcEnd[source] ?? 0

			if (end <= start) continue

			fcStart[kept] = start
			fcEnd[kept] = end
			at[kept] = this.at[source] ?? 0
			length[kept] = this.length[source] ?? 0
			reached = end
			kept++
		}

		return { fcStart, fcEnd, at, length, count: kept }
	}
}

export function readChpxRuns(word: Bytes, table: Bytes, fc: number, lcb: number): FkpRuns {
	const builder = new RunBuilder()

	for (const pn of readBte(table, fc, lcb, word).pns) {
		const page = pn * PAGE
		const crun = word.u8(page + PAGE - 1)

		// crun is 1..0x65 (§2.9.33); rgfc and rgb must fit before the count byte.
		if (crun === 0 || 4 * (crun + 1) + crun > PAGE - 1) continue

		for (let k = 0; k < crun; k++) {
			const start = word.u32(page + 4 * k)
			const end = word.u32(page + 4 * (k + 1))
			const offset = word.u8(page + 4 * (crun + 1) + k) * 2

			if (offset === 0) {
				builder.push(start, end, 0, 0)
				continue
			}

			const cb = word.u8(page + offset)

			builder.push(start, end, page + offset + 1, offset + 1 + cb <= PAGE - 1 ? cb : 0)
		}
	}

	return builder.build()
}

export function readPapxRuns(word: Bytes, table: Bytes, fc: number, lcb: number): FkpRuns {
	const builder = new RunBuilder()

	for (const pn of readBte(table, fc, lcb, word).pns) {
		const page = pn * PAGE
		const cpara = word.u8(page + PAGE - 1)

		// cpara is 1..0x1D; each BxPap is 13 bytes (§2.9.174).
		if (cpara === 0 || 4 * (cpara + 1) + 13 * cpara > PAGE - 1) continue

		for (let k = 0; k < cpara; k++) {
			const start = word.u32(page + 4 * k)
			const end = word.u32(page + 4 * (k + 1))
			const offset = word.u8(page + 4 * (cpara + 1) + 13 * k) * 2

			if (offset === 0) {
				builder.push(start, end, 0, 0)
				continue
			}

			// PapxInFkp (§2.9.175): cb, or 0 then cb'; the GrpPrlAndIstd is 2*cb-1 or 2*cb' bytes.
			const cb = word.u8(page + offset)
			let at = page + offset + 1
			let size = 2 * cb - 1

			if (cb === 0) {
				size = 2 * word.u8(at)
				at++
			}

			builder.push(start, end, at, size >= 2 && at + size <= page + PAGE - 1 ? size : 0)
		}
	}

	return builder.build()
}

// Index of the last run with fcStart <= fc, or -1.
export function runAtOrBefore(runs: FkpRuns, fc: number): number {
	let low = 0
	let high = runs.count - 1

	if (high < 0 || (runs.fcStart[0] ?? 0) > fc) return -1

	while (low < high) {
		const middle = (low + high + 1) >> 1

		if ((runs.fcStart[middle] ?? 0) <= fc) low = middle
		else high = middle - 1
	}

	return low
}

// The run covering fc, or -1.
export function runCovering(runs: FkpRuns, fc: number): number {
	const k = runAtOrBefore(runs, fc)

	return k >= 0 && fc < (runs.fcEnd[k] ?? 0) ? k : -1
}
