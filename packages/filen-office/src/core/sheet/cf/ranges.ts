import type { SheetRange } from "../protocol"

// A rule's ranges (an sqref may hold thousands, as copying formatted cells leaves them) found by position without
// reading them all: parts spanning few rows are listed in buckets of rows, wide ones are kept apart (few in practice).

const BUCKET_ROWS = 64
// A part spanning more buckets than this is a wide one.
const BUCKETS_PER_PART = 16
// Few parts are read in turn.
const LINEAR_PARTS = 16

export function intersects(a: SheetRange, b: SheetRange): boolean {
	return a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol
}

export function contains(range: SheetRange, row: number, col: number): boolean {
	return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}

export function boundsOf(ranges: readonly SheetRange[]): SheetRange {
	let startRow = Number.POSITIVE_INFINITY
	let startCol = Number.POSITIVE_INFINITY
	let endRow = -1
	let endCol = -1

	for (const range of ranges) {
		startRow = Math.min(startRow, range.startRow)
		startCol = Math.min(startCol, range.startCol)
		endRow = Math.max(endRow, range.endRow)
		endCol = Math.max(endCol, range.endCol)
	}

	return { startRow, startCol, endRow, endCol }
}

export class RangeIndex {
	readonly ranges: readonly SheetRange[]
	readonly bounds: SheetRange
	private readonly buckets: Map<number, number[]> | null
	private readonly wide: number[] = []
	// Marks a part as seen in one walk (a part listed in several buckets is visited once).
	private readonly seen: Uint32Array
	private walk = 0

	constructor(ranges: readonly SheetRange[]) {
		this.ranges = ranges
		this.bounds = boundsOf(ranges)
		this.seen = new Uint32Array(ranges.length)

		if (ranges.length <= LINEAR_PARTS) {
			this.buckets = null

			return
		}

		const buckets = new Map<number, number[]>()

		ranges.forEach((range, index) => {
			const first = Math.floor(range.startRow / BUCKET_ROWS)
			const last = Math.floor(range.endRow / BUCKET_ROWS)

			if (last - first >= BUCKETS_PER_PART) {
				this.wide.push(index)

				return
			}

			for (let bucket = first; bucket <= last; bucket++) {
				const list = buckets.get(bucket)

				if (list === undefined) buckets.set(bucket, [index])
				else list.push(index)
			}
		})

		this.buckets = buckets
	}

	// Each part that meets `range`, once.
	each(range: SheetRange, visit: (part: SheetRange, index: number) => void): void {
		if (!intersects(this.bounds, range)) return

		if (this.buckets === null) {
			this.ranges.forEach((part, index) => {
				if (intersects(part, range)) visit(part, index)
			})

			return
		}

		for (const index of this.wide) {
			const part = this.ranges[index] as SheetRange

			if (intersects(part, range)) visit(part, index)
		}

		const walk = this.nextWalk()
		const first = Math.floor(Math.max(range.startRow, this.bounds.startRow) / BUCKET_ROWS)
		const last = Math.floor(Math.min(range.endRow, this.bounds.endRow) / BUCKET_ROWS)

		for (let bucket = first; bucket <= last; bucket++) {
			const list = this.buckets.get(bucket)

			if (list === undefined) continue

			for (const index of list) {
				if (this.seen[index] === walk) continue

				this.seen[index] = walk

				const part = this.ranges[index] as SheetRange

				if (intersects(part, range)) visit(part, index)
			}
		}
	}

	any(range: SheetRange): boolean {
		let found = false

		this.each(range, () => {
			found = true
		})

		return found
	}

	// The first part (in the sqref's order) holding a cell, or -1.
	first(row: number, col: number): number {
		let first = -1

		this.each({ startRow: row, startCol: col, endRow: row, endCol: col }, (_part, index) => {
			if (first < 0 || index < first) first = index
		})

		return first
	}

	private nextWalk(): number {
		if (this.walk === 0xffffffff) {
			this.seen.fill(0)
			this.walk = 0
		}

		return ++this.walk
	}
}
