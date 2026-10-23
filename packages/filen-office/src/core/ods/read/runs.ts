import type { SheetRange } from "../../sheet/protocol"
import type { XlsxStyleRuns } from "../../xlsx/model/types"
import { crossInBand, shiftAxisSpan, type AxisShift } from "../../xlsx/refs/axis"

// Formats of empty cells kept as runs (XlsxStyleRuns): runs of rows, each with ascending runs of columns. A run of
// rows stands for every row of a repeated table:table-row and for consecutive rows whose runs are equal, so a
// styled block of a million rows costs one entry.
export class StyleRuns implements XlsxStyleRuns {
	readonly rowExtent: number
	readonly colExtent: number
	private readonly rowStarts: Uint32Array
	private readonly rowEnds: Uint32Array
	// rowStarts.length + 1 offsets into the column arrays.
	private readonly firstRun: Uint32Array
	private readonly colStarts: Uint32Array
	private readonly colEnds: Uint32Array
	private readonly styles: Uint32Array

	constructor(
		rowStarts: Uint32Array,
		rowEnds: Uint32Array,
		firstRun: Uint32Array,
		colStarts: Uint32Array,
		colEnds: Uint32Array,
		styles: Uint32Array
	) {
		this.rowStarts = rowStarts
		this.rowEnds = rowEnds
		this.firstRun = firstRun
		this.colStarts = colStarts
		this.colEnds = colEnds
		this.styles = styles
		this.rowExtent = rowEnds.length === 0 ? 0 : (rowEnds[rowEnds.length - 1] ?? 0)

		let colExtent = 0

		for (let k = 0; k < colEnds.length; k++) colExtent = Math.max(colExtent, colEnds[k] ?? 0)

		this.colExtent = colExtent
	}

	get rowRuns(): number {
		return this.rowStarts.length
	}

	get cellRuns(): number {
		return this.colStarts.length
	}

	// The row run holding `row`, -1 for none.
	private rowRun(row: number): number {
		let low = 0
		let high = this.rowStarts.length - 1

		while (low <= high) {
			const mid = (low + high) >>> 1

			if ((this.rowStarts[mid] ?? 0) <= row) low = mid + 1
			else high = mid - 1
		}

		return high >= 0 && row < (this.rowEnds[high] ?? 0) ? high : -1
	}

	styleAt(row: number, col: number): number {
		const run = this.rowRun(row)

		if (run < 0) return -1

		let low = this.firstRun[run] ?? 0
		let high = (this.firstRun[run + 1] ?? 0) - 1

		while (low <= high) {
			const mid = (low + high) >>> 1

			if ((this.colStarts[mid] ?? 0) <= col) low = mid + 1
			else high = mid - 1
		}

		return high >= (this.firstRun[run] ?? 0) && col < (this.colEnds[high] ?? 0) ? (this.styles[high] ?? 0) : -1
	}

	forEach(range: SheetRange, visit: (startRow: number, startCol: number, endRow: number, endCol: number, style: number) => void): void {
		let run = this.rowRun(range.startRow)

		if (run < 0) {
			// The first run starting inside the range.
			let low = 0
			let high = this.rowStarts.length - 1

			while (low <= high) {
				const mid = (low + high) >>> 1

				if ((this.rowStarts[mid] ?? 0) < range.startRow) low = mid + 1
				else high = mid - 1
			}

			run = low
		}

		for (; run < this.rowStarts.length && (this.rowStarts[run] ?? 0) <= range.endRow; run++) {
			const startRow = Math.max(range.startRow, this.rowStarts[run] ?? 0)
			const endRow = Math.min(range.endRow, (this.rowEnds[run] ?? 0) - 1)

			for (let k = this.firstRun[run] ?? 0; k < (this.firstRun[run + 1] ?? 0); k++) {
				const startCol = Math.max(range.startCol, this.colStarts[k] ?? 0)
				const endCol = Math.min(range.endCol, (this.colEnds[k] ?? 0) - 1)

				if (startCol <= endCol) visit(startRow, startCol, endRow, endCol, this.styles[k] ?? 0)
			}
		}
	}

	estimateBytes(): number {
		return 4 * (this.rowStarts.length * 3 + 1 + this.colStarts.length * 3)
	}

	// Each run as a rectangle [startRow, endRow), [startCol, endCol).
	private rectangles(): Rectangle[] {
		const out: Rectangle[] = []

		for (let run = 0; run < this.rowStarts.length; run++) {
			for (let k = this.firstRun[run] ?? 0; k < (this.firstRun[run + 1] ?? 0); k++)
				out.push({
					startRow: this.rowStarts[run] ?? 0,
					endRow: this.rowEnds[run] ?? 0,
					startCol: this.colStarts[k] ?? 0,
					endCol: this.colEnds[k] ?? 0,
					style: this.styles[k] ?? 0
				})
		}

		return out
	}

	shifted(shift: AxisShift): StyleRuns | undefined {
		const rows = shift.axis === "rows"
		const moved: Rectangle[] = []

		for (const rectangle of this.rectangles()) {
			// A band edit moves only the band's part of a run across it (refs/axis.ts).
			const parts = shift.band === undefined ? [rectangle] : splitAtBand(rectangle, shift)

			for (const part of parts) {
				const start = rows ? part.startRow : part.startCol
				const end = rows ? part.endRow : part.endCol

				if (!crossInBand(rows ? part.startCol : part.startRow, shift)) {
					moved.push(part)

					continue
				}

				const span = shiftAxisSpan(start, end - 1, shift)

				if (span === null) continue

				moved.push(rows ? { ...part, startRow: span[0], endRow: span[1] + 1 } : { ...part, startCol: span[0], endCol: span[1] + 1 })
			}
		}

		return fromRectangles(moved)
	}
}

interface Rectangle {
	readonly startRow: number
	readonly endRow: number
	readonly startCol: number
	readonly endCol: number
	readonly style: number
}

// The parts of a run inside and outside a band edit's band (columns for an edit of rows, rows for one of columns).
function splitAtBand(rectangle: Rectangle, shift: AxisShift): Rectangle[] {
	const band = shift.band

	if (band === undefined) return [rectangle]

	const rows = shift.axis === "rows"
	const low = rows ? rectangle.startCol : rectangle.startRow
	const high = rows ? rectangle.endCol : rectangle.endRow
	const cuts = [low, Math.max(low, Math.min(high, band.start)), Math.max(low, Math.min(high, band.end + 1)), high]
	const out: Rectangle[] = []

	for (let k = 0; k < 3; k++) {
		const from = cuts[k] ?? 0
		const to = cuts[k + 1] ?? 0

		if (to > from) out.push(rows ? { ...rectangle, startCol: from, endCol: to } : { ...rectangle, startRow: from, endRow: to })
	}

	return out
}

// Rectangles that do not overlap, as runs: swept row boundary by row boundary.
function fromRectangles(rectangles: readonly Rectangle[]): StyleRuns | undefined {
	if (rectangles.length === 0) return undefined

	const bounds = [...new Set(rectangles.flatMap(rectangle => [rectangle.startRow, rectangle.endRow]))].sort((a, b) => a - b)
	const byStart = [...rectangles].sort((a, b) => a.startRow - b.startRow)
	const builder = new StyleRunsBuilder()
	let active: Rectangle[] = []
	let next = 0

	for (let k = 0; k + 1 < bounds.length; k++) {
		const from = bounds[k] ?? 0
		const to = bounds[k + 1] ?? 0

		active = active.filter(rectangle => rectangle.endRow > from)

		while (next < byStart.length && (byStart[next]?.startRow ?? 0) <= from) {
			const rectangle = byStart[next++]

			if (rectangle !== undefined && rectangle.endRow > from) active.push(rectangle)
		}

		if (active.length === 0) continue

		active.sort((a, b) => a.startCol - b.startCol)

		for (const rectangle of active) builder.add(rectangle.startCol, rectangle.endCol, rectangle.style)

		builder.endRows(from, to)
	}

	return builder.finish()
}

// Collects runs row by row, ascending: a row run equal to the one before and adjacent to it extends it.
export class StyleRunsBuilder {
	private rowStarts: number[] = []
	private rowEnds: number[] = []
	private firstRun: number[] = [0]
	private colStarts: number[] = []
	private colEnds: number[] = []
	private styles: number[] = []
	// The row under way's column runs.
	private pendingStarts: number[] = []
	private pendingEnds: number[] = []
	private pendingStyles: number[] = []

	// A run of columns [start, end) of the row under way; ascending, not overlapping.
	add(start: number, end: number, style: number): void {
		const last = this.pendingEnds.length - 1

		if (last >= 0 && this.pendingEnds[last] === start && this.pendingStyles[last] === style) {
			this.pendingEnds[last] = end

			return
		}

		this.pendingStarts.push(start)
		this.pendingEnds.push(end)
		this.pendingStyles.push(style)
	}

	get pending(): boolean {
		return this.pendingStarts.length > 0
	}

	// Ends the runs added since the last call: they hold for rows [startRow, endRow).
	endRows(startRow: number, endRow: number): void {
		const count = this.pendingStarts.length

		if (count === 0) return

		const previous = this.rowStarts.length - 1

		if (previous >= 0 && this.rowEnds[previous] === startRow && this.sameAsLast(count)) {
			this.rowEnds[previous] = endRow
		} else {
			this.rowStarts.push(startRow)
			this.rowEnds.push(endRow)

			for (let k = 0; k < count; k++) {
				this.colStarts.push(this.pendingStarts[k] ?? 0)
				this.colEnds.push(this.pendingEnds[k] ?? 0)
				this.styles.push(this.pendingStyles[k] ?? 0)
			}

			this.firstRun.push(this.colStarts.length)
		}

		if (this.pendingStarts.length > 0) {
			this.pendingStarts.length = 0
			this.pendingEnds.length = 0
			this.pendingStyles.length = 0
		}
	}

	private sameAsLast(count: number): boolean {
		const from = this.firstRun[this.firstRun.length - 2] ?? 0

		if (this.colStarts.length - from !== count) return false

		for (let k = 0; k < count; k++) {
			if (
				this.colStarts[from + k] !== this.pendingStarts[k] ||
				this.colEnds[from + k] !== this.pendingEnds[k] ||
				this.styles[from + k] !== this.pendingStyles[k]
			) {
				return false
			}
		}

		return true
	}

	finish(): StyleRuns | undefined {
		if (this.rowStarts.length === 0) return undefined

		return new StyleRuns(
			Uint32Array.from(this.rowStarts),
			Uint32Array.from(this.rowEnds),
			Uint32Array.from(this.firstRun),
			Uint32Array.from(this.colStarts),
			Uint32Array.from(this.colEnds),
			Uint32Array.from(this.styles)
		)
	}
}
