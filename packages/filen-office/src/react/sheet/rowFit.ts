import {
	TILE_COLS,
	TILE_ROWS,
	tileKey,
	tileKeyCol,
	tileKeyRow,
	tileKeySheet,
	type SheetAxisLayout,
	type SheetChange,
	type SheetId,
	type SheetRenderStyle
} from "../../core/sheet/protocol"
// The axis.fit command's type (erased: no core code comes with it).
import type {} from "../../core/xlsx/edit/formatFeature"
import type { SheetGridModel } from "./model"
import { TEXT_PADDING } from "./paint/cells"
import { layoutText, textFont, type TextMeasurer, type TextRun } from "./paint/text"
import { entryRuns, entryText, regionHoldsTile, type GridTile } from "./tiles"

// Rows without a height of their own fit their wrapped text, as Excel fits them when wrapped text changes (wrap
// turned on or off, an entry, a column's width): the core measures no text, so the grid measures the wrapped cells
// of the rows an edit touched, as it draws them, and sends their heights (axis.fit), one step a user can undo. Only
// rows whose tiles the grid holds (on screen or about to be) are measured, never the sheet at open: a file's
// saved heights show as they are.

// A line of text and the row's padding beyond its lines, as autofit measures them.
const LINE_HEIGHT = 1.2
const ROW_PADDING = 2

export interface TextMeasure {
	// Null while the grid draws nothing (no canvas yet).
	measurer(): TextMeasurer | null
	// The theme's font stack, for a family the file names but the platform lacks.
	font(): string
}

// A row's cells in one tile as a string (column, style and text of each), and whether any wraps: a row is fitted
// when this changed and it wraps before or after.
interface RowLook {
	readonly look: string
	readonly wraps: boolean
}

function rowLooks(model: SheetGridModel, tile: GridTile): Map<number, RowLook> {
	const looks = new Map<number, { look: string; wraps: boolean }>()

	for (let i = 0; i < tile.count; i++) {
		const position = tile.positions[i] ?? 0
		const row = tile.tileRow * TILE_ROWS + Math.floor(position / TILE_COLS)
		const style = tile.styles[i] ?? 0
		const text = entryText(tile, i)
		const part = `${String(position % TILE_COLS)}\u0000${String(style)}\u0000${text}\u0000`
		const known = looks.get(row)
		const wraps = text !== "" && wrapsText(model.style(style))

		if (known === undefined) looks.set(row, { look: part, wraps })
		else {
			known.look += part
			known.wraps ||= wraps
		}
	}

	return looks
}

// The columns whose widths differ between two layouts of a sheet's columns; null: every column (the default width).
function changedCols(before: SheetAxisLayout, after: SheetAxisLayout): Set<number> | null {
	if (before.defaultSize !== after.defaultSize) return null

	const sizes = new Map<number, number>()
	const changed = new Set<number>()

	before.indices.forEach((col, k) => sizes.set(col, before.sizes[k] ?? 0))
	after.indices.forEach((col, k) => {
		if (sizes.get(col) !== after.sizes[k]) changed.add(col)

		sizes.delete(col)
	})

	for (const col of sizes.keys()) changed.add(col)

	return changed
}

export class RowFitter {
	private readonly model: SheetGridModel
	private measure: TextMeasure | null = null
	// Tiles an edit changes, with their rows' looks before it; the rows a column's new width rewraps.
	private readonly before = new Map<number, Map<number, RowLook>>()
	private readonly rewrapped = new Map<SheetId, Set<number>>()
	private scheduled = false

	constructor(model: SheetGridModel) {
		this.model = model
	}

	// Measures with the grid's text measurer while attached; returns the detach.
	attach(measure: TextMeasure): () => void {
		this.measure = measure

		return () => {
			if (this.measure === measure) this.measure = null

			this.before.clear()
			this.rewrapped.clear()
		}
	}

	// What an edit changes, before it is folded in: the held tiles it patches or drops, kept as they look now; and
	// the rows with wrapped text in a column whose width changes.
	changed(change: SheetChange): void {
		if (this.measure === null) return

		const model = this.model
		const keep = (key: number): void => {
			const tile = model.tiles.peek(key)

			if (tile !== undefined && !this.before.has(key)) this.before.set(key, rowLooks(model, tile))
		}

		for (const patch of change.patches) keep(tileKey(patch.sheet, patch.tileRow, patch.tileCol))

		for (const region of change.invalidated)
			for (const key of model.tiles.keys())
				if (regionHoldsTile(region, tileKeySheet(key), tileKeyRow(key), tileKeyCol(key))) keep(key)

		for (const patch of change.layouts) {
			const cols = patch.layout.cols
			const held = model.layout(patch.sheet)?.cols

			if (cols === undefined || held === undefined) continue

			const changed = changedCols(held, cols)

			if (changed !== null && changed.size === 0) continue

			let rows = this.rewrapped.get(patch.sheet)

			for (const key of model.tiles.keys()) {
				const tile = model.tiles.peek(key)

				if (tile === undefined || tile.sheet !== patch.sheet) continue

				for (let i = 0; i < tile.count; i++) {
					const position = tile.positions[i] ?? 0

					if (changed !== null && !changed.has(tile.tileCol * TILE_COLS + (position % TILE_COLS))) continue
					if (!wrapsText(model.style(tile.styles[i] ?? 0))) continue

					if (rows === undefined) {
						rows = new Set()
						this.rewrapped.set(patch.sheet, rows)
					}

					rows.add(tile.tileRow * TILE_ROWS + Math.floor(position / TILE_COLS))
				}
			}
		}

		this.schedule()
	}

	// Tiles that came in: a dropped one an edit changed is compared now.
	arrived(keys: Iterable<number>): void {
		if (this.measure === null || this.before.size === 0) return

		for (const key of keys) {
			if (this.before.has(key)) {
				this.schedule()

				return
			}
		}
	}

	private schedule(): void {
		if (this.scheduled || (this.before.size === 0 && this.rewrapped.size === 0)) return

		this.scheduled = true
		// After the change is folded in and drawn, so the widths and tiles measured are the new ones.
		setTimeout(() => {
			this.scheduled = false
			this.run()
		}, 0)
	}

	private run(): void {
		const measurer = this.measure?.measurer() ?? null
		const snapshot = this.model.getSnapshot()

		if (measurer === null || this.measure === null || snapshot.readOnly || snapshot.writability.state === "viewOnly") {
			this.before.clear()
			this.rewrapped.clear()

			return
		}

		const rows = new Map<SheetId, Set<number>>()
		const rowsOf = (sheet: SheetId): Set<number> => {
			let set = rows.get(sheet)

			if (set === undefined) {
				set = new Set()
				rows.set(sheet, set)
			}

			return set
		}
		// Rows that wrapped before an edit: sent even when nothing wraps now (wrap turned off).
		const wrappedBefore = new Set<number>()

		for (const [key, before] of this.before) {
			const tile = this.model.tiles.peek(key)

			// Dropped and not back yet: compared when it arrives.
			if (tile === undefined) continue

			this.before.delete(key)

			const after = rowLooks(this.model, tile)
			const sheet = tileKeySheet(key)

			for (const row of new Set([...before.keys(), ...after.keys()])) {
				const was = before.get(row)
				const now = after.get(row)

				if (was?.look === now?.look || (was?.wraps !== true && now?.wraps !== true)) continue

				rowsOf(sheet).add(row)

				if (was?.wraps === true) wrappedBefore.add(row)
			}
		}

		for (const [sheet, set] of this.rewrapped) for (const row of set) rowsOf(sheet).add(row)

		this.rewrapped.clear()

		for (const [sheet, set] of rows) {
			const fitted = measureRows(this.model, sheet, set, measurer, this.measure.font(), wrappedBefore)

			if (fitted !== null)
				void this.model.command(
					{
						type: "axis.fit",
						sheet,
						rows: fitted.rows,
						heights: fitted.heights,
						firstCol: fitted.firstCol,
						lastCol: fitted.lastCol
					},
					false
				)
		}
	}
}

// Text drawn over more than one line when it is longer than its cell (paint/cells.ts).
function wrapsText(style: SheetRenderStyle): boolean {
	return (style.wrap === true || style.align === "justify" || style.align === "distributed") && (style.rotation ?? 0) === 0
}

interface MeasuredRows {
	readonly rows: Uint32Array
	readonly heights: Float32Array
	readonly firstCol: number
	readonly lastCol: number
}

// Each row's tallest wrapped cell among the tiles held, at 100 % (NaN: no wrapped text there, for a row that held
// some before), over the columns those tiles cover. Null when no row has wrapped text to fit.
export function measureRows(
	model: SheetGridModel,
	sheet: SheetId,
	rows: ReadonlySet<number>,
	measurer: TextMeasurer,
	fallback: string,
	wrappedBefore: ReadonlySet<number>
): MeasuredRows | null {
	const layout = model.layout(sheet)
	const axes = model.axes(sheet)

	if (layout === null || axes === null) return null

	const merges = model.merges(layout)
	const base = model.style(0)
	const tiles = new Map<number, GridTile[]>()
	let firstCol = Number.POSITIVE_INFINITY
	let lastCol = -1

	for (const key of model.tiles.keys()) {
		const tile = model.tiles.peek(key)

		if (tile === undefined || tile.sheet !== sheet) continue

		const band = tiles.get(tile.tileRow)

		if (band === undefined) tiles.set(tile.tileRow, [tile])
		else band.push(tile)

		firstCol = Math.min(firstCol, tile.tileCol * TILE_COLS)
		lastCol = Math.max(lastCol, tile.tileCol * TILE_COLS + TILE_COLS - 1)
	}

	const heights = new Map<number, number>()

	for (const row of rows) {
		const band = tiles.get(Math.floor(row / TILE_ROWS))

		if (band === undefined) continue

		let tallest = Number.NaN

		for (const tile of band) {
			for (let i = 0; i < tile.count; i++) {
				const position = tile.positions[i] ?? 0

				if (Math.floor(position / TILE_COLS) !== row % TILE_ROWS) continue

				const col = tile.tileCol * TILE_COLS + (position % TILE_COLS)
				const style = model.style(tile.styles[i] ?? 0)

				// Excel fits no row to a merged cell's text.
				if (!wrapsText(style)) continue
				if (merges.some(merge => row >= merge.startRow && row <= merge.endRow && col >= merge.startCol && col <= merge.endCol))
					continue

				const text = entryText(tile, i)

				if (text === "") continue

				const raw = entryRuns(tile, i)
				const runs: TextRun[] | null = raw === null ? null : []

				for (let k = 0; raw !== null && runs !== null && k + 1 < raw.length; k += 2)
					runs.push({ start: raw[k] ?? 0, style: model.style(raw[k + 1] ?? 0) })

				const indent =
					(style.indent ?? 0) > 0 ? (style.indent ?? 0) * measurer.width(textFont(style, base, fallback, 1).css, "   ") : 0
				const width = Math.max(1, axes.cols.size(col) / axes.scale - 2 * TEXT_PADDING - 1 - indent)
				const lines = layoutText(measurer, text, runs, style, { wrap: true, width, base, fallback, scale: 1 })
				const height = lines.reduce((sum, line) => sum + Math.round(line.px * LINE_HEIGHT), 0) + ROW_PADDING

				tallest = Number.isNaN(tallest) ? height : Math.max(tallest, height)
			}
		}

		if (!Number.isNaN(tallest) || wrappedBefore.has(row)) heights.set(row, tallest)
	}

	if (heights.size === 0) return null

	const sorted = [...heights.keys()].sort((a, b) => a - b)

	return {
		rows: Uint32Array.from(sorted),
		heights: Float32Array.from(sorted, row => heights.get(row) ?? Number.NaN),
		firstCol,
		lastCol
	}
}
