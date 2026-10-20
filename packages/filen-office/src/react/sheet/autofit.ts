import { MAX_COL_WIDTH, MAX_ROW_HEIGHT, MIN_SIZE } from "../../core/sheet/sizes/sizes"
import { TILE_COLS, TILE_ROWS, type SheetAxis, type SheetId } from "../../core/sheet/protocol"
import { createGridAxis } from "./axis"
import type { SheetGridModel } from "./model"
import { TEXT_PADDING } from "./paint/cells"
import { layoutText, textFont, type TextMeasurer } from "./paint/text"
import { entryText } from "./tiles"

// Autofit: the session names each column's or row's longest texts with their styles (fitSample; the core has no
// fonts), and the grid measures them with the canvas it draws with, at 100 %. A wrapped text fits a row at its
// column's width when its cell is on screen (its tile held), else by its own line breaks.

// Texts measured per index.
const FIT_SAMPLE = 32
// Indices one autofit measures (a select-all on a wide sheet fits the first ones).
export const MAX_FIT_INDICES = 1024
const LINE_HEIGHT = 1.2
const ROW_PADDING = 2

function clamp(axis: SheetAxis, px: number): number {
	return Math.round(Math.min(axis === "cols" ? MAX_COL_WIDTH : MAX_ROW_HEIGHT, Math.max(MIN_SIZE, px)))
}

// Pixel sizes at 100 % for `indices`, NaN where an index goes back to the default.
export async function autofitSizes(
	model: SheetGridModel,
	sheet: SheetId,
	axis: SheetAxis,
	indices: readonly number[],
	measurer: TextMeasurer,
	fallback: string
): Promise<{ readonly indices: Uint32Array; readonly sizes: Float32Array }> {
	const chosen = indices.slice(0, MAX_FIT_INDICES)
	const samples = await Promise.all(chosen.map(index => model.session.query("fitSample", { sheet, axis, index, limit: FIT_SAMPLE })))
	const base = model.style(0)
	const layout = model.layout(sheet)
	const defaultRow = layout?.rows.defaultSize ?? 20
	const sizes = new Float32Array(chosen.length)

	chosen.forEach((index, k) => {
		const sample = samples[k]
		let largest = 0

		sample?.texts.forEach((text, i) => {
			const style = model.style(sample.styles[i] ?? 0)
			const font = textFont(style, base, fallback, 1)
			const wraps = (sample.wraps[i] ?? 0) !== 0

			if (axis === "cols") {
				const lines = wraps ? text.split(/\r?\n|\r/u) : [text.replace(/\r?\n|\r/gu, "")]

				for (const line of lines) largest = Math.max(largest, measurer.width(font.css, line))

				return
			}

			const width = wraps ? wrapWidth(model, sheet, index, text) : null
			const lines =
				width === null
					? wraps
						? text.split(/\r?\n|\r/u).length
						: 1
					: layoutText(measurer, text, null, style, { wrap: true, width, base, fallback, scale: 1 }).length

			largest = Math.max(largest, lines * Math.round(font.px * LINE_HEIGHT))
		})

		if (largest === 0) {
			sizes[k] = Number.NaN
		} else if (axis === "cols") {
			sizes[k] = clamp("cols", Math.ceil(largest) + 2 * TEXT_PADDING + 2)
		} else {
			const height = clamp("rows", largest + ROW_PADDING)

			sizes[k] = height <= defaultRow ? Number.NaN : height
		}
	})

	return { indices: Uint32Array.from(chosen), sizes }
}

// The text width a wrapped text has in its row: its column's width at 100 %, found from the tiles held.
function wrapWidth(model: SheetGridModel, sheet: SheetId, row: number, text: string): number | null {
	const layout = model.layout(sheet)

	if (layout === null) return null

	for (const key of model.tiles.keys()) {
		const tile = model.tiles.peek(key)

		if (tile === undefined || tile.sheet !== sheet || tile.tileRow !== Math.floor(row / TILE_ROWS)) continue

		for (let i = 0; i < tile.count; i++) {
			const position = tile.positions[i] ?? 0

			if (Math.floor(position / TILE_COLS) !== row % TILE_ROWS || entryText(tile, i) !== text) continue

			const col = tile.tileCol * TILE_COLS + (position % TILE_COLS)
			const axes = model.axes(sheet)
			// The width as drawn (kept sizes and zoom included), back at 100 %.
			const width = axes === null ? createGridAxis(layout.cols, col + 1, 1).size(col) : axes.cols.size(col) / axes.scale

			return Math.max(1, width - 2 * TEXT_PADDING - 1)
		}
	}

	return null
}
