import { CellFlag, TILE_COLS, TILE_ROWS } from "../../sheet/protocol"
import { defineSheetFeature, type SheetReadContext, type SheetTileBuilder } from "../../sheet/registry"
import { isPlainGeneral } from "../../sheet/numfmt/general"
import { DEFAULT_THEME_COLORS, resolveColor } from "../../sheet/style/color"
import type { SheetFormatChoices } from "../../sheet/style/theme"
import type { SheetStore } from "../model/store"
import { CellKind } from "../model/types"
import type { XlsxDocument } from "./document"
import { clearFormats } from "./format"
import { applyBorders } from "./formatBorders"
import { fitRows } from "./formatHeights"

// The formatting the toolbar adds beyond format.apply: borders, what the pickers offer from the workbook, and the
// values of numbers shown in General (their tile layer), so a narrow cell rounds them as Excel does.

declare module "../../sheet/protocol" {
	interface SheetCommandMap {
		// Rows without a height of their own fit their content (formatHeights.ts): `heights` are each row's tallest
		// wrapped cell the grid measured over columns firstCol..lastCol, in px at 100 % (NaN: none there).
		"axis.fit": {
			readonly sheet: SheetId
			readonly rows: Uint32Array
			readonly heights: Float32Array
			readonly firstCol: number
			readonly lastCol: number
		}
	}
}

const THEME_COLORS = 12

export function formatChoices(document: XlsxDocument): SheetFormatChoices {
	const theme = document.workbook.theme
	const palette = document.styles.palette
	const themeColors: number[] = []

	for (let index = 0; index < THEME_COLORS; index++) {
		const rgb = resolveColor({ theme: index }, theme, palette)

		themeColors.push(rgb >= 0 ? rgb : (DEFAULT_THEME_COLORS[index] ?? 0))
	}

	const fonts: string[] = []
	const seen = new Set<string>()

	for (const font of document.workbook.styles.fonts) {
		const name = font.name?.trim() ?? ""

		if (name === "" || seen.has(name.toLowerCase())) continue

		seen.add(name.toLowerCase())
		fonts.push(name)
	}

	return { themeColors, majorFont: theme?.majorFont ?? "", minorFont: theme?.minorFont ?? "", fonts }
}

// The "general" layer of a tile: every entry showing a number through plain General, with its value.
export function generalValues(context: SheetReadContext<XlsxDocument>, tile: SheetTileBuilder): void {
	const document = context.doc
	// A sheet still parsing shows its first rows' preview, as fillTile does.
	const model = document.model(tile.sheet) ?? document.previews.get(tile.sheet)?.model ?? null

	if (model === null || tile.count === 0) return

	const store = model.store as SheetStore
	const general = new Map<number, boolean>()
	const entries: number[] = []
	const values: number[] = []
	const shown = document.state.bind(document)

	store.forEachCell(
		{ startRow: tile.firstRow, startCol: tile.firstCol, endRow: tile.firstRow + TILE_ROWS - 1, endCol: tile.firstCol + TILE_COLS - 1 },
		(row, col, slot) => {
			const kind = store.kind(slot)
			const formula = store.formula(slot)
			let value: number

			if (kind === CellKind.NUMBER) {
				value = store.value(slot)
			} else if (kind === CellKind.EMPTY && formula >= 0) {
				const result = shown(tile.sheet, row, col).shown

				if (result === undefined || result.error || typeof result.value !== "number") return

				value = result.value
			} else {
				return
			}

			const xf = store.style(slot)
			let plain = general.get(xf)

			if (plain === undefined) {
				plain = isPlainGeneral(document.numFmt(xf))
				general.set(xf, plain)
			}

			if (!plain) return

			const entry = tile.find(row, col)

			if (entry < 0 || (tile.flags(entry) & (CellFlag.NUMBER | CellFlag.CLIPPED)) !== CellFlag.NUMBER) return

			entries.push(entry)
			values.push(value)
		}
	)

	if (entries.length === 0) return

	tile.layer("general", {
		entries: Uint16Array.from(entries),
		values: Float64Array.from(values),
		decimal: context.settings.locale?.decimal ?? "."
	})
}

export const formattingFeature = defineSheetFeature<XlsxDocument>({
	id: "formatting",
	commands: {
		"format.clear": {
			info: { write: true, protection: "formatCells" },
			run: (context, command) => clearFormats(context.doc, command.sheet, command.ranges, context.changes, context.slicing)
		},
		"axis.fit": {
			info: { write: true },
			run: (context, command) => {
				if (command.rows.length !== command.heights.length) return Promise.resolve({ refused: { reason: "tooLarge" } })

				const heights = new Map<number, number>()

				command.rows.forEach((row, k) => heights.set(row, command.heights[k] ?? Number.NaN))

				return fitRows(
					context.doc,
					command.sheet,
					{ heights, firstCol: command.firstCol, lastCol: command.lastCol },
					context.changes
				)
			}
		},
		"format.borders": {
			info: { write: true, protection: "formatCells" },
			run: (context, command) =>
				applyBorders(context.doc, command.sheet, command.ranges, command.edges, context.changes, context.slicing)
		}
	},
	queries: {
		"format.choices": { run: context => formatChoices(context.doc) }
	},
	tiles: generalValues
})
