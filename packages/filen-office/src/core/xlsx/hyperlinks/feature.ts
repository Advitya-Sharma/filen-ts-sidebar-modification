import { CellFlag, SHEET_MAX_COLS, TILE_COLS, TILE_ROWS, type SheetCell, type SheetRange } from "../../sheet/protocol"
import { defineSheetFeature, type SheetTileBuilder } from "../../sheet/registry"
import { annotationStore } from "../comments/store"
import type { XlsxDocument } from "../edit/document"
import { FormulaTable } from "../model/formulas"
import type { SheetStore } from "../model/store"
import { FormulaKind, type XlsxFormulaTable, type XlsxSheetModel } from "../model/types"
import { linksOf, removeLinks, setLinks, type LinkTarget } from "./edit"
import { linkAt, type LinkEntry, type LinksSection } from "./links"
import { hyperlinkArguments, resolvePlace, stringLiteral, type LinkPlace } from "./location"

// Hyperlinks: cells under a link of the sheet's hyperlinks section, and cells whose formula is a HYPERLINK call, carry
// CellFlag.HYPERLINK; link.get says where one goes (an external URL, which the app opens through the shared link
// policy, or a place in the workbook); link.set and link.remove edit the section (Ctrl+K).

declare module "../../sheet/protocol" {
	interface SheetCommandMap {
		// A link on every cell of `ranges`, replacing the links there: an external URL, or a location in the workbook
		// ("Sheet2!A1", a defined name). `text`: what each range's first cell shows; `tooltip`: Excel's ScreenTip.
		"link.set": {
			readonly sheet: SheetId
			readonly ranges: readonly SheetRange[]
			readonly target: { readonly url: string } | { readonly location: string }
			readonly tooltip?: string
			readonly text?: string
		}
		// The links over `ranges` removed, with the look a link gives.
		"link.remove": { readonly sheet: SheetId; readonly ranges: readonly SheetRange[] }
	}
	interface SheetQueryMap {
		"link.get": { request: SheetCell; result: SheetLink | null }
	}
	interface SheetRefusalReasons {
		// A link target the shared link policy refuses, or a location that names no cells of the workbook.
		linkTarget: true
	}
}

export interface SheetLink {
	// An external target as written (the app judges it before opening it), null for a place in the workbook.
	readonly url: string | null
	// A place in the workbook as written, and where it is (null when it names nothing there).
	readonly location: string | null
	readonly place: LinkPlace | null
	readonly tooltip: string | null
	readonly display: string | null
	// The cells the link covers; null for a HYPERLINK formula's.
	readonly range: SheetRange | null
	// A HYPERLINK formula's (not editable as a link: its formula is).
	readonly formula: boolean
}

const HYPERLINK_CALL = /^\s*(?:_xlfn\.)?HYPERLINK\s*\(/i

// Whether any formula of a sheet calls HYPERLINK first. The table is append-only (an edit's formula is a new one;
// a structure edit rewrites references, never a call's name), so only formulas added since are read.
const tablesWithLinks = new WeakMap<XlsxFormulaTable, { length: number; any: boolean }>()

function anyHyperlinkFormula(formulas: XlsxFormulaTable): boolean {
	const known = tablesWithLinks.get(formulas) ?? { length: 0, any: false }

	if (!known.any && known.length < formulas.length) {
		let any = false

		for (let id = known.length; id < formulas.length && !any; id++) any = HYPERLINK_CALL.test(formulas.text(id))

		tablesWithLinks.set(formulas, { length: formulas.length, any })

		return any
	}

	return known.any
}

// Links by the tiles they reach, per read section; a link over more than WIDE tiles is checked by every tile.
const WIDE = 64
const byTile = new WeakMap<LinksSection, { tiles: Map<number, LinkEntry[]>; wide: LinkEntry[] }>()

function tileOf(row: number, col: number): number {
	return Math.floor(row / TILE_ROWS) * (SHEET_MAX_COLS / TILE_COLS) + Math.floor(col / TILE_COLS)
}

function tileLinks(section: LinksSection, firstRow: number, firstCol: number): readonly LinkEntry[] {
	let index = byTile.get(section)

	if (index === undefined) {
		const tiles = new Map<number, LinkEntry[]>()
		const wide: LinkEntry[] = []

		for (const link of section.links) {
			if (link.range.startRow < 0) continue

			const firstTileRow = Math.floor(link.range.startRow / TILE_ROWS)
			const lastTileRow = Math.floor(link.range.endRow / TILE_ROWS)
			const firstTileCol = Math.floor(link.range.startCol / TILE_COLS)
			const lastTileCol = Math.floor(link.range.endCol / TILE_COLS)

			if ((lastTileRow - firstTileRow + 1) * (lastTileCol - firstTileCol + 1) > WIDE) {
				wide.push(link)
				continue
			}

			for (let tileRow = firstTileRow; tileRow <= lastTileRow; tileRow++)
				for (let tileCol = firstTileCol; tileCol <= lastTileCol; tileCol++) {
					const key = tileRow * (SHEET_MAX_COLS / TILE_COLS) + tileCol
					const list = tiles.get(key)

					if (list === undefined) tiles.set(key, [link])
					else list.push(link)
				}
		}

		index = { tiles, wide }
		byTile.set(section, index)
	}

	const own = index.tiles.get(tileOf(firstRow, firstCol)) ?? []

	return index.wide.length === 0 ? own : [...own, ...index.wide]
}

// A formula's text as its shared master writes it (a call's name moves with it unchanged).
function leadText(formulas: XlsxFormulaTable, id: number): string {
	if (formulas.kind(id) !== FormulaKind.SHARED_CHILD || !(formulas instanceof FormulaTable)) return formulas.text(id)

	const master = formulas.master(formulas.sharedIndex(id))

	return master < 0 ? "" : formulas.text(master)
}

function flagLinks(document: XlsxDocument, tile: SheetTileBuilder): void {
	const model = document.model(tile.sheet)

	if (model === null) return

	const tileRange: SheetRange = {
		startRow: tile.firstRow,
		startCol: tile.firstCol,
		endRow: tile.firstRow + TILE_ROWS - 1,
		endCol: tile.firstCol + TILE_COLS - 1
	}

	const section = linksOf(document.workbook, model)

	for (const link of section === null || section === undefined ? [] : tileLinks(section, tile.firstRow, tile.firstCol)) {
		const startRow = Math.max(link.range.startRow, tileRange.startRow)
		const endRow = Math.min(link.range.endRow, tileRange.endRow)
		const startCol = Math.max(link.range.startCol, tileRange.startCol)
		const endCol = Math.min(link.range.endCol, tileRange.endCol)

		if (link.range.startRow < 0 || startRow > endRow || startCol > endCol) continue

		// A link's first cell is marked though it is blank; the rest where they have an entry.
		if (startRow === link.range.startRow && startCol === link.range.startCol)
			tile.addFlags(tile.entry(startRow, startCol), CellFlag.HYPERLINK)

		for (let row = startRow; row <= endRow; row++)
			for (let col = startCol; col <= endCol; col++) {
				const entry = tile.find(row, col)

				if (entry >= 0) tile.addFlags(entry, CellFlag.HYPERLINK)
			}
	}

	const store = model.store as SheetStore
	const formulas = store.formulas

	if (!anyHyperlinkFormula(formulas)) return

	store.forEachCell(tileRange, (row, col, slot) => {
		const id = store.formula(slot)

		if (id < 0 || !HYPERLINK_CALL.test(leadText(formulas, id))) return

		const entry = tile.find(row, col)

		if (entry >= 0) tile.addFlags(entry, CellFlag.HYPERLINK)
	})
}

// Where a HYPERLINK formula goes: its first argument as a string, a cell's text, or (without a friendly name) what the
// cell itself shows. Null when that is not known without calculating it.
function formulaLink(document: XlsxDocument, model: XlsxSheetModel, row: number, col: number): SheetLink | null {
	const slot = model.store.locate(row, col)
	const id = slot < 0 ? -1 : (model.store as SheetStore).formula(slot)

	if (id < 0) return null

	const args = hyperlinkArguments(document.formulaText(model, id, row, col))

	if (args === null || args.length === 0) return null

	const first = args[0] ?? ""
	let target = stringLiteral(first)

	if (target === null) {
		const place = /^[A-Za-z$'_]/.test(first) ? resolvePlace(document.workbook, model.id, first) : null

		if (place !== null && place.range.startRow === place.range.endRow && place.range.startCol === place.range.endCol)
			target = document.cellView(place.sheet, place.range.startRow, place.range.startCol).text
		else if (args.length === 1) target = document.cellView(model.id, row, col).text
	}

	if (target === null || target.trim() === "") return null

	const internal = target.trim().startsWith("#")

	return {
		url: internal ? null : target.trim(),
		location: internal ? target.trim().slice(1) : null,
		place: internal ? resolvePlace(document.workbook, model.id, target) : null,
		tooltip: null,
		display: null,
		range: null,
		formula: true
	}
}

async function linkAtCell(document: XlsxDocument, cell: SheetCell): Promise<SheetLink | null> {
	if (document.entry(cell.sheet)?.kind !== "worksheet") return null

	const model = await document.ready(cell.sheet)
	const link = linkAt(linksOf(document.workbook, model) ?? null, cell.row, cell.col)

	if (link === null) return formulaLink(document, model, cell.row, cell.col)

	const relationship =
		link.relationshipId === null
			? undefined
			: annotationStore(document.workbook)
					.relationships(cell.sheet)
					.find(found => found.id === link.relationshipId)

	return {
		url: relationship?.target ?? null,
		location: link.location,
		place: link.location === null ? null : resolvePlace(document.workbook, cell.sheet, link.location),
		tooltip: link.tooltip,
		display: link.display,
		range: link.range,
		formula: false
	}
}

export const linksFeature = defineSheetFeature<XlsxDocument>({
	id: "links",
	commands: {
		"link.set": {
			info: { write: true, protection: "insertHyperlinks" },
			run: (context, command) =>
				setLinks(
					context.doc,
					command.sheet,
					command.ranges,
					command.target as LinkTarget,
					command.tooltip,
					command.text,
					context.changes
				)
		},
		"link.remove": {
			info: { write: true, protection: "insertHyperlinks" },
			run: (context, command) => removeLinks(context.doc, command.sheet, command.ranges, context.changes)
		}
	},
	queries: {
		"link.get": { run: (context, request) => linkAtCell(context.doc, request) }
	},
	tiles: (context, tile) => flagLinks(context.doc, tile)
})
