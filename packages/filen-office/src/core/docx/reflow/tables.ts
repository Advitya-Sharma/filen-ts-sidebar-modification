import type { Node } from "prosemirror-model"
import { refItem, refPart } from "../layout/edits"
import { bottomEdge, cellFill, cellFormat, cellItems, endEdge, rowFormat, startEdge, tableRows, topEdge } from "../layout/table"
import type { DocxModel } from "../model/index"
import { Item, RevisionView } from "../model/kinds"
import type { StoryPart } from "../model/part"
import type { PropsStore, ResolvedBorder, ResolvedCellMargins, WidthProps } from "../props/index"
import type { CellFormat, PropertyResolver, TableFormat } from "../styles/index"
import { borderCss, shadingCss } from "./css"

// A table's look in the reflowed view as the pages resolve it (layout/table.ts): its style's and its own
// properties, each row's exceptions, each cell's conditional formatting and its own. Every edge is drawn once,
// by the cell the pages draw it with (the top of the first row, every cell's bottom, every cell's start edge
// and the last cell's end edge), where two borders meet the one the pages keep (`heavier`).

export interface TableLook {
	// The table's declarations (direction, shading behind its cells).
	readonly table: string
	// Per row, per cell (all of them, as the editing document holds them), the cell's declarations.
	readonly cells: readonly (readonly string[])[]
	// The item of each cell, per row, in the table's part (a table inside a cell is found among its children).
	readonly items: readonly (readonly number[])[]
	readonly part: StoryPart
}

interface Slot {
	readonly item: number
	readonly format: CellFormat
	// The grid column it starts at.
	readonly column: number
	readonly continued: boolean
}

const looks = new WeakMap<StoryPart, Map<number, TableLook>>()

function px(twips: number): string {
	return `${String(Math.round((twips / 15) * 100) / 100)}px`
}

function margin(width: WidthProps): string {
	return width.type === "dxa" || width.type === "nil" ? px(width.value) : "0px"
}

function padding(margins: ResolvedCellMargins): string {
	return `padding: ${margin(margins.top)} ${margin(margins.end)} ${margin(margins.bottom)} ${margin(margins.start)}`
}

// The tables directly inside cell `cell` of `part`, in order.
export function cellTables(part: StoryPart, cell: number): number[] {
	const out: number[] = []

	for (let item = cell + 1; item < part.end(cell);) {
		const kind = part.kinds[item]

		if (kind === Item.Table) {
			out.push(item)
			item = part.end(item)
		} else if (kind === Item.Sdt || kind === Item.CustomXml) {
			item++
		} else {
			item = part.end(item)
		}
	}

	return out
}

function edge(side: string, border: ResolvedBorder | undefined): string {
	return `border-${side}: ${border === undefined || border.style === "nil" ? "none" : (borderCss(border) ?? "none")}`
}

// The look of the table at `item` of `part` (`columns` grid columns), kept per part and item.
export function tableLook(part: StoryPart, item: number, columns: number, resolver: PropertyResolver, store: PropsStore): TableLook {
	let byItem = looks.get(part)
	const cached = byItem?.get(item)

	if (cached !== undefined) return cached

	const document = part.document
	const tblPr = store.table(document, part.tables.tblPr[part.aux[item] ?? -1] ?? -1)
	const direct = resolver.table(tblPr)
	const rows = tableRows(part, item, RevisionView.Markup)
	const infos: { table: TableFormat; slots: Slot[] }[] = rows.map((row, rowIndex) => {
		const position = { rowIndex, rowCount: rows.length }
		const { table, props } = rowFormat(resolver, store, part, item, row, direct, position)
		let column = props.gridBefore

		return {
			table,
			slots: cellItems(part, row, RevisionView.Markup).map(cell => {
				const own = store.cell(document, part.cells.tcPr[part.aux[cell] ?? -1] ?? -1)
				const { format, span } = cellFormat(resolver, table, own, position, column, columns)
				const slot = { item: cell, format, column, continued: format.props.vMerge === "continue" }

				column += span

				return slot
			})
		}
	})
	const cells = infos.map((info, index) => {
		const tableBorders = info.table.props.borders
		const tableShading = info.table.props.shading
		const last = index === infos.length - 1
		const next = infos[index + 1]

		return info.slots.map((slot, k) => {
			const props = slot.format.props
			const borders = props.borders
			const out: string[] = [
				padding(props.margins),
				`vertical-align: ${props.vAlign === "center" ? "middle" : props.vAlign === "bottom" ? "bottom" : "top"}`
			]
			const below = next?.slots.find(candidate => candidate.column === slot.column)

			out.push(edge("top", index === 0 && !slot.continued ? topEdge(borders, tableBorders, true) : undefined))
			out.push(
				edge("bottom", below?.continued === true ? undefined : bottomEdge(borders, tableBorders, last, below?.format.props.borders))
			)
			out.push(edge("inline-start", startEdge(borders, tableBorders, info.slots[k - 1]?.format.props.borders, k === 0)))
			out.push(edge("inline-end", k === info.slots.length - 1 ? endEdge(borders, tableBorders) : undefined))

			const fill = shadingCss(cellFill(props.shading, tableShading))

			if (fill !== undefined) out.push(`background-color: ${fill}`)

			return out.join("; ")
		})
	})
	const look: TableLook = {
		table: direct.props.bidiVisual ? "direction: rtl" : "",
		cells,
		items: infos.map(info => info.slots.map(slot => slot.item)),
		part
	}

	if (byItem === undefined) {
		byItem = new Map()
		looks.set(part, byItem)
	}

	byItem.set(item, look)

	return look
}

// The part and item the layout lays a top-level table out from (a ref of the editor's block list).
export function laidTable(model: DocxModel, laid: number | undefined): { readonly part: StoryPart; readonly item: number } | undefined {
	if (laid === undefined || laid < 0) return undefined

	const part = model.parts[refPart(laid)]
	const item = refItem(laid)

	return part === undefined || part.kinds[item] !== Item.Table ? undefined : { part, item }
}

// The source table a table node stands for in `part` (a story the editor keeps no block list for).
export function sourceTable(part: StoryPart, table: Node): { readonly part: StoryPart; readonly item: number } | undefined {
	const src = table.attrs["src"]

	return typeof src === "number" && src >= 0 && part.kinds[src] === Item.Table ? { part, item: src } : undefined
}
