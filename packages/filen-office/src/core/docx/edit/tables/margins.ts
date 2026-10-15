import type { Node } from "prosemirror-model"
import { Item, type StoryPart } from "../../model/index"
import type { CellMarginsProps, WidthProps } from "../../props/index"
import { marginTwips, resolveCellMargins } from "../../styles/cellMargins"
import type { DocxEditContext } from "../context"
import { MAIN, dataAttr, type ElementData } from "../copies"
import { NO_SOURCE } from "../schema"
import { propertySource, tableOverrides } from "./map"

// A table's left and right cell margins (w:tblCellMar), which keep the text of every cell off its borders: what a
// column must leave room for, and what a picture in a cell may not take. Resolved as layout resolves them
// (styles/cellMargins.ts): the table's own, else its style hierarchy's (the document's default table style for a
// table naming none), each side on its own, else the standard's default.

// The w:tblCellMar a copy of another part's table carries (its source w:tblPr's).
function copied(copy: ElementData | undefined): CellMarginsProps | undefined {
	const holder = copy?.children.find(entry => entry.uri === MAIN && entry.name === "tblCellMar")

	if (holder === undefined) return undefined

	const out: { start?: WidthProps; end?: WidthProps } = {}

	for (const child of holder.children) {
		if (child.uri !== MAIN) continue

		const side =
			child.name === "left" || child.name === "start" ? "start" : child.name === "right" || child.name === "end" ? "end" : null
		const type = dataAttr(child, "type")
		const value = Number(dataAttr(child, "w") ?? "0")

		if (side !== null)
			out[side] = {
				value: Number.isFinite(value) ? value : 0,
				type: type === "nil" || type === "pct" || type === "auto" ? type : "dxa"
			}
	}

	return out
}

// The left and right cell margins of `table` together (twips); of a table not made yet when `table` is null (its
// style `styleId`).
export function tableCellMargins(part: StoryPart, context: DocxEditContext, table: Node | null, styleId?: string | null): number {
	let style = styleId ?? null
	let own: CellMarginsProps | undefined

	if (table !== null) {
		const src = propertySource(part, table, Item.Table)

		own =
			src === NO_SOURCE
				? copied(tableOverrides(table).copy)
				: context.store.table(part.document, part.tables.tblPr[part.aux[src] ?? -1] ?? -1).cellMargins
		style = (table.attrs["styleId"] as string | null) ?? null
	}

	// The style hierarchy, its root first.
	const layers: (CellMarginsProps | undefined)[] = [own]

	for (
		let current = context.resolver.tableStyle(style ?? undefined), depth = 0;
		current !== undefined && depth < 64;
		current = context.styles.parent(current), depth++
	)
		layers.unshift(current.table.cellMargins)

	const margins = resolveCellMargins(layers)

	return Math.max(0, Math.round(marginTwips(margins.start))) + Math.max(0, Math.round(marginTwips(margins.end)))
}

// The narrowest a grid column of a table with cell margins `margins` may get: room for the margins, at least a
// twip.
export function minColumn(margins: number): number {
	return Math.max(1, margins)
}
