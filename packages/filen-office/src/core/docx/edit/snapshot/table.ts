import type { Node } from "prosemirror-model"
import { Item, type StoryPart } from "../../model/index"
import { EMPTY_TABLE, type PropsStore, type TableProps } from "../../props/index"
import type { DocxStyles, PropertyResolver } from "../../styles/index"
import type { DocxCellVAlign, DocxTableAlign } from "../commands/tables"
import type { DocxEditContext } from "../context"
import type { DocxEditorSession } from "../session"
import { MAIN, dataAttr } from "../copies"
import { colorKey } from "../formats"
import { NO_SOURCE, type ColorAttrs } from "../schema"
import { cellOverrides, propertySource, tableOverrides, type TableRect } from "../tables/map"
import { canMerge, canSplit, tableAt, type TableAt } from "../tables/ops"
import type { SnapshotInput } from "./index"

// The table the selection is in: its size, the cells the selection touches and what table commands need.
// Read from the rows the selection touches only (never the whole table), and the same object while the
// answer is.

export interface DocxTableStyle {
	readonly id: string
	readonly name: string
}

export interface DocxTableContext {
	readonly rows: number
	// Grid columns (the table's w:tblGrid).
	readonly columns: number
	// The selection's cells as grid rectangle (rows and grid columns, end exclusive).
	readonly fromRow: number
	readonly toRow: number
	readonly fromColumn: number
	readonly toColumn: number
	// The selected cells can be merged, or the cell split.
	readonly mergeable: boolean
	readonly splittable: boolean
	// The first selected row repeats as a header row.
	readonly headerRow: boolean
	readonly styleId: string | null
	// Tables around this one (0: not nested).
	readonly depth: number
	// Grid column widths in twips (what a border drag sends back with table.column.width).
	readonly widths: readonly number[]
	// Direct alignment of the table, null when its style or Word's default decides.
	readonly align: DocxTableAlign | null
	// Laid out right to left (w:bidiVisual): its first column is on the right, so dragging a border to the
	// right narrows the column the border ends.
	readonly rtl: boolean
	// Of the selected cells: vertical alignment and fill ("RRGGBB", or a theme colour as "theme:accent1/tint:51"),
	// "mixed" when they differ, null when none is set on them.
	readonly vAlign: DocxCellVAlign | "mixed" | null
	readonly shading: string | "mixed" | null
	// The document's table styles (for table.style).
	readonly styles: readonly DocxTableStyle[]
}

const previous = new WeakMap<DocxEditorSession, DocxTableContext>()
// Right to left as a table's style decides, per resolver (a document's styles), by the table's direct properties
// and the style a command may have set.
const styleDirections = new WeakMap<PropertyResolver, WeakMap<TableProps, Map<string | undefined, boolean>>>()
const OFF = /^(0|false|off)$/i
const styleLists = new WeakMap<DocxStyles, readonly DocxTableStyle[]>()

function tableStyles(styles: DocxStyles): readonly DocxTableStyle[] {
	let list = styleLists.get(styles)

	if (list === undefined) {
		list = Object.freeze(
			styles.styles
				.filter(style => style.type === "table")
				.map(style => Object.freeze({ id: style.id, name: style.name ?? style.id }))
		)
		styleLists.set(styles, list)
	}

	return list
}

// Whether a table is laid out right to left (w:bidiVisual): as its own properties say (those of the table a copy
// was made from), else its style. Constant time: the style's answer is kept per properties and style.
export function tableRightToLeft(part: StoryPart, context: DocxEditContext, table: Node): boolean {
	const src = propertySource(part, table, Item.Table)

	if (src === NO_SOURCE) {
		const copied = tableOverrides(table).copy?.children.find(entry => entry.uri === MAIN && entry.name === "bidiVisual")
		const value = copied === undefined ? undefined : dataAttr(copied, "val")

		if (copied !== undefined) return value === undefined || !OFF.test(value)
	}

	const direct = src === NO_SOURCE ? EMPTY_TABLE : context.store.table(part.document, part.tables.tblPr[part.aux[src] ?? -1] ?? -1)

	if (direct.bidiVisual !== undefined) return direct.bidiVisual

	const styleId = (table.attrs["styleId"] as string | null) ?? undefined
	const resolver = context.resolver
	let byProps = styleDirections.get(resolver)

	if (byProps === undefined) {
		byProps = new WeakMap()
		styleDirections.set(resolver, byProps)
	}

	let byStyle = byProps.get(direct)

	if (byStyle === undefined) {
		byStyle = new Map()
		byProps.set(direct, byStyle)
	}

	let rtl = byStyle.get(styleId)

	if (rtl === undefined) {
		const { styleId: _source, ...rest } = direct

		rtl = resolver.table(styleId === direct.styleId ? direct : styleId === undefined ? rest : { ...rest, styleId }).props.bidiVisual
		byStyle.set(styleId, rtl)
	}

	return rtl
}

function alignOf(value: string | undefined): DocxTableAlign | null {
	switch (value) {
		case "left":
		case "start":
			return "start"
		case "center":
			return "center"
		case "right":
		case "end":
			return "end"
		default:
			return null
	}
}

function vAlignOf(value: string | undefined): DocxCellVAlign | null {
	return value === "top" || value === "center" || value === "bottom" ? value : null
}

// A cell's vertical alignment and fill: what commands set, else its (or its model's) w:tcPr, or the w:tcPr of
// another part's cell it is a copy of (its `copy`).
function cellValues(part: StoryPart, store: PropsStore, cell: Node): { vAlign: DocxCellVAlign | null; shading: string | null } {
	const overrides = cellOverrides(cell)
	const source = propertySource(part, cell, Item.Cell)
	const props = source === NO_SOURCE ? undefined : store.cell(part.document, part.cells.tcPr[part.aux[source] ?? -1] ?? -1)
	const copied = (name: string, attribute: string): string | undefined => {
		const child = overrides.copy?.children.find(entry => entry.uri === MAIN && entry.name === name)

		return child === undefined ? undefined : dataAttr(child, attribute)
	}
	const tone = (value: string | undefined): number | null => (value === undefined ? null : Number.parseInt(value, 16))
	const own = overrides.shadingTheme
	const color: ColorAttrs | undefined =
		overrides.shading !== undefined
			? overrides.shading === null
				? undefined
				: {
						value: overrides.shading,
						themeColor: own?.theme ?? null,
						themeTint: own?.tint ?? null,
						themeShade: own?.shade ?? null
					}
			: props?.shading?.fill !== undefined
				? {
						value: props.shading.fill.val ?? null,
						themeColor: props.shading.fill.themeColor ?? null,
						themeTint: props.shading.fill.themeTint ?? null,
						themeShade: props.shading.fill.themeShade ?? null
					}
				: copied("shd", "fill") === undefined && copied("shd", "themeFill") === undefined
					? undefined
					: {
							value: copied("shd", "fill") ?? null,
							themeColor: copied("shd", "themeFill") ?? null,
							themeTint: tone(copied("shd", "themeFillTint")),
							themeShade: tone(copied("shd", "themeFillShade"))
						}
	const key = color === undefined ? "auto" : colorKey(color)

	return {
		vAlign: overrides.vAlign ?? vAlignOf(props?.vAlign ?? copied("vAlign", "val")),
		// As the colour pickers key colours: "RRGGBB", or the theme colour ("theme:accent1/tint:51").
		shading: key === "auto" ? null : key
	}
}

function selectedValues(at: TableAt, part: StoryPart, store: PropsStore): Pick<DocxTableContext, "vAlign" | "shading"> {
	const rect: TableRect = at.cells
	let vAlign: DocxCellVAlign | "mixed" | null | undefined
	let shading: string | "mixed" | null | undefined

	for (let r = rect.top; r < rect.bottom; r++) {
		const row = at.map.row(r)

		for (let index = 0; index < row.starts.length; index++) {
			const start = row.starts[index] ?? 0

			if (start >= rect.right || start + (row.spans[index] ?? 1) <= rect.left) continue

			const values = cellValues(part, store, at.map.cell(r, index))

			vAlign = vAlign === undefined || vAlign === values.vAlign ? values.vAlign : "mixed"
			shading = shading === undefined || shading === values.shading ? values.shading : "mixed"
		}
	}

	return { vAlign: vAlign ?? null, shading: shading ?? null }
}

function same(a: DocxTableContext, b: DocxTableContext): boolean {
	for (const key of Object.keys(a) as (keyof DocxTableContext)[]) if (a[key] !== b[key]) return false

	return true
}

export function tableOf(input: SnapshotInput): DocxTableContext | null {
	const { session, state, context } = input
	const part = session.story.part
	const at = tableAt(state, part, context.store)

	if (at === null) return null

	const table = at.table
	const overrides = tableOverrides(table)
	const model = propertySource(part, table, Item.Table)
	const copiedJc = overrides.copy?.children.find(entry => entry.uri === MAIN && entry.name === "jc")
	const sourceAlign =
		model !== NO_SOURCE
			? context.store.table(part.document, part.tables.tblPr[part.aux[model] ?? -1] ?? -1).alignment
			: copiedJc === undefined
				? undefined
				: dataAttr(copiedJc, "val")
	const widths = table.attrs["grid"] as readonly number[]
	const next: DocxTableContext = {
		rows: at.map.height,
		columns: Math.max(widths.length, at.map.row(at.cells.top).end),
		fromRow: at.cells.top,
		toRow: at.cells.bottom,
		fromColumn: at.cells.left,
		toColumn: at.cells.right,
		mergeable: canMerge(at),
		splittable: canSplit(at),
		headerRow: at.map.row(at.cells.top).header,
		styleId: table.attrs["styleId"] as string | null,
		depth: at.depth,
		widths,
		align: overrides.align ?? alignOf(sourceAlign),
		rtl: tableRightToLeft(part, context, table),
		...selectedValues(at, part, context.store),
		styles: tableStyles(context.styles)
	}
	const last = previous.get(session)

	if (last !== undefined && same(last, next)) return last

	const frozen = Object.freeze(next)

	previous.set(session, frozen)

	return frozen
}
