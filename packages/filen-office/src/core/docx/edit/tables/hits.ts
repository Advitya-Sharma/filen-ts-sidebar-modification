import type { Node } from "prosemirror-model"
import { LINE_FLOATS, LINE_INTS, LineFloat, LineInt } from "../../../text/index"
import { Item } from "../../model/index"
import { EMPTY_TABLE } from "../../props/index"
import { GRAPHIC_FLOATS, GRAPHIC_INTS, PageGraphic, SOURCE_INTS, SourceFlag, SourceInt, type PageDisplayList } from "../../layout/index"
import type { ObjectHit, ObjectHitProvider } from "../../paged/objects"
import type { PagedEditor } from "../../paged/editor"
import { docxSchema } from "../schema"
import { propertySource, tableMap, type TableMap } from "./map"
import { paddedGrid } from "./ops"

// Column borders of the active story's tables, for dragging a column's width: the vertical border line a
// press is on (within TOLERANCE screen px), as the grid column it ends (`part` "column:<index>", what
// table.column.width takes) at the table's position. A table's left edge (its right edge in a right-to-left
// table) and borderless cells have no border to drag.
//
// The page holds borders and text lines, not cells, so the cell boundary a border stands on comes from the
// text: a border is as tall as the row that drew it, a cell's lines lie between its borders (a cell a vertical
// merge continues has the lines of the merge's first cell) and the row's cells are in grid order (reversed when
// the table is right to left). A border among a cell's lines is a nested table's, or no table's.

const TOLERANCE = 3
// How far a cell's text starts inside its border (Word's default cell margin, px): for placing a boundary
// between cells that have no text to go by.
const MARGIN = 108 / 15
const nodes = docxSchema.nodes

// The table cell a text line is in, at one nesting level.
interface Level {
	readonly tablePos: number
	readonly table: Node
	readonly row: number
	readonly cell: number
}

interface PlacedLine {
	readonly left: number
	readonly right: number
	readonly middle: number
	// Outermost first.
	readonly levels: readonly Level[]
}

interface Range {
	readonly left: number
	readonly right: number
}

// The story's lines in table cells on the page, with the cells they are in.
function placedLines(editor: PagedEditor, page: PageDisplayList): PlacedLine[] {
	const text = page.text
	const doc = editor.session.state.doc
	const out: PlacedLine[] = []

	for (let line = 0; line < text.lineCount; line++) {
		const source = text.lineInts[line * LINE_INTS + LineInt.SOURCE] ?? -1

		if (((page.sourceInts[source * SOURCE_INTS + SourceInt.FLAGS] ?? 0) & SourceFlag.IN_TABLE) === 0) continue

		const f = line * LINE_FLOATS
		const left = text.lineFloats[f + LineFloat.X] ?? 0
		const hit = editor.geometry.atLine(page, line, left)

		if (hit === undefined || hit.atom) continue

		const $pos = doc.resolve(hit.pos)
		const levels: Level[] = []

		for (let depth = 1; depth + 2 <= $pos.depth; depth++) {
			const table = $pos.node(depth)

			if (table.type === nodes.table)
				levels.push({ tablePos: $pos.before(depth), table, row: $pos.index(depth), cell: $pos.index(depth + 1) })
		}

		if (levels.length === 0) continue

		out.push({
			left,
			right: left + (text.lineFloats[f + LineFloat.WIDTH] ?? 0),
			middle: (text.lineFloats[f + LineFloat.TOP] ?? 0) + (text.lineFloats[f + LineFloat.HEIGHT] ?? 0) / 2,
			levels
		})
	}

	return out
}

// Whether a table is laid out right to left (w:bidiVisual), as its own properties or its style say.
export function rightToLeft(editor: PagedEditor, table: Node): boolean {
	const part = editor.session.story.part
	const context = editor.session.context
	const src = propertySource(part, table, Item.Table)
	const direct = src >= 0 ? context.store.table(part.document, part.tables.tblPr[part.aux[src] ?? -1] ?? -1) : EMPTY_TABLE

	if (direct.bidiVisual !== undefined) return direct.bidiVisual

	const styleId = (table.attrs["styleId"] as string | null) ?? undefined
	const { styleId: _source, ...rest } = direct

	return context.resolver.table(styleId === direct.styleId ? direct : styleId === undefined ? rest : { ...rest, styleId }).props
		.bidiVisual
}

// The cell boundary of row `r` of `map` at x `edge`, given the x range of each cell's lines: the number of
// cells before it (1 to the row's cell count); the cell it is inside of; or undefined when none fits.
function boundary(
	editor: PagedEditor,
	map: TableMap,
	r: number,
	ranges: readonly (Range | undefined)[],
	edge: number
): { readonly cells: number } | { readonly inside: number } | undefined {
	const row = map.row(r)
	const count = row.starts.length

	for (let index = 0; index < count; index++) {
		const range = ranges[index]

		if (range !== undefined && range.left < edge - 1 && range.right > edge + 1) return { inside: index }
	}

	const known = ranges.flatMap((range, index) => (range === undefined ? [] : [{ index, ...range }]))
	const first = known[0]
	const last = known[known.length - 1]
	const rtl =
		first !== undefined && last !== undefined && first.index !== last.index ? last.left < first.left : rightToLeft(editor, map.table)
	const before = (range: Range): boolean => (rtl ? range.left >= edge - 1 : range.right <= edge + 1)
	const fits: number[] = []

	for (let k = 1; k <= count; k++) if (known.every(cell => before(cell) === cell.index < k)) fits.push(k)

	if (fits.length <= 1) return fits.length === 0 ? undefined : { cells: fits[0] ?? 1 }

	// Cells without text next to the border: the boundary the grid puts nearest it, placed from the cells with
	// text.
	const offsets = [0]

	for (const width of paddedGrid(map, map.width)) offsets.push((offsets[offsets.length - 1] ?? 0) + width)

	const anchors = known.map(cell => ({
		at: offsets[row.starts[cell.index] ?? 0] ?? 0,
		x: rtl ? cell.right + MARGIN : cell.left - MARGIN
	}))
	const a = anchors[0] ?? { at: 0, x: edge }
	const b = anchors[anchors.length - 1] ?? a
	const slope = b.at !== a.at ? (b.x - a.x) / (b.at - a.at) : rtl ? -1 / 15 : 1 / 15
	let best = fits[0] ?? 1
	let distance = Infinity

	for (const k of fits) {
		const away = Math.abs(a.x + slope * ((offsets[row.starts[k] ?? row.end] ?? 0) - a.at) - edge)

		if (away < distance) {
			distance = away
			best = k
		}
	}

	return { cells: best }
}

// The table and grid column of the border at x `edge` over [top, bottom], among the tables at nesting level
// `level` of the lines `within` keeps.
function columnAt(
	editor: PagedEditor,
	lines: readonly PlacedLine[],
	level: number,
	within: (line: PlacedLine) => boolean,
	edge: number,
	top: number,
	bottom: number
): { readonly tablePos: number; readonly column: number } | undefined {
	const inBand = (line: PlacedLine): boolean => line.middle >= top - 1 && line.middle <= bottom + 1
	const own = lines.filter(line => line.levels.length > level && within(line))
	let table: Level | undefined
	let distance = Infinity

	// The table of the band's line nearest the border.
	for (const line of own) {
		if (!inBand(line)) continue

		const away = edge < line.left ? line.left - edge : edge > line.right ? edge - line.right : 0

		if (away < distance) {
			distance = away
			table = line.levels[level]
		}
	}

	if (table === undefined) return undefined

	const tablePos = table.tablePos
	const map = tableMap(table.table, editor.session.story.part, editor.session.context.store)
	const lineAt = (line: PlacedLine): Level | undefined => line.levels[level]
	const ofTable = own.filter(line => lineAt(line)?.tablePos === tablePos)
	// The row that drew the border: the last with lines in its band (rows above reach into it only with merges).
	let r = -1

	for (const line of ofTable) if (inBand(line)) r = Math.max(r, lineAt(line)?.row ?? -1)

	const row = map.row(r)
	const ranges: (Range | undefined)[] = []

	for (let index = 0; index < row.starts.length; index++) {
		let cellRow = r
		let cell = index

		// A cell a merge continues has the lines of the merge's first cell.
		if (map.continues(r, index)) {
			cellRow = map.chain(r, index).top
			cell = map.cellExactly(cellRow, row.starts[index] ?? 0, row.spans[index] ?? 1)
		}

		let range: Range | undefined

		for (const line of ofTable) {
			const at = lineAt(line)

			if (at?.row !== cellRow || at.cell !== cell) continue

			range = range === undefined ? line : { left: Math.min(range.left, line.left), right: Math.max(range.right, line.right) }
		}

		ranges.push(range === undefined ? undefined : { left: range.left, right: range.right })
	}

	const found = boundary(editor, map, r, ranges, edge)
	const cells = found !== undefined && "cells" in found ? found.cells : 0
	// A border shorter than the row (a nested table's) leaves some of the row's lines out of its band.
	const whole = ofTable.every(line => lineAt(line)?.row !== r || inBand(line))

	if (cells === 0 || !whole) {
		const inside = found !== undefined && "inside" in found ? found.inside : -1
		const nested = columnAt(
			editor,
			ofTable,
			level + 1,
			line => lineAt(line)?.row === r && (inside < 0 || lineAt(line)?.cell === inside),
			edge,
			top,
			bottom
		)

		// Without a nested table there, a row whose exact height cuts its lines off still drew the border.
		if (nested !== undefined || cells === 0) return nested
	}

	const cell = cells - 1

	return { tablePos, column: (row.starts[cell] ?? 0) + (row.spans[cell] ?? 1) - 1 }
}

// `zoom`: screen px per page px (the tolerance is the same on screen at every zoom).
function borderAt(editor: PagedEditor, pageIndex: number, x: number, y: number, zoom = 1): ObjectHit | undefined {
	const page = editor.page(pageIndex)

	if (page === undefined) return undefined

	const floats = page.graphicFloats
	const ints = page.graphicInts
	let border = -1
	let distance = TOLERANCE / (zoom > 0 ? zoom : 1)

	for (let k = 0; k < page.graphicCount; k++) {
		if (ints[k * GRAPHIC_INTS] !== PageGraphic.BORDER) continue

		const left = floats[k * GRAPHIC_FLOATS] ?? 0
		const top = floats[k * GRAPHIC_FLOATS + 1] ?? 0
		const width = floats[k * GRAPHIC_FLOATS + 2] ?? 0
		const height = floats[k * GRAPHIC_FLOATS + 3] ?? 0

		if (height <= width || y < top || y > top + height) continue

		const away = Math.abs(left + width / 2 - x)

		if (away <= distance) {
			distance = away
			border = k
		}
	}

	if (border < 0) return undefined

	const box = {
		x: floats[border * GRAPHIC_FLOATS] ?? 0,
		y: floats[border * GRAPHIC_FLOATS + 1] ?? 0,
		width: floats[border * GRAPHIC_FLOATS + 2] ?? 0,
		height: floats[border * GRAPHIC_FLOATS + 3] ?? 0
	}
	const found = columnAt(editor, placedLines(editor, page), 0, () => true, box.x + box.width / 2, box.y, box.y + box.height)

	return found === undefined
		? undefined
		: { kind: "tableBorder", story: editor.activeStory, pos: found.tablePos, box, part: `column:${String(found.column)}` }
}

export const TABLE_HITS: readonly ObjectHitProvider[] = [{ kind: "tableBorder", hit: borderAt }]
