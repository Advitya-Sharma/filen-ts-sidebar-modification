import { Fragment, type ContentMatch, type Mark, type Node, type NodeType } from "prosemirror-model"
import { closeHistory } from "prosemirror-history"
import { TextSelection, type EditorState, type Transaction } from "prosemirror-state"
import { Item, RevisionFlag, type StoryPart } from "../../model/index"
import type { PropsStore } from "../../props/index"
import { minuteDate, paragraphChange, type ParagraphChange } from "../../review/track"
import { MAX_TABLE_COLUMNS, MAX_TABLE_ROWS, type DocxCellVAlign, type DocxTableAlign } from "../commands/tables"
import type { DocxEditContext } from "../context"
import { markCommand } from "../invariants/core"
import { NO_MARKERS, NO_SOURCE, docxSchema, type MarkerList } from "../schema"
import {
	cellOverrides,
	propertySource,
	rowOverrides,
	tableMap,
	tableOverrides,
	type CellOverrides,
	type MapRow,
	type RowOverrides,
	type TableMap,
	type TableOverrides,
	type TableRect,
	type TrackStamp
} from "./map"

// Table structure and property edits as transactions of a story's session. Edits of a few rows are steps
// on those rows, made from the end of the table backwards (positions read before stay valid); edits of
// every row (columns, widths) rebuild the table in one step, keeping the nodes of what did not change. Rows
// and cells an edit adds are new (no source), modelled on a neighbour whose properties they copy (their
// row_props / cell_props `from`). A cell's width in twips (w:tcW) is its `colwidth` once an edit changed it.
// While changes are tracked, inserted and deleted rows and cells are recorded (revision bits with a stamp)
// instead of being removed; merges, splits and property changes apply untracked.

const nodes = docxSchema.nodes
const marks = docxSchema.marks

// The narrowest grid column an edit leaves (twips).
export const MIN_COLUMN = 115
// Grid column width when a table has no grid to go by (twips, 1").
const DEFAULT_COLUMN = 1440
// Text width when nothing says better (twips: Letter with 1" margins).
const DEFAULT_TEXT_WIDTH = 9360
// A cell's default left and right margins together (twips).
const CELL_MARGINS = 216

// A command's transaction (it records its own tracked changes), an undo step of its own.
function finish(tr: Transaction): Transaction {
	return closeHistory(markCommand(tr))
}

// ---- where the selection is -----------------------------------------------------------------------------

export interface TableAt {
	readonly table: Node
	readonly pos: number
	readonly map: TableMap
	// The cells at the selection's ends (row, index).
	readonly anchor: { readonly row: number; readonly index: number }
	readonly head: { readonly row: number; readonly index: number }
	// The rows and grid columns the selection touches, and that rectangle grown to whole merged cells.
	readonly cells: TableRect
	readonly rect: TableRect
	// Tables around it.
	readonly depth: number
}

const located = new WeakMap<EditorState, TableAt | null>()

// The innermost table holding both ends of the selection, or null.
export function tableAt(state: EditorState, part: StoryPart, store: PropsStore): TableAt | null {
	const cached = located.get(state)

	if (cached !== undefined) return cached

	const found = locate(state, part, store)

	located.set(state, found)

	return found
}

function locate(state: EditorState, part: StoryPart, store: PropsStore): TableAt | null {
	const { $from, $to } = state.selection

	for (let d = Math.min($from.depth, $to.depth) - 2; d >= 1; d--) {
		const table = $from.node(d)

		if (table.type !== nodes.table || $to.node(d) !== table || $from.before(d) !== $to.before(d)) continue

		const pos = $from.before(d)
		const map = tableMap(table, part, store)
		const anchor = { row: $from.index(d), index: $from.index(d + 1) }
		const head = { row: $to.index(d), index: $to.index(d + 1) }
		const a = map.row(anchor.row)
		const b = map.row(head.row)
		const aStart = a.starts[anchor.index] ?? 0
		const bStart = b.starts[head.index] ?? 0
		const cells = {
			top: Math.min(anchor.row, head.row),
			bottom: Math.max(anchor.row, head.row) + 1,
			left: Math.min(aStart, bStart),
			right: Math.max(aStart + (a.spans[anchor.index] ?? 1), bStart + (b.spans[head.index] ?? 1))
		}
		let depth = 0

		for (let k = d - 1; k >= 1; k--) if ($from.node(k).type === nodes.table) depth++

		return { table, pos, map, anchor, head, cells, rect: map.normalize(cells), depth }
	}

	return null
}

// Whether the selection's cells can be merged: more than one cell, every row of the area covered by cells.
export function canMerge(at: TableAt): boolean {
	return at.map.areasIn(at.rect) > 1 && tiles(at.map, at.rect)
}

// Whether the selection is one cell (a merged one counts as one), which a split divides.
export function canSplit(at: TableAt): boolean {
	return at.map.areasIn(at.rect) === 1 && tiles(at.map, at.rect)
}

// Every row of `rect` has cells covering exactly its columns.
function tiles(map: TableMap, rect: TableRect): boolean {
	for (let r = rect.top; r < rect.bottom; r++) {
		const row = map.row(r)
		let covered = 0

		for (let index = 0; index < row.starts.length; index++) {
			const start = row.starts[index] ?? 0
			const end = start + (row.spans[index] ?? 1)

			if (end <= rect.left || start >= rect.right) continue
			if (start < rect.left || end > rect.right) return false

			covered += end - start
		}

		if (covered !== rect.right - rect.left) return false
	}

	return true
}

// ---- tracking -------------------------------------------------------------------------------------------

// The stamp edits are recorded with: undefined while changes are not tracked, null when they are but there
// is no author to record (the command is refused).
export function trackStamp(context: DocxEditContext): TrackStamp | null | undefined {
	if (!context.review.tracking) return undefined

	const author = context.identity.author

	if (author === undefined || author === "") return null

	return { author, date: minuteDate(context.identity.now()), id: context.ids.next("revision") }
}

// ---- building blocks ------------------------------------------------------------------------------------

function withoutKeys<T extends object>(value: T, drop: readonly string[]): Partial<T> {
	const out: Record<string, unknown> = {}

	for (const [key, entry] of Object.entries(value)) if (!drop.includes(key) && entry !== undefined) out[key] = entry

	return out as Partial<T>
}

function sameValue(a: object, b: object): boolean {
	return JSON.stringify(withoutKeys(a, [])) === JSON.stringify(withoutKeys(b, []))
}

// `marks` with the row_props / cell_props mark of `type` holding `value` (none when empty).
function withValue(type: Mark["type"], current: readonly Mark[], value: object): readonly Mark[] {
	const clean = withoutKeys(value, [])

	return Object.keys(clean).length === 0 ? type.removeFromSet(current) : type.create({ value: clean }).addToSet(current)
}

// Sets a row's or cell's overrides to `value` (none when empty), as a step.
function setOverrides(tr: Transaction, pos: number, node: Node, value: object): void {
	const type = node.type === nodes.table_row ? marks.row_props : marks.cell_props
	const clean = withoutKeys(value, [])

	if (Object.keys(clean).length === 0) {
		if (type.isInSet(node.marks) !== undefined) tr.removeNodeMark(pos, type)
	} else {
		tr.addNodeMark(pos, type.create({ value: clean }))
	}
}

// A row with other overrides (as a node).
function rowWith(row: Node, value: RowOverrides): Node {
	return sameValue(rowOverrides(row), value) ? row : row.mark(withValue(marks.row_props, row.marks, value))
}

// A cell with another span or width (as a node).
function cellWith(
	cell: Node,
	attrs: { readonly colspan?: number; readonly colwidth?: number | null; readonly vMerge?: string | null }
): Node {
	const next: Record<string, unknown> = { ...cell.attrs, ...attrs }

	return Object.keys(attrs).every(key => next[key] === cell.attrs[key]) ? cell : cell.type.create(next, cell.content, cell.marks)
}

// A new paragraph formatted as `model` (the first paragraph of the cell a new cell is modelled on).
function emptyParagraph(model: Node | null | undefined): Node {
	const attrs = model?.type === nodes.paragraph ? model.attrs : undefined

	return nodes.paragraph.create({
		styleId: attrs?.["styleId"] ?? null,
		align: attrs?.["align"] ?? null,
		indent: attrs?.["indent"] ?? null
	})
}

// A new cell modelled on `model` (its properties copied), `span` grid columns and `colwidth` twips wide.
function newCell(
	part: StoryPart,
	model: Node | undefined,
	span: number,
	vMerge: string | null,
	colwidth: number | null,
	stamp: TrackStamp | undefined
): Node {
	const overrides: Record<string, unknown> = model === undefined ? {} : withoutKeys(cellOverrides(model), ["from", "ins", "del"])
	const source = model === undefined ? NO_SOURCE : propertySource(part, model, Item.Cell)

	if (source !== NO_SOURCE) overrides["from"] = source
	if (stamp !== undefined) overrides["ins"] = stamp

	return nodes.table_cell.create(
		{ colspan: span, vMerge, colwidth, revision: stamp === undefined ? 0 : RevisionFlag.Inserted },
		emptyParagraph(model?.firstChild),
		withValue(marks.cell_props, [], overrides)
	)
}

// A new row modelled on `model` (properties copied) holding `cells`.
function newRow(part: StoryPart, model: Node | undefined, cells: readonly Node[], stamp: TrackStamp | undefined): Node {
	const overrides: Record<string, unknown> = model === undefined ? {} : withoutKeys(rowOverrides(model), ["from", "ins", "del"])
	const source = model === undefined ? NO_SOURCE : propertySource(part, model, Item.Row)

	if (source !== NO_SOURCE) overrides["from"] = source
	if (stamp !== undefined) overrides["ins"] = stamp

	return nodes.table_row.create(
		{ revision: stamp === undefined ? 0 : RevisionFlag.Inserted },
		cells,
		withValue(marks.row_props, [], overrides)
	)
}

// A cell holding nothing but one empty paragraph: what a merge leaves out.
function emptyCell(cell: Node): boolean {
	const only = cell.childCount === 1 ? cell.firstChild : null

	return (
		only !== null &&
		only.type === nodes.paragraph &&
		only.content.size === 0 &&
		only.marks.length === 0 &&
		(only.attrs["before"] as MarkerList).length === 0 &&
		only.attrs["sectPr"] === NO_SOURCE
	)
}

function caretIn(tr: Transaction, pos: number): Transaction {
	const at = Math.max(0, Math.min(pos, tr.doc.content.size))

	return tr.setSelection(TextSelection.near(tr.doc.resolve(at)))
}

// The end of row `r` (the position after it).
function rowEnd(map: TableMap, pos: number, r: number): number {
	return map.rowPos(pos, r) + map.table.child(r).nodeSize
}

function cellEnd(map: TableMap, pos: number, r: number, index: number): number {
	return map.cellPos(pos, r, index) + map.cell(r, index).nodeSize
}

// Grid widths padded to `width` columns.
export function paddedGrid(map: TableMap, width: number): number[] {
	const grid = [...map.grid]
	const known = grid.filter(value => value > 0)
	const fill = known.length === 0 ? DEFAULT_COLUMN : Math.round(known.reduce((sum, value) => sum + value, 0) / known.length)

	for (let c = 0; c < grid.length; c++) if (!((grid[c] ?? 0) > 0)) grid[c] = fill

	while (grid.length < width) grid.push(fill)

	return grid
}

function widthOf(grid: readonly number[], start: number, span: number): number {
	let width = 0

	for (let c = start; c < start + span; c++) width += grid[c] ?? 0

	return width
}

// `widths` scaled to add up to `total`, as integers.
function scaled(widths: readonly number[], total: number): number[] {
	const sum = widths.reduce((acc, value) => acc + value, 0)

	if (sum <= 0) return widths.map(() => Math.max(MIN_COLUMN, Math.floor(total / Math.max(1, widths.length))))

	const out = widths.map(value => Math.max(MIN_COLUMN, Math.round((value * total) / sum)))
	const drift = total - out.reduce((acc, value) => acc + value, 0)
	const last = out.length - 1

	if (last >= 0) out[last] = Math.max(MIN_COLUMN, (out[last] ?? 0) + drift)

	return out
}

// Replaces the table with one made of `rows` (the old nodes where nothing changed) and `attrs`; returns where
// cell `index` of row `r` of it starts.
function replaceTable(tr: Transaction, at: TableAt, rows: readonly Node[], attrs: object): (r: number, index: number) => number {
	const table = at.table.type.create({ ...at.table.attrs, ...attrs }, Fragment.fromArray([...rows]), at.table.marks)

	tr.replaceWith(at.pos, at.pos + at.table.nodeSize, table)

	return (r, index) => {
		let pos = at.pos + 1

		for (let k = 0; k < Math.min(r, table.childCount - 1); k++) pos += table.child(k).nodeSize

		const row = table.child(Math.min(r, table.childCount - 1))

		pos++

		for (let k = 0; k < Math.min(index, row.childCount - 1); k++) pos += row.child(k).nodeSize

		return pos
	}
}

// The flags of a paragraph mark's own revision.
const MARK_FLAGS = RevisionFlag.Inserted | RevisionFlag.Deleted | RevisionFlag.MovedFrom | RevisionFlag.MovedTo

// `nodeMarks` with a paragraph_change mark holding `change` (none when it records nothing).
function withChange(nodeMarks: readonly Mark[], change: ParagraphChange): readonly Mark[] {
	const others = marks.paragraph_change.removeFromSet(nodeMarks)

	return Object.keys(change).length === 0 ? others : marks.paragraph_change.create({ value: change }).addToSet(others)
}

// ---- insert and delete a table --------------------------------------------------------------------------

// Where the selection allows a new table: the paragraph at its start, in a container blocks go into.
export function canInsertTable(state: EditorState): boolean {
	const $from = state.selection.$from

	if ($from.parent.type !== nodes.paragraph || $from.depth < 1) return false

	const container = $from.node($from.depth - 1)

	if (tablesAnywhere(container.type)) return true

	const index = $from.index($from.depth - 1)

	return container.canReplaceWith(index, index, nodes.table)
}

const anywhere = new Map<NodeType, boolean>()

// Whether a container's content takes a table before any of its blocks (a repetition of blocks: matching its
// content from the start, linear in its children, is not needed on every snapshot).
function tablesAnywhere(type: NodeType): boolean {
	let known = anywhere.get(type)

	if (known === undefined) {
		const start = type.contentMatch
		const loop = start.matchType(nodes.table)

		// Every block leads from the start, and from after a block, to the same state, which takes a table.
		const same = (match: ContentMatch | null): boolean => match === null || match === loop

		known =
			loop !== null &&
			loop.matchType(nodes.table) === loop &&
			Object.values(nodes).every(other => same(start.matchType(other)) && same(loop.matchType(other)))
		anywhere.set(type, known)
	}

	return known
}

// The width a new table fills: the enclosing cell's, else the section's text width (twips).
function availableWidth(state: EditorState, context: DocxEditContext, part: StoryPart, store: PropsStore): number {
	const $from = state.selection.$from

	for (let d = $from.depth - 1; d >= 2; d--) {
		if ($from.node(d).type !== nodes.table_cell) continue

		const cell = $from.node(d)
		const map = tableMap($from.node(d - 2), part, store)
		const row = map.row($from.index(d - 2))
		const index = $from.index(d - 1)
		const start = row.starts[index] ?? 0
		const span = row.spans[index] ?? 1
		const width = (cell.attrs["colwidth"] as number | null) ?? widthOf(paddedGrid(map, start + span), start, span)

		return Math.max(MIN_COLUMN, width - CELL_MARGINS)
	}

	const model = context.model
	const body = model.body === undefined ? undefined : model.parts[model.body.part]
	const section = model.sections[model.sections.length - 1]

	if (body === undefined || section === undefined || section.sectPr < 0) return DEFAULT_TEXT_WIDTH

	const props = store.section(body.document, section.sectPr)
	const page = props.pageSize?.width
	const margins = props.pageMargins

	if (page === undefined) return DEFAULT_TEXT_WIDTH

	const columns = Math.max(1, props.columns?.count ?? 1)
	const text = page - (margins?.left ?? 1440) - (margins?.right ?? 1440) - (margins?.gutter ?? 0)
	const space = props.columns?.space ?? 720

	return Math.max(MIN_COLUMN * 2, Math.round((text - space * (columns - 1)) / columns))
}

// `joined`: the transaction whose undo step it joins (the selection the table replaces, deleted just before).
export function insertTable(
	state: EditorState,
	context: DocxEditContext,
	part: StoryPart,
	rows: number,
	cols: number,
	joined?: Transaction
): Transaction | null {
	if (!canInsertTable(state) || rows < 1 || cols < 1 || rows > MAX_TABLE_ROWS || cols > MAX_TABLE_COLUMNS) return null

	const stamp = trackStamp(context)

	if (stamp === null) return null

	const column = Math.max(MIN_COLUMN, Math.floor(availableWidth(state, context, part, context.store) / cols))
	const gridStyle = context.styles.byName("Table Grid", "table")
	const props: TableOverrides = gridStyle === undefined ? { borders: true, look: true } : { look: true }
	const tableRows: Node[] = []

	for (let r = 0; r < rows; r++) {
		const cells = Array.from({ length: cols }, () =>
			nodes.table_cell.create({ colspan: 1, colwidth: column }, nodes.paragraph.create())
		)

		tableRows.push(
			nodes.table_row.create(
				{ revision: stamp === undefined ? 0 : RevisionFlag.Inserted },
				cells,
				stamp === undefined ? [] : [marks.row_props.create({ value: { ins: stamp } })]
			)
		)
	}

	const table = nodes.table.create({ styleId: gridStyle?.id ?? null, grid: Array.from({ length: cols }, () => column), props }, tableRows)
	const tr = state.tr
	const $from = state.selection.$from
	const paragraph = $from.parent
	let at: number

	if (paragraph.content.size === 0 || $from.parentOffset === 0) {
		at = $from.before()
		tr.insert(at, table)
	} else {
		// Split at the caret, the paragraph replaced by its halves around the table in one step. The first half
		// ends with a new paragraph mark (inserted by the change while tracking; rejecting it joins the halves
		// again once the table is gone) and keeps the paragraph's own marks; the original mark, with its revision,
		// its section and a section change, ends the second.
		const from = $from.before()
		const change = paragraphChange(paragraph)
		const revision = paragraph.attrs["revision"] as number
		const own = withoutKeys(change ?? {}, ["inserted", "deleted"])
		const first = paragraph.type.create(
			{
				...paragraph.attrs,
				sectPr: NO_SOURCE,
				revision: (revision & ~MARK_FLAGS) | (stamp === undefined ? 0 : RevisionFlag.Inserted)
			},
			paragraph.content.cut(0, $from.parentOffset),
			withChange(
				paragraph.marks.filter(mark => mark.type !== marks.section_change),
				stamp === undefined ? own : { ...own, inserted: stamp }
			)
		)
		const second = paragraph.type.create(
			{ ...paragraph.attrs, before: NO_MARKERS },
			paragraph.content.cut($from.parentOffset),
			withChange(
				paragraph.marks.filter(mark => mark.type === marks.section_change),
				withoutKeys(change ?? {}, ["props", "source"])
			)
		)

		tr.replaceWith(from, from + paragraph.nodeSize, [first, table, second])
		at = from + first.nodeSize
	}

	caretIn(tr, at + 4)

	// One undo step with the deletion before it.
	return joined === undefined ? finish(tr) : markCommand(tr.setMeta("appendedTransaction", joined))
}

export function deleteTable(state: EditorState, context: DocxEditContext, at: TableAt): Transaction | null {
	const stamp = trackStamp(context)

	if (stamp === null) return null

	return finish(stamp === undefined ? removeTable(state.tr, at) : trackRowDeletion(state.tr, at, 0, at.map.height, stamp))
}

function removeTable(tr: Transaction, at: TableAt): Transaction {
	const end = at.pos + at.table.nodeSize

	if (tr.doc.resolve(at.pos).parent.childCount === 1) tr.replaceWith(at.pos, end, nodes.paragraph.create())
	else tr.delete(at.pos, end)

	return caretIn(tr, at.pos)
}

// Takes out the tables at `positions` (the editor's own, after a save refused them), as one step.
export function removeNewTables(state: EditorState, positions: readonly number[]): Transaction | null {
	const tr = state.tr

	for (const pos of [...positions].sort((a, b) => b - a)) {
		const table = tr.doc.nodeAt(pos)

		if (table === null || table.type !== nodes.table) continue

		if (tr.doc.resolve(pos).parent.childCount === 1) tr.replaceWith(pos, pos + table.nodeSize, nodes.paragraph.create())
		else tr.delete(pos, pos + table.nodeSize)
	}

	return tr.docChanged ? finish(tr) : null
}

// ---- rows -----------------------------------------------------------------------------------------------

export function insertRows(state: EditorState, context: DocxEditContext, part: StoryPart, at: TableAt, below: boolean): Transaction | null {
	const stamp = trackStamp(context)

	if (stamp === null) return null

	const { map, pos, cells: selected } = at
	const count = selected.bottom - selected.top

	if (map.height + count > MAX_TABLE_ROWS) return null

	const ref = below ? selected.bottom - 1 : selected.top
	const refRow = map.row(ref)
	const made: Node[] = []

	for (let k = 0; k < count; k++) {
		const cells: Node[] = []

		for (let index = 0; index < refRow.starts.length; index++) {
			const model = map.cell(ref, index)
			let vMerge: string | null = null

			// Inside a vertical merge, the new cell continues it.
			if (below) {
				const next = ref + 1 < map.height ? map.cellExactly(ref + 1, refRow.starts[index] ?? 0, refRow.spans[index] ?? 1) : -1

				if (next >= 0 && map.continues(ref + 1, next)) vMerge = "continue"
			} else if (map.continues(ref, index)) {
				vMerge = "continue"
			}

			cells.push(newCell(part, model, refRow.spans[index] ?? 1, vMerge, model.attrs["colwidth"] as number | null, stamp))
		}

		made.push(newRow(part, refRow.node, cells, stamp))
	}

	const tr = state.tr
	const insertAt = below ? rowEnd(map, pos, selected.bottom - 1) : map.rowPos(pos, selected.top)

	tr.insert(insertAt, made)

	return finish(caretIn(tr, insertAt + 3))
}

// A row the author inserted in this editor while changes were tracked: deleting it takes it out for good.
function ownInsertion(row: Node, author: string): boolean {
	return ((row.attrs["revision"] as number) & RevisionFlag.Inserted) !== 0 && rowOverrides(row).ins?.author === author
}

// Rows [top, bottom) deleted as a tracked change: each is marked deleted, except rows the author inserted while
// tracking, which go (the whole table when nothing else is left).
function trackRowDeletion(tr: Transaction, at: TableAt, top: number, bottom: number, stamp: TrackStamp): Transaction {
	const { map, pos } = at
	let left = map.height

	for (let r = top; r < bottom; r++) if (ownInsertion(map.table.child(r), stamp.author)) left--

	if (left === 0) return removeTable(tr, at)

	for (let r = bottom - 1; r >= top; r--) {
		const row = map.table.child(r)

		if (ownInsertion(row, stamp.author)) {
			// A run of them at once.
			let first = r

			while (first > top && ownInsertion(map.table.child(first - 1), stamp.author)) first--

			removeRows(tr, map, pos, first, r + 1)
			r = first

			continue
		}

		const rowPos = map.rowPos(pos, r)
		const revision = row.attrs["revision"] as number

		if ((revision & RevisionFlag.Deleted) === 0) {
			tr.setNodeAttribute(rowPos, "revision", revision | RevisionFlag.Deleted)
			setOverrides(tr, rowPos, row, { ...rowOverrides(row), del: stamp })
		}
	}

	return caretIn(tr, tr.mapping.map(map.rowPos(pos, top)) + 3)
}

// Takes rows [top, bottom) out of the table (not all of them), as steps. A vertical merge starting in them
// that goes on below is held by the first row left, with its content.
export function removeRows(tr: Transaction, map: TableMap, pos: number, top: number, bottom: number): void {
	if (bottom < map.height) {
		const row = map.row(bottom)

		for (let index = row.starts.length - 1; index >= 0; index--) {
			if (!map.continues(bottom, index)) continue

			const chain = map.chain(bottom, index)

			if (chain.top < top) continue

			const restart = map.cellExactly(chain.top, row.starts[index] ?? 0, row.spans[index] ?? 1)
			const cellPos = map.cellPos(pos, bottom, index)
			const cell = map.cell(bottom, index)
			const content: Node[] = []

			if (restart >= 0) map.cell(chain.top, restart).forEach(block => content.push(block))
			if (restart < 0 || !emptyCell(cell)) cell.forEach(block => content.push(block))

			tr.replaceWith(
				cellPos,
				cellPos + cell.nodeSize,
				cell.type.create({ ...cell.attrs, vMerge: chain.bottom - bottom > 1 ? "restart" : null }, content, cell.marks)
			)
		}
	}

	tr.delete(map.rowPos(pos, top), rowEnd(map, pos, bottom - 1))
}

export function deleteRows(state: EditorState, context: DocxEditContext, at: TableAt): Transaction | null {
	const stamp = trackStamp(context)

	if (stamp === null) return null

	const { map, pos } = at
	const { top, bottom } = at.cells
	const tr = state.tr

	if (stamp !== undefined) return finish(trackRowDeletion(tr, at, top, bottom, stamp))
	if (top === 0 && bottom === map.height) return finish(removeTable(tr, at))

	removeRows(tr, map, pos, top, bottom)

	return finish(caretIn(tr, (bottom < map.height ? map.rowPos(pos, top) : map.rowPos(pos, top - 1)) + 3))
}

// ---- columns --------------------------------------------------------------------------------------------

// The widths of the columns an insertion adds: one per selected cell of the row with the most of them, as wide
// as that cell (Word inserts as many columns as are selected, each as wide as the one it copies; a cell
// spanning grid columns is one column).
function addedWidths(map: TableMap, rect: TableRect, grid: readonly number[]): number[] {
	let widest: number[] = []

	for (let r = rect.top; r < rect.bottom; r++) {
		const row = map.row(r)
		const widths: number[] = []

		for (let index = 0; index < row.starts.length; index++) {
			const start = row.starts[index] ?? 0
			const span = row.spans[index] ?? 1

			if (start >= rect.left && start + span <= rect.right) widths.push(widthOf(grid, start, span))
		}

		if (widths.length > widest.length) widest = widths
	}

	return widest.length > 0 ? widest : [widthOf(grid, rect.left, rect.right - rect.left)]
}

export function insertColumns(
	state: EditorState,
	context: DocxEditContext,
	part: StoryPart,
	at: TableAt,
	right: boolean
): Transaction | null {
	const stamp = trackStamp(context)

	if (stamp === null) return null

	const { map, cells: selected } = at
	const width = map.width
	const old = paddedGrid(map, width)
	const added = addedWidths(map, selected, old)
	const count = added.length

	if (width + count > MAX_TABLE_COLUMNS) return null

	const boundary = right ? selected.right : selected.left
	const refColumn = right ? selected.right - 1 : selected.left
	const total = old.reduce((sum, value) => sum + value, 0)
	const grid = scaled([...old.slice(0, boundary), ...added, ...old.slice(boundary)], total)
	// Old grid column -> new.
	const moved = (column: number): number => (column < boundary ? column : column + count)
	const rows: Node[] = []
	let caret = -1

	for (let r = 0; r < map.height; r++) {
		const row = map.row(r)
		const node = row.node
		const cells: Node[] = []
		let overrides: RowOverrides = rowOverrides(node)
		let insertAt = -1

		if (boundary < row.before) overrides = { ...overrides, gridBefore: row.before + count }
		else if (boundary > row.end) {
			if (boundary <= row.end + row.after) overrides = { ...overrides, gridAfter: row.after + count }
		} else {
			const covering = map.cellAt(r, boundary)

			insertAt = covering >= 0 && (row.starts[covering] ?? 0) < boundary ? -1 : covering >= 0 ? covering : row.starts.length
		}

		for (let index = 0; index <= row.starts.length; index++) {
			if (index === insertAt) {
				const modelIndex = map.cellAt(r, refColumn) >= 0 ? map.cellAt(r, refColumn) : Math.min(index, row.starts.length - 1)
				const model = modelIndex >= 0 ? map.cell(r, modelIndex) : undefined

				if (r === at.anchor.row) caret = cells.length

				for (let k = 0; k < count; k++) cells.push(newCell(part, model, 1, null, grid[boundary + k] ?? DEFAULT_COLUMN, stamp))
			}

			if (index === row.starts.length) break

			const start = row.starts[index] ?? 0
			const span = row.spans[index] ?? 1
			const grows = start < boundary && boundary < start + span
			const newSpan = grows ? span + count : span

			if (grows && r === at.anchor.row) caret = cells.length

			cells.push(cellWith(map.cell(r, index), { colspan: newSpan, colwidth: widthOf(grid, moved(start), newSpan) }))
		}

		rows.push(rowWith(node.copy(Fragment.fromArray(cells)), overrides))
	}

	const tr = state.tr
	const cellPos = replaceTable(tr, at, rows, { grid })

	return finish(caretIn(tr, cellPos(at.anchor.row, Math.max(0, caret)) + 2))
}

export function deleteColumns(state: EditorState, context: DocxEditContext, at: TableAt): Transaction | null {
	const stamp = trackStamp(context)

	if (stamp === null) return null

	const { map } = at
	const { left, right } = at.cells
	const tr = state.tr

	// Tracked: the cells inside the columns are marked deleted (w:cellDel, which Word reads as a cell deletion
	// to accept or reject, though its own column deletion asks and applies untracked). A cell holds one tracked
	// change at most (CT_TcPr), so columns of cells the author inserted while tracking go for good, and inserted
	// cells of another kind keep the columns from being deleted as a tracked change.
	const inside = (r: number, index: number): boolean => {
		const start = map.row(r).starts[index] ?? 0

		return start >= left && start + (map.row(r).spans[index] ?? 1) <= right
	}
	let cells = 0
	let own = 0
	let inserted = 0

	if (stamp !== undefined) {
		for (let r = 0; r < map.height; r++) {
			for (let index = 0; index < map.row(r).starts.length; index++) {
				if (!inside(r, index)) continue

				const cell = map.cell(r, index)

				cells++

				if (((cell.attrs["revision"] as number) & RevisionFlag.Inserted) === 0) continue

				inserted++

				if (cellOverrides(cell).ins?.author === stamp.author) own++
			}
		}

		if (inserted > 0 && own < cells) return null
	}

	if (stamp !== undefined && inserted === 0) {
		const rows: Node[] = []

		for (let r = 0; r < map.height; r++) {
			const row = map.row(r)
			const cells: Node[] = []

			for (let index = 0; index < row.starts.length; index++) {
				const cell = map.cell(r, index)
				const revision = cell.attrs["revision"] as number

				if (!inside(r, index) || (revision & RevisionFlag.Deleted) !== 0) {
					cells.push(cell)

					continue
				}

				cells.push(
					cell.type.create(
						{ ...cell.attrs, revision: revision | RevisionFlag.Deleted },
						cell.content,
						withValue(marks.cell_props, cell.marks, { ...cellOverrides(cell), del: stamp })
					)
				)
			}

			rows.push(cells.every((cell, index) => cell === row.node.child(index)) ? row.node : row.node.copy(Fragment.fromArray(cells)))
		}

		replaceTable(tr, at, rows, {})

		return finish(caretIn(tr, at.pos + 3))
	}

	if (left === 0 && right >= map.width) return finish(removeTable(tr, at))

	const grid = [...map.grid]

	grid.splice(left, Math.max(0, Math.min(right, grid.length) - left))

	const count = right - left
	const overlap = (from: number, to: number): number => Math.max(0, Math.min(to, right) - Math.max(from, left))
	const rows: Node[] = []

	for (let r = 0; r < map.height; r++) {
		const row = map.row(r)
		const cells: Node[] = []

		for (let index = 0; index < row.starts.length; index++) {
			const start = row.starts[index] ?? 0
			const span = row.spans[index] ?? 1
			const cut = overlap(start, start + span)
			const cell = map.cell(r, index)

			if (cut === span) continue

			const newStart = start < left ? start : start >= right ? start - count : left

			cells.push(cut === 0 ? cell : cellWith(cell, { colspan: span - cut, colwidth: widthOf(grid, newStart, span - cut) }))
		}

		// A row whose cells all go goes with them.
		if (cells.length === 0) continue

		const before = row.before - overlap(0, row.before)
		const after = row.after - overlap(row.end, row.end + row.after)
		let node =
			cells.length === row.starts.length && cells.every((cell, index) => cell === row.node.child(index))
				? row.node
				: row.node.copy(Fragment.fromArray(cells))

		if (before !== row.before || after !== row.after)
			node = rowWith(node, { ...rowOverrides(node), gridBefore: before, gridAfter: after })

		rows.push(node)
	}

	if (rows.length === 0) return finish(removeTable(tr, at))

	const cellPos = replaceTable(tr, at, rows, { grid })

	return finish(caretIn(tr, cellPos(Math.min(at.anchor.row, rows.length - 1), 0) + 2))
}

// ---- merge and split ------------------------------------------------------------------------------------

export function mergeCells(state: EditorState, at: TableAt): Transaction | null {
	if (!canMerge(at)) return null

	const { map, pos, rect } = at
	const span = rect.right - rect.left
	const width = widthOf(paddedGrid(map, rect.right), rect.left, span)
	const tall = rect.bottom - rect.top > 1
	const moved: Node[] = []
	const markers: MarkerList[] = []
	const groups: { row: number; first: number; last: number }[] = []

	for (let r = rect.top; r < rect.bottom; r++) {
		const row = map.row(r)
		let first = -1
		let last = -1

		for (let index = 0; index < row.starts.length; index++) {
			const start = row.starts[index] ?? 0

			if (start < rect.left || start >= rect.right) continue
			if (first < 0) first = index

			last = index
		}

		groups.push({ row: r, first, last })

		for (let index = first; index <= last; index++) {
			const cell = map.cell(r, index)

			if (r === rect.top && index === first) continue

			// Markers between the cells that go stay with the merged cell.
			if (r === rect.top || index > first) markers.push(cell.attrs["before"] as MarkerList, cell.attrs["end"] as MarkerList)
			if (!emptyCell(cell)) cell.forEach(block => moved.push(block))
		}
	}

	const tr = state.tr

	for (let k = groups.length - 1; k >= 0; k--) {
		const { row: r, first, last } = groups[k] as { row: number; first: number; last: number }
		const keep = map.cell(r, first)
		let content: Node[]
		let end = keep.attrs["end"] as MarkerList

		if (r === rect.top) {
			content = []
			keep.forEach(block => content.push(block))
			content.push(...moved)

			if (content.at(-1)?.type === nodes.table) content.push(nodes.paragraph.create())

			const kept = markers.flat()

			if (kept.length > 0) end = [...end, ...kept]
		} else {
			content = [emptyParagraph(keep.firstChild)]
		}

		tr.replaceWith(
			map.cellPos(pos, r, first),
			cellEnd(map, pos, r, last),
			keep.type.create(
				{ ...keep.attrs, colspan: span, colwidth: width, vMerge: r === rect.top ? (tall ? "restart" : null) : "continue", end },
				content,
				keep.marks
			)
		)

		// The rows a merge spans repeat as header rows or not, as its first row does.
		const header = map.row(rect.top).header

		if (map.row(r).header !== header)
			setOverrides(tr, map.rowPos(pos, r), map.row(r).node, { ...rowOverrides(map.row(r).node), header })
	}

	return finish(caretIn(tr, map.cellPos(pos, rect.top, groups[0]?.first ?? 0) + 2))
}

// Splits the selected cell into `rows` x `cols` cells of equal width. A cell merged over several rows splits
// into a divisor of them; a cell of one row into as many rows as asked (new rows, the other cells of its row
// merged down into them), as Word does.
export function splitCell(state: EditorState, part: StoryPart, at: TableAt, rows: number, cols: number): Transaction | null {
	if (!canSplit(at)) return null

	const { map, rect } = at
	const height = rect.bottom - rect.top

	if (rows < 1 || cols < 1 || (rows === 1 && cols === 1) || (rows > height ? height !== 1 : height % rows !== 0)) return null

	const old = paddedGrid(map, map.width)
	const total = widthOf(old, rect.left, rect.right - rect.left)
	const tolerance = Math.max(20, Math.round(total / 100))
	// The sub-cells' boundaries (offsets from the cell's left edge) and the grid lines the cell has.
	const wanted = Array.from({ length: cols - 1 }, (_, k) => Math.round(((k + 1) * total) / cols))
	const lines: number[] = []

	for (let c = rect.left, x = 0; c < rect.right - 1; c++) {
		x += old[c] ?? 0
		lines.push(x)
	}

	// Per old grid column, where new grid lines cut it (offsets from its left edge).
	const cuts = new Map<number, number[]>()

	for (const target of wanted) {
		if (lines.some(line => Math.abs(line - target) <= tolerance)) continue

		for (let c = rect.left, x = 0; c < rect.right; c++) {
			const end = x + (old[c] ?? 0)

			if (target > x && target < end) {
				cuts.set(c, [...(cuts.get(c) ?? []), target - x])

				break
			}

			x = end
		}
	}

	let added = 0

	for (const list of cuts.values()) added += list.length

	if (map.width + added > MAX_TABLE_COLUMNS) return null

	const grid: number[] = []
	// Old grid column -> its first new grid column (one past the last for old.length).
	const shift: number[] = []

	for (let c = 0; c < old.length; c++) {
		shift.push(grid.length)

		let used = 0

		for (const cut of (cuts.get(c) ?? []).sort((a, b) => a - b)) {
			grid.push(cut - used)
			used = cut
		}

		grid.push((old[c] ?? 0) - used)
	}

	shift.push(grid.length)

	// The sub-cells' first grid columns.
	const left = shift[rect.left] ?? 0
	const right = shift[rect.right] ?? grid.length
	const bounds = [left]

	for (let c = left, x = 0; c < right - 1; c++) {
		x += grid[c] ?? 0

		const k = bounds.length

		if (k < cols && Math.abs(x - (wanted[k - 1] ?? Number.NaN)) <= tolerance) bounds.push(c + 1)
	}

	bounds.push(right)

	if (bounds.length !== cols + 1) return null

	const respan = (start: number, span: number): number => (shift[start + span] ?? grid.length) - (shift[start] ?? 0)
	// The merge of the sub-cells in row k of the cell.
	const vMergeAt = (k: number): string | null => {
		if (rows >= height) return null

		const size = height / rows

		return size === 1 ? null : k % size === 0 ? "restart" : "continue"
	}
	const subCells = (cell: Node, k: number): Node[] => {
		const vMerge = vMergeAt(k)
		const out = [cellWith(cell, { colspan: (bounds[1] ?? 0) - left, colwidth: widthOf(grid, left, (bounds[1] ?? 0) - left), vMerge })]

		for (let c = 1; c < cols; c++) {
			const from = bounds[c] ?? 0
			const to = bounds[c + 1] ?? 0

			out.push(newCell(part, cell, to - from, vMerge, widthOf(grid, from, to - from), undefined))
		}

		return out
	}
	const out: Node[] = []

	for (let r = 0; r < map.height; r++) {
		const row = map.row(r)
		const node = row.node
		const cells: Node[] = []

		for (let index = 0; index < row.starts.length; index++) {
			const start = row.starts[index] ?? 0
			const span = row.spans[index] ?? 1
			const cell = map.cell(r, index)

			if (r >= rect.top && r < rect.bottom && start === rect.left) cells.push(...subCells(cell, r - rect.top))
			else cells.push(added === 0 ? cell : cellWith(cell, { colspan: respan(start, span) }))
		}

		let updated =
			cells.every((cell, index) => cell === node.child(index)) && cells.length === node.childCount
				? node
				: node.copy(Fragment.fromArray(cells))

		if (added > 0) {
			const before = shift[row.before] ?? row.before
			const after = respan(row.end, row.after)

			if (before !== row.before || after !== row.after)
				updated = rowWith(updated, { ...rowOverrides(updated), gridBefore: before, gridAfter: after })
		}

		out.push(updated)

		// One row into several: new rows below it, the other cells of the row merged down into them.
		if (height === 1 && rows > 1 && r === rect.top) {
			const merged = updated.copy(
				Fragment.fromArray(
					cells.map((cell, index) =>
						index >= row.starts.findIndex(start => start === rect.left) &&
						index < row.starts.findIndex(start => start === rect.left) + cols
							? cell
							: cellWith(cell, { vMerge: (cell.attrs["vMerge"] as string | null) ?? "restart" })
					)
				)
			)

			out[out.length - 1] = merged

			const first = row.starts.findIndex(start => start === rect.left)

			for (let k = 1; k < rows; k++) {
				out.push(
					newRow(
						part,
						merged,
						cells.map((cell, index) =>
							newCell(
								part,
								cell,
								cell.attrs["colspan"] as number,
								index >= first && index < first + cols ? null : "continue",
								cell.attrs["colwidth"] as number | null,
								undefined
							)
						),
						undefined
					)
				)
			}
		}
	}

	const tr = state.tr
	const cellPos = replaceTable(tr, at, out, added === 0 ? {} : { grid })

	return finish(caretIn(tr, cellPos(rect.top, map.cellAt(rect.top, rect.left)) + 2))
}

// ---- properties -----------------------------------------------------------------------------------------

// Repeats rows [0, bottom of the selection) at the top of each page, or stops the selected header rows and
// those below them from repeating. The header rows end where no vertical merge runs on.
export function toggleHeaderRows(state: EditorState, at: TableAt): Transaction | null {
	const { map, pos } = at
	const on = !map.row(at.cells.top).header
	const whole = map.normalize({
		top: on ? at.cells.bottom - 1 : at.cells.top,
		bottom: on ? at.cells.bottom : at.cells.top + 1,
		left: 0,
		right: map.width
	})
	const from = on ? 0 : whole.top
	const to = on ? whole.bottom : map.height
	const tr = state.tr

	for (let r = to - 1; r >= from; r--) {
		const row = map.row(r)

		if (row.header === on) continue

		setOverrides(tr, map.rowPos(pos, r), row.node, { ...rowOverrides(row.node), header: on })
	}

	return tr.docChanged ? finish(tr) : null
}

export function setTableStyle(state: EditorState, context: DocxEditContext, at: TableAt, styleId: string): Transaction | null {
	const style = context.styles.byId(styleId)

	if (style === undefined || style.type !== "table" || style.id === at.table.attrs["styleId"]) return null

	return finish(state.tr.setNodeAttribute(at.pos, "styleId", style.id))
}

export function alignTable(state: EditorState, at: TableAt, align: DocxTableAlign): Transaction | null {
	const current = tableOverrides(at.table)

	if (current.align === align) return null

	return finish(state.tr.setNodeAttribute(at.pos, "props", { ...current, align }))
}

// The table with grid `grid` (and `attrs`): the cells whose columns changed width take their new width.
function regrid(state: EditorState, at: TableAt, grid: readonly number[], attrs: object = {}): Transaction | null {
	const { map } = at

	if (grid.length === map.grid.length && grid.every((value, k) => value === map.grid[k])) return null

	const old = paddedGrid(map, grid.length)
	const rows: Node[] = []

	for (let r = 0; r < map.height; r++) {
		const row: MapRow = map.row(r)
		const cells: Node[] = []

		for (let index = 0; index < row.starts.length; index++) {
			const start = row.starts[index] ?? 0
			const span = row.spans[index] ?? 1
			const cell = map.cell(r, index)
			const width = widthOf(grid, start, span)

			cells.push(width === widthOf(old, start, span) && cell.attrs["colwidth"] === null ? cell : cellWith(cell, { colwidth: width }))
		}

		rows.push(cells.every((cell, index) => cell === row.node.child(index)) ? row.node : row.node.copy(Fragment.fromArray(cells)))
	}

	const tr = state.tr
	const cellPos = replaceTable(tr, at, rows, { ...attrs, grid })

	return finish(caretIn(tr, cellPos(at.anchor.row, at.anchor.index) + 2))
}

// The table's preferred width (w:tblW) as an edit left it or the source has it.
function tableWidth(part: StoryPart, store: PropsStore, table: Node): { readonly value: number; readonly type: string } | undefined {
	const own = tableOverrides(table).width

	if (own !== undefined) return own

	const src = table.attrs["src"] as number

	if (src === NO_SOURCE || part.kinds[src] !== Item.Table) return undefined

	return store.table(part.document, part.tables.tblPr[part.aux[src] ?? -1] ?? -1).width
}

// Grid column `index` to `twips`, the column right of it giving or taking the difference (a border drag);
// the last column changes the table's width, and with it a width in percent (laid out, the grid is scaled to
// it, which would move every border; Word sets the percentage the drag gives).
export function setColumnWidth(
	state: EditorState,
	part: StoryPart,
	store: PropsStore,
	at: TableAt,
	index: number,
	twips: number
): Transaction | null {
	const grid = paddedGrid(at.map, at.map.width)
	const before = grid.reduce((sum, value) => sum + value, 0)

	if (index >= grid.length) return null

	let width = Math.max(MIN_COLUMN, twips)

	if (index + 1 < grid.length) {
		const pair = (grid[index] ?? 0) + (grid[index + 1] ?? 0)

		width = Math.min(width, pair - MIN_COLUMN)

		if (width < MIN_COLUMN) return null

		grid[index + 1] = pair - width
	}

	grid[index] = width

	const preferred = index + 1 < grid.length ? undefined : tableWidth(part, store, at.table)

	if (preferred?.type !== "pct" || !(preferred.value > 0) || before <= 0) return regrid(state, at, grid)

	const after = grid.reduce((sum, value) => sum + value, 0)
	const value = Math.max(1, Math.round((preferred.value * after) / before))

	return regrid(state, at, grid, { props: { ...tableOverrides(at.table), width: { value, type: "pct" } } })
}

// The selected grid columns (all of them when the selection is in one cell) made equally wide.
export function distributeColumns(state: EditorState, at: TableAt): Transaction | null {
	const grid = paddedGrid(at.map, at.map.width)
	const whole = at.map.areasIn(at.rect) <= 1
	const left = whole ? 0 : at.rect.left
	const right = whole ? grid.length : at.rect.right
	const count = right - left

	if (count < 2) return null

	const total = widthOf(grid, left, count)
	const each = Math.floor(total / count)

	for (let c = left; c < right; c++) grid[c] = c === right - 1 ? total - each * (count - 1) : each

	return regrid(state, at, grid)
}

// Sets `patch` on every cell the selection touches.
function patchCells(state: EditorState, at: TableAt, patch: CellOverrides): Transaction | null {
	const { map, pos } = at
	const { top, bottom, left, right } = at.cells
	const tr = state.tr

	for (let r = bottom - 1; r >= top; r--) {
		const row = map.row(r)

		for (let index = row.starts.length - 1; index >= 0; index--) {
			const start = row.starts[index] ?? 0

			if (start >= right || start + (row.spans[index] ?? 1) <= left) continue

			const cell = map.cell(r, index)
			const current = cellOverrides(cell)
			const next = { ...current, ...patch }

			if (!sameValue(next, current)) setOverrides(tr, map.cellPos(pos, r, index), cell, next)
		}
	}

	return tr.docChanged ? finish(tr) : null
}

export function shadeCells(state: EditorState, at: TableAt, fill: string | null): Transaction | null {
	return patchCells(state, at, { shading: fill === null ? null : fill.toUpperCase() })
}

export function alignCells(state: EditorState, at: TableAt, align: DocxCellVAlign): Transaction | null {
	return patchCells(state, at, { vAlign: align })
}

// ---- for accepting and rejecting tracked changes --------------------------------------------------------

// Takes back the tracked insertion or deletion (`flag`) of cell `index` of row `r` of the table at `pos`: the
// bit and its stamp go, the cell stays (accept a cell insertion, reject a cell deletion).
export function clearCellRevision(tr: Transaction, pos: number, r: number, index: number, flag: number): void {
	const table = tr.doc.nodeAt(pos)

	if (table === null || table.type !== nodes.table || r >= table.childCount || index >= table.child(r).childCount) return

	let cellPos = pos + 1

	for (let k = 0; k < r; k++) cellPos += table.child(k).nodeSize

	const row = table.child(r)

	cellPos++

	for (let k = 0; k < index; k++) cellPos += row.child(k).nodeSize

	const cell = row.child(index)

	tr.setNodeAttribute(cellPos, "revision", (cell.attrs["revision"] as number) & ~flag)
	setOverrides(tr, cellPos, cell, withoutKeys(cellOverrides(cell), [flag === RevisionFlag.Inserted ? "ins" : "del"]))
}

// Takes cells (row, index) out of the table at `pos` (accept a cell deletion, reject a cell insertion), as
// steps from the end: the row's later cells move left, a row losing every cell goes. False, with nothing done,
// when that would leave no row, or move a cell a vertical merge runs through (in the row, or in the rows
// next to it from the first cell taken out on).
export function removeCells(
	tr: Transaction,
	map: TableMap,
	pos: number,
	cells: readonly { readonly row: number; readonly index: number }[]
): boolean {
	const byRow = new Map<number, number[]>()

	for (const { row, index } of cells) {
		if (row < 0 || row >= map.height || index < 0 || index >= map.row(row).starts.length) return false

		byRow.set(
			row,
			[...new Set([...(byRow.get(row) ?? []), index])].sort((a, b) => a - b)
		)
	}

	let emptied = 0

	for (const [r, indices] of byRow) {
		const row = map.row(r)
		const from = row.starts[indices[0] ?? 0] ?? 0
		const merged = (k: number): boolean => {
			if (k < 0 || k >= map.height) return false

			const other = map.row(k)

			return other.starts.some((start, index) => start + (other.spans[index] ?? 1) > from && map.vMerge(k, index) !== null)
		}

		if (merged(r) || merged(r - 1) || merged(r + 1)) return false
		if (indices.length === row.starts.length) emptied++
	}

	if (emptied === map.height) return false

	for (const r of [...byRow.keys()].sort((a, b) => b - a)) {
		const indices = byRow.get(r) ?? []

		if (indices.length === map.row(r).starts.length) {
			tr.delete(map.rowPos(pos, r), rowEnd(map, pos, r))

			continue
		}

		for (const index of [...indices].reverse()) tr.delete(map.cellPos(pos, r, index), cellEnd(map, pos, r, index))
	}

	return true
}

// Takes back the tracked insertion or deletion (`flag`) of row `r` of the table at `pos`: the bit and its
// stamp go, the row stays.
export function clearRowRevision(tr: Transaction, pos: number, r: number, flag: number): void {
	const table = tr.doc.nodeAt(pos)

	if (table === null || table.type !== nodes.table || r >= table.childCount) return

	let rowPos = pos + 1

	for (let k = 0; k < r; k++) rowPos += table.child(k).nodeSize

	const row = table.child(r)

	tr.setNodeAttribute(rowPos, "revision", (row.attrs["revision"] as number) & ~flag)
	setOverrides(tr, rowPos, row, withoutKeys(rowOverrides(row), [flag === RevisionFlag.Inserted ? "ins" : "del"]))
}
