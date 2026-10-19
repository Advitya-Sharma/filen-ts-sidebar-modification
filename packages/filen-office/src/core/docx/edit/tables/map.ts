import type { Mark, Node } from "prosemirror-model"
import { Item, Revision, type StoryPart } from "../../model/index"
import { Ns } from "../../model/namespaces"
import type { PropsStore, TableCellProps, TableProps, TableRowProps } from "../../props/index"
import type { DocxThemeColor } from "../commands/formatting"
import type { DocxCellVAlign, DocxTableAlign } from "../commands/tables"
import type { ElementData } from "../copies"
import { NO_SOURCE, docxSchema } from "../schema"

// The grid of a table as the editor holds it, and what table commands record on tables, rows and cells.
// Slots come from each row's w:gridBefore, its cells' spans (colspan) and w:gridAfter; a vertically merged
// cell is a chain of cells in one grid column, the first w:vMerge "restart" and the rest "continue". Rows are
// read lazily, so a snapshot looks only at the rows the selection touches.

const nodes = docxSchema.nodes
const marks = docxSchema.marks

// Who changed a row or cell while changes were tracked, and when (minute precision, as an ISO date).
export interface TrackStamp {
	readonly author: string
	readonly date: string
	// The tracked change it records (a "revision" id of the editing context; every row or cell one command
	// changed shares it). The file's w:id is the writer's own.
	readonly id?: number
}

// A source's own tracked property changes (w:tblPrChange, w:tblGridChange, w:trPrChange, w:tblPrExChange,
// w:tcPrChange, w:cellMerge) once accepted (the records go) or rejected (the properties before them are back).
export type SourceResolution = "accepted" | "rejected"

// What table commands set on a table (its `props`): absent keys are as the source has them.
export interface TableOverrides {
	// A source table (item of the story's part) whose properties a new table starts from (a copy of it).
	readonly from?: number
	// The w:tblPr of another part's table a copy of it starts from.
	readonly copy?: ElementData
	readonly align?: DocxTableAlign
	// w:tblW: twips for "dxa", fiftieths of a percent for "pct".
	readonly width?: { readonly value: number; readonly type: "dxa" | "pct" | "auto" }
	// A table the editor made: single borders (when no grid style applies) and Word's look of a new table.
	readonly borders?: boolean
	readonly look?: boolean
	// Changes recorded while tracking: of the table's properties (w:tblPrChange) and of its grid
	// (w:tblGridChange).
	readonly change?: TableChange
	readonly gridChange?: GridChange
	readonly source?: SourceResolution
}

// A property change recorded while tracking: its stamp (the id is the change's, shared by everything one
// command changed) and what the properties were before it. One per property element, the first: a later
// tracked edit of the same element adds to it.
export interface TableChange extends TrackStamp {
	readonly styleId: string | null
	readonly props: TableOverrides
}

export interface GridChange extends TrackStamp {
	readonly grid: readonly number[]
}

export interface RowChange extends TrackStamp {
	readonly props: RowOverrides
}

export interface CellChange extends TrackStamp {
	readonly props: CellOverrides
	readonly colwidth: number | null
	// The cell was merged or split (untracked, as Word does) after the record: the width before was another
	// span's; rejected, the cell takes the width of the columns it spans now.
	readonly reshaped?: boolean
}

// What table commands record on a row (its row_props mark value).
export interface RowOverrides {
	// A source row (item of the story's part) whose properties a new row starts from.
	readonly from?: number
	// The w:trPr of another part's row a copy of it starts from.
	readonly copy?: ElementData
	readonly header?: boolean
	readonly gridBefore?: number
	readonly gridAfter?: number
	// The stamps of the row's tracked insertion or deletion (its revision bits say which applies).
	readonly ins?: TrackStamp
	readonly del?: TrackStamp
	readonly change?: RowChange
	readonly source?: SourceResolution
}

// A theme colour a cell's fill follows: tint or shade 0–255 (none: the colour itself).
export interface ShadingTheme {
	readonly theme: DocxThemeColor
	readonly tint?: number
	readonly shade?: number
}

// What table commands record on a cell (its cell_props mark value).
export interface CellOverrides {
	readonly from?: number
	// The w:tcPr of another part's cell a copy of it starts from.
	readonly copy?: ElementData
	// "RRGGBB", or null for no fill.
	readonly shading?: string | null
	// The theme colour the fill is (w:themeFill with w:themeFillTint or w:themeFillShade), set with `shading`.
	readonly shadingTheme?: ShadingTheme
	readonly vAlign?: DocxCellVAlign
	readonly ins?: TrackStamp
	readonly del?: TrackStamp
	readonly change?: CellChange
	readonly source?: SourceResolution
}

// Overrides keys that are no property: the model, tracked insertions and deletions, change records.
export const RECORD_KEYS: readonly string[] = ["from", "copy", "ins", "del", "change", "gridChange", "source"]

const NONE: Readonly<Record<string, never>> = Object.freeze({})

export function tableOverrides(table: Node): TableOverrides {
	return (table.attrs["props"] as TableOverrides | null) ?? NONE
}

function markValue<T>(node: Node, type: Mark["type"]): T | undefined {
	return type.isInSet(node.marks)?.attrs["value"] as T | undefined
}

export function rowOverrides(row: Node): RowOverrides {
	return markValue<RowOverrides>(row, marks.row_props) ?? NONE
}

export function cellOverrides(cell: Node): CellOverrides {
	return markValue<CellOverrides>(cell, marks.cell_props) ?? NONE
}

// The source item a table, row or cell takes its properties from: its own, or the one it was modelled on.
export function propertySource(part: StoryPart, node: Node, kind: number): number {
	const own = node.attrs["src"] as number
	const from = (
		node.type === nodes.table ? tableOverrides(node) : node.type === nodes.table_row ? rowOverrides(node) : cellOverrides(node)
	).from

	if (own !== NO_SOURCE && part.kinds[own] === kind) return own
	if (from !== undefined && part.kinds[from] === kind) return from

	return NO_SOURCE
}

// ---- the source's own tracked property changes ---------------------------------------------------------

// Revision kinds a table, row or cell holds in its property elements.
const PROPERTY_KINDS: ReadonlySet<number> = new Set([
	Revision.TableProperties,
	Revision.TableGrid,
	Revision.RowProperties,
	Revision.TablePropertiesException,
	Revision.CellProperties,
	Revision.CellMerge
])
const NO_CHANGES: ReadonlyMap<number, number> = new Map()
const changesByPart = new WeakMap<StoryPart, Map<number, Map<number, number>>>()

// The tracked property changes of source item `item`: their change elements (w:tblPrChange...) by Revision
// kind. Read once per part.
export function sourceChanges(part: StoryPart, item: number): ReadonlyMap<number, number> {
	let byItem = changesByPart.get(part)

	if (byItem === undefined) {
		byItem = new Map()

		const revisions = part.revisions

		for (let k = 0; k < revisions.count; k++) {
			const kind = revisions.kind[k] ?? 0

			if (!PROPERTY_KINDS.has(kind)) continue

			const of = revisions.item[k] ?? -1
			let kinds = byItem.get(of)

			if (kinds === undefined) {
				kinds = new Map()
				byItem.set(of, kinds)
			}

			if (!kinds.has(kind)) kinds.set(kind, revisions.node[k] ?? -1)
		}

		changesByPart.set(part, byItem)
	}

	return item === NO_SOURCE ? NO_CHANGES : (byItem.get(item) ?? NO_CHANGES)
}

// The source item of a table, row or cell of `kind` (an Item) that holds its tracked property changes: its
// own (a copy of one holds none), or -1.
export function ownItem(part: StoryPart, node: Node, kind: number): number {
	const src = node.attrs["src"] as number

	return src !== NO_SOURCE && part.kinds[src] === kind ? src : NO_SOURCE
}

// Whether `node` (a table, row or cell) holds the source's tracked change of `kind` (a Revision), unresolved.
export function unresolvedChange(part: StoryPart, node: Node, kind: number): boolean {
	const item = ownItem(part, node, node.type === nodes.table ? Item.Table : node.type === nodes.table_row ? Item.Row : Item.Cell)

	if (item === NO_SOURCE || !sourceChanges(part, item).has(kind)) return false

	const overrides =
		node.type === nodes.table ? tableOverrides(node) : node.type === nodes.table_row ? rowOverrides(node) : cellOverrides(node)

	return overrides.source === undefined
}

// The source properties of a table, row or cell (`item` of the part) as the editor holds them: its own, or
// with its tracked change rejected, those before it.
export function sourceTableProps(part: StoryPart, store: PropsStore, item: number, rejected: boolean): TableProps {
	const props = store.table(part.document, part.tables.tblPr[part.aux[item] ?? -1] ?? -1)

	return rejected ? (props.change?.previous ?? props) : props
}

export function sourceRowProps(part: StoryPart, store: PropsStore, item: number, rejected: boolean): TableRowProps {
	const props = store.row(part.document, part.rows.trPr[part.aux[item] ?? -1] ?? -1)

	return rejected ? (props.change?.previous ?? props) : props
}

export function sourceCellProps(part: StoryPart, store: PropsStore, item: number, rejected: boolean): TableCellProps {
	const props = store.cell(part.document, part.cells.tcPr[part.aux[item] ?? -1] ?? -1)

	return rejected ? (props.change?.previous ?? props) : props
}

// The grid a source w:tblGridChange holds (the grid before it), or undefined.
export function previousGrid(part: StoryPart, change: number | undefined): number[] | undefined {
	if (change === undefined || change < 0) return undefined

	const reader = part.reader
	const document = part.document

	for (let child = document.firstChildElement(change); child >= 0; child = document.nextSiblingElement(child)) {
		if (!reader.is(child, Ns.W, "tblGrid")) continue

		const out: number[] = []

		for (let col = document.firstChildElement(child); col >= 0; col = document.nextSiblingElement(col))
			if (reader.is(col, Ns.W, "gridCol")) out.push(reader.int(col, Ns.W, "w", 0))

		return out
	}

	return undefined
}

// ---- grids ----------------------------------------------------------------------------------------------

export interface RowGrid {
	readonly before: number
	readonly after: number
	readonly header: boolean
}

const PLAIN_ROW: RowGrid = Object.freeze({ before: 0, after: 0, header: false })
const rowGrids = new WeakMap<Node, RowGrid>()
// Source rows' own, by item: a row's node changes with every keystroke in it, its source does not.
const sourceGrids = new WeakMap<StoryPart, Map<number, RowGrid>>()

function gridOf(before: number, after: number, header: boolean): RowGrid {
	return before === 0 && after === 0 && !header ? PLAIN_ROW : { before, after, header }
}

function sourceGrid(part: StoryPart, store: PropsStore, item: number): RowGrid {
	let grids = sourceGrids.get(part)

	if (grids === undefined) {
		grids = new Map()
		sourceGrids.set(part, grids)
	}

	let grid = grids.get(item)

	if (grid === undefined) {
		const props = store.row(part.document, part.rows.trPr[part.aux[item] ?? -1] ?? -1)

		grid = gridOf(props.gridBefore ?? 0, props.gridAfter ?? 0, props.header ?? false)
		grids.set(item, grid)
	}

	return grid
}

// A row's grid columns before and after its cells, and whether it repeats as a header row.
export function rowGrid(part: StoryPart, store: PropsStore, row: Node): RowGrid {
	let grid = rowGrids.get(row)

	if (grid !== undefined) return grid

	const source = propertySource(part, row, Item.Row)
	const overrides = rowOverrides(row)
	let base = PLAIN_ROW

	if (source !== NO_SOURCE && overrides.source === "rejected") {
		const props = sourceRowProps(part, store, source, true)

		base = gridOf(props.gridBefore ?? 0, props.gridAfter ?? 0, props.header ?? false)
	} else if (source !== NO_SOURCE) {
		base = sourceGrid(part, store, source)
	}

	if (marks.row_props.isInSet(row.marks) === undefined) return base

	grid = gridOf(overrides.gridBefore ?? base.before, overrides.gridAfter ?? base.after, overrides.header ?? base.header)
	rowGrids.set(row, grid)

	return grid
}

export interface MapRow {
	readonly node: Node
	readonly before: number
	readonly after: number
	readonly header: boolean
	// Per cell: its first grid column, span and offset from the row's content start.
	readonly starts: readonly number[]
	readonly spans: readonly number[]
	readonly offsets: readonly number[]
	// The column after the last cell.
	readonly end: number
}

// A rectangle of grid slots: rows [top, bottom), grid columns [left, right).
export interface TableRect {
	readonly top: number
	readonly bottom: number
	readonly left: number
	readonly right: number
}

export function sameRect(a: TableRect, b: TableRect): boolean {
	return a.top === b.top && a.bottom === b.bottom && a.left === b.left && a.right === b.right
}

// A row's record in TableMap.pool: before, after, end, header (0/1), the cell count n, then n starts, n spans and
// n offsets.
const BEFORE = 0
const AFTER = 1
const END = 2
const HEADER = 3
const COUNT = 4
const CELLS = 5
// Rows kept as MapRow objects (row() gives them to callers; the table's rows stay in the pool).
const VIEWS = 32

export class TableMap {
	readonly table: Node
	private readonly part: StoryPart
	private readonly store: PropsStore
	// Rows read so far (a snapshot reads the selection's only), packed: per row the start of its record in
	// `pool` (-1 before it is read). A long table's rows as objects and arrays would take several times this.
	private records: Int32Array | undefined
	private pool = new Int32Array(64)
	private used = 0
	private readonly views = new Map<number, MapRow>()
	private rowOffsets: Int32Array | undefined

	constructor(table: Node, part: StoryPart, store: PropsStore) {
		this.table = table
		this.part = part
		this.store = store
	}

	get height(): number {
		return this.table.childCount
	}

	get grid(): readonly number[] {
		return this.table.attrs["grid"] as readonly number[]
	}

	// Grid columns: the tblGrid's, or as many as the widest row needs (reads every row).
	get width(): number {
		let width = this.grid.length

		for (let r = 0; r < this.height; r++) {
			const at = this.record(r)

			width = Math.max(width, (this.pool[at + END] ?? 0) + (this.pool[at + AFTER] ?? 0))
		}

		return width
	}

	private offsets(): Int32Array {
		if (this.rowOffsets === undefined) {
			const out = new Int32Array(this.height)

			this.table.forEach((_row, offset, index) => {
				out[index] = offset
			})
			this.rowOffsets = out
		}

		return this.rowOffsets
	}

	// Row `r`'s record in the pool, read now when it was not.
	private record(r: number): number {
		const records = (this.records ??= new Int32Array(this.height).fill(-1))
		const found = records[r] ?? -1

		if (found >= 0) return found

		const node = this.table.child(r)
		const grid = rowGrid(this.part, this.store, node)
		const count = node.childCount
		const at = this.reserve(CELLS + count * 3)
		const pool = this.pool
		let column = grid.before

		pool[at + BEFORE] = grid.before
		pool[at + AFTER] = grid.after
		pool[at + HEADER] = grid.header ? 1 : 0
		pool[at + COUNT] = count
		node.forEach((cell, offset, index) => {
			const span = Math.max(1, cell.attrs["colspan"] as number)

			pool[at + CELLS + index] = column
			pool[at + CELLS + count + index] = span
			pool[at + CELLS + 2 * count + index] = offset
			column += span
		})
		pool[at + END] = column
		records[r] = at

		return at
	}

	private reserve(size: number): number {
		const at = this.used

		if (at + size > this.pool.length) {
			const grown = new Int32Array(Math.max(at + size, this.pool.length * 2))

			grown.set(this.pool.subarray(0, at))
			this.pool = grown
		}

		this.used = at + size

		return at
	}

	// Cell counts and cells of row `r` without a MapRow (what the map's own walks read).
	private count(r: number): number {
		return this.pool[this.record(r) + COUNT] ?? 0
	}

	private start(r: number, index: number): number {
		return this.pool[this.record(r) + CELLS + index] ?? 0
	}

	private span(r: number, index: number): number {
		const at = this.record(r)

		return this.pool[at + CELLS + (this.pool[at + COUNT] ?? 0) + index] ?? 1
	}

	row(r: number): MapRow {
		let row = this.views.get(r)

		if (row !== undefined) return row

		const at = this.record(r)
		const pool = this.pool
		const count = pool[at + COUNT] ?? 0
		const cells = at + CELLS

		row = {
			node: this.table.child(r),
			before: pool[at + BEFORE] ?? 0,
			after: pool[at + AFTER] ?? 0,
			header: pool[at + HEADER] === 1,
			starts: Array.from(pool.subarray(cells, cells + count)),
			spans: Array.from(pool.subarray(cells + count, cells + 2 * count)),
			offsets: Array.from(pool.subarray(cells + 2 * count, cells + 3 * count)),
			end: pool[at + END] ?? 0
		}

		if (this.views.size >= VIEWS) {
			const oldest = this.views.keys().next()

			if (oldest.done !== true) this.views.delete(oldest.value)
		}

		this.views.set(r, row)

		return row
	}

	cell(r: number, index: number): Node {
		return this.table.child(r).child(index)
	}

	// The position of cell `index` of row `r`, the table being at `tablePos` (reads every row once).
	cellPos(tablePos: number, r: number, index: number): number {
		const at = this.record(r)

		return this.rowPos(tablePos, r) + 1 + (this.pool[at + CELLS + 2 * (this.pool[at + COUNT] ?? 0) + index] ?? 0)
	}

	rowPos(tablePos: number, r: number): number {
		return tablePos + 1 + (this.offsets()[r] ?? 0)
	}

	// The cell of row `r` covering grid column `column`, or -1.
	cellAt(r: number, column: number): number {
		const at = this.record(r)
		const pool = this.pool
		const count = pool[at + COUNT] ?? 0
		let low = 0
		let high = count - 1

		while (low <= high) {
			const middle = (low + high) >> 1
			const start = pool[at + CELLS + middle] ?? 0

			if (column < start) high = middle - 1
			else if (column >= start + (pool[at + CELLS + count + middle] ?? 1)) low = middle + 1
			else return middle
		}

		return -1
	}

	// The cell of row `r` starting at grid column `column` with span `span`, or -1.
	cellExactly(r: number, column: number, span: number): number {
		const index = this.cellAt(r, column)

		return index >= 0 && this.start(r, index) === column && this.span(r, index) === span ? index : -1
	}

	vMerge(r: number, index: number): "restart" | "continue" | null {
		return this.cell(r, index).attrs["vMerge"] as "restart" | "continue" | null
	}

	// Whether cell `index` of row `r` continues a merge from the row above (a cell there in the same columns
	// takes part in one).
	continues(r: number, index: number): boolean {
		if (r === 0 || this.vMerge(r, index) !== "continue") return false

		const above = this.cellExactly(r - 1, this.start(r, index), this.span(r, index))

		return above >= 0 && this.vMerge(r - 1, above) !== null
	}

	// The rows [top, bottom) of the vertical merge holding cell `index` of row `r` (just its row when none).
	chain(r: number, index: number): { top: number; bottom: number } {
		const column = this.start(r, index)
		const span = this.span(r, index)
		let top = r
		let at = index

		while (this.continues(top, at)) {
			at = this.cellExactly(top - 1, column, span)
			top--
		}

		let bottom = r + 1

		if (this.vMerge(r, index) !== null) {
			for (; bottom < this.height; bottom++) {
				const below = this.cellExactly(bottom, column, span)

				if (below < 0 || !this.continues(bottom, below)) break
			}
		}

		return { top, bottom }
	}

	// The smallest rectangle holding `rect` that cuts no cell and no vertical merge.
	normalize(rect: TableRect): TableRect {
		let { top, bottom, left, right } = rect

		for (;;) {
			const before = { top, bottom, left, right }

			for (let r = top; r < bottom; r++) {
				const count = this.count(r)

				for (let index = 0; index < count; index++) {
					const start = this.start(r, index)
					const end = start + this.span(r, index)

					if (end <= left || start >= right) continue

					left = Math.min(left, start)
					right = Math.max(right, end)

					if (this.vMerge(r, index) !== null) {
						const chain = this.chain(r, index)

						top = Math.min(top, chain.top)
						bottom = Math.max(bottom, chain.bottom)
					}
				}
			}

			if (sameRect(before, { top, bottom, left, right })) return before
		}
	}

	// The cells starting in `rect` (row, index), row by row.
	cellsIn(rect: TableRect): { row: number; index: number }[] {
		const out: { row: number; index: number }[] = []

		for (let r = rect.top; r < rect.bottom; r++) {
			const count = this.count(r)

			for (let index = 0; index < count; index++) {
				const start = this.start(r, index)

				if (start >= rect.left && start + this.span(r, index) <= rect.right) out.push({ row: r, index })
			}
		}

		return out
	}

	// The distinct merged areas in `rect` (a vertical merge counts once).
	areasIn(rect: TableRect): number {
		let count = 0

		for (const { row, index } of this.cellsIn(rect)) if (!this.continues(row, index)) count++

		return count
	}
}

const maps = new WeakMap<Node, TableMap>()

// The map of `table` (one per table node; its rows are read as they are asked for).
export function tableMap(table: Node, part: StoryPart, store: PropsStore): TableMap {
	let map = maps.get(table)

	if (map === undefined) {
		map = new TableMap(table, part, store)
		maps.set(table, map)
	}

	return map
}
