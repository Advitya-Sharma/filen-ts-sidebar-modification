import type { Mark, Node } from "prosemirror-model"
import { Item, type StoryPart } from "../../model/index"
import type { PropsStore } from "../../props/index"
import type { DocxCellVAlign, DocxTableAlign } from "../commands/tables"
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

// What table commands set on a table (its `props`): absent keys are as the source has them.
export interface TableOverrides {
	readonly align?: DocxTableAlign
	// w:tblW: twips for "dxa", fiftieths of a percent for "pct".
	readonly width?: { readonly value: number; readonly type: "dxa" | "pct" | "auto" }
	// A table the editor made: single borders (when no grid style applies) and Word's look of a new table.
	readonly borders?: boolean
	readonly look?: boolean
}

// What table commands record on a row (its row_props mark value).
export interface RowOverrides {
	// A source row (item of the story's part) whose properties a new row starts from.
	readonly from?: number
	readonly header?: boolean
	readonly gridBefore?: number
	readonly gridAfter?: number
	// The stamps of the row's tracked insertion or deletion (its revision bits say which applies).
	readonly ins?: TrackStamp
	readonly del?: TrackStamp
}

// What table commands record on a cell (its cell_props mark value).
export interface CellOverrides {
	readonly from?: number
	// "RRGGBB", or null for no fill.
	readonly shading?: string | null
	readonly vAlign?: DocxCellVAlign
	readonly ins?: TrackStamp
	readonly del?: TrackStamp
}

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

// The source item a row or cell takes its properties from: its own, or the one it was modelled on.
export function propertySource(part: StoryPart, node: Node, kind: number): number {
	const own = node.attrs["src"] as number
	const from = (node.type === nodes.table_row ? rowOverrides(node) : cellOverrides(node)).from

	if (own !== NO_SOURCE && part.kinds[own] === kind) return own
	if (from !== undefined && part.kinds[from] === kind) return from

	return NO_SOURCE
}

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
	const base = source === NO_SOURCE ? PLAIN_ROW : sourceGrid(part, store, source)

	if (marks.row_props.isInSet(row.marks) === undefined) return base

	const overrides = rowOverrides(row)

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

export class TableMap {
	readonly table: Node
	private readonly part: StoryPart
	private readonly store: PropsStore
	// Rows read so far (a snapshot reads the selection's only).
	private readonly rows = new Map<number, MapRow>()
	private rowOffsets: number[] | undefined

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

		for (let r = 0; r < this.height; r++) width = Math.max(width, this.row(r).end + this.row(r).after)

		return width
	}

	private offsets(): number[] {
		if (this.rowOffsets === undefined) {
			const out: number[] = []

			this.table.forEach((_row, offset) => out.push(offset))
			this.rowOffsets = out
		}

		return this.rowOffsets
	}

	row(r: number): MapRow {
		let row = this.rows.get(r)

		if (row !== undefined) return row

		const node = this.table.child(r)
		const grid = rowGrid(this.part, this.store, node)
		const starts: number[] = []
		const spans: number[] = []
		const offsets: number[] = []
		let column = grid.before

		node.forEach((cell, offset) => {
			const span = Math.max(1, cell.attrs["colspan"] as number)

			starts.push(column)
			spans.push(span)
			offsets.push(offset)
			column += span
		})

		row = {
			node,
			before: grid.before,
			after: grid.after,
			header: grid.header,
			starts,
			spans,
			offsets,
			end: column
		}
		this.rows.set(r, row)

		return row
	}

	cell(r: number, index: number): Node {
		return this.table.child(r).child(index)
	}

	// The position of cell `index` of row `r`, the table being at `tablePos` (reads every row once).
	cellPos(tablePos: number, r: number, index: number): number {
		return this.rowPos(tablePos, r) + 1 + (this.row(r).offsets[index] ?? 0)
	}

	rowPos(tablePos: number, r: number): number {
		return tablePos + 1 + (this.offsets()[r] ?? 0)
	}

	// The cell of row `r` covering grid column `column`, or -1.
	cellAt(r: number, column: number): number {
		const row = this.row(r)
		let low = 0
		let high = row.starts.length - 1

		while (low <= high) {
			const middle = (low + high) >> 1
			const start = row.starts[middle] ?? 0

			if (column < start) high = middle - 1
			else if (column >= start + (row.spans[middle] ?? 1)) low = middle + 1
			else return middle
		}

		return -1
	}

	// The cell of row `r` starting at grid column `column` with span `span`, or -1.
	cellExactly(r: number, column: number, span: number): number {
		const index = this.cellAt(r, column)
		const row = this.row(r)

		return index >= 0 && row.starts[index] === column && row.spans[index] === span ? index : -1
	}

	vMerge(r: number, index: number): "restart" | "continue" | null {
		return this.cell(r, index).attrs["vMerge"] as "restart" | "continue" | null
	}

	// Whether cell `index` of row `r` continues a merge from the row above (a cell there in the same columns
	// takes part in one).
	continues(r: number, index: number): boolean {
		if (r === 0 || this.vMerge(r, index) !== "continue") return false

		const row = this.row(r)
		const above = this.cellExactly(r - 1, row.starts[index] ?? 0, row.spans[index] ?? 1)

		return above >= 0 && this.vMerge(r - 1, above) !== null
	}

	// The rows [top, bottom) of the vertical merge holding cell `index` of row `r` (just its row when none).
	chain(r: number, index: number): { top: number; bottom: number } {
		const row = this.row(r)
		const column = row.starts[index] ?? 0
		const span = row.spans[index] ?? 1
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
				const row = this.row(r)

				for (let index = 0; index < row.starts.length; index++) {
					const start = row.starts[index] ?? 0
					const end = start + (row.spans[index] ?? 1)

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
			const row = this.row(r)

			for (let index = 0; index < row.starts.length; index++) {
				const start = row.starts[index] ?? 0

				if (start >= rect.left && start + (row.spans[index] ?? 1) <= rect.right) out.push({ row: r, index })
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
