import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetId, type SheetRange } from "../../sheet/protocol"
import { defaultColWidth } from "../../sheet/sizes/sizes"
import type { SheetChangeSink } from "../../sheet/registry"
import type { RowProperties, SheetStore } from "../model/store"
import { RowFlag, type XlsxColRange, type XlsxSheetModel } from "../model/types"
import type { XlsxDocument } from "./document"
import { NO_ROW, rowState, type RowState } from "./history"

// Formatting whole columns or whole rows as Excel writes it: the columns' <col style> (the rows' s with
// customFormat) take the format, so every cell without one of its own shows it, past the used area too; cells
// that have one are patched; and wherever an empty cell would show something other than the format it should
// (a styled row's cells under formatted columns, as a row's style beats a column's, and a styled column's cells in
// formatted rows; the range's first or last row or column where a border differs, the grid's last row or column
// included; the cells before a range that starts inside the grid) it gets a cell of its own. A row without a style
// of its own takes one where that writes fewer cells: select-all puts its first and last rows' lines on those
// rows' styles, not on a cell per column.

// A cell's new format from its old one and where it is (row, col). A patch depends on the place only as the
// range's first or last row or column, so one inside place stands for all. `cell`: the answer is written to that
// cell (else it is an axis's format, or compared with one).
export type PositionedPatch = (xf: number, row: number, col: number, cell: boolean) => number

export type WholeAxis = "cols" | "rows" | "all"

const LAST_ROW = SHEET_MAX_ROWS - 1
const LAST_COL = SHEET_MAX_COLS - 1

// How a range reaching the grid's last row (column) is formatted: through its columns (rows), select-all through
// both. A range starting below the first row (right of the first column) is formatted this way only when it is
// too large to format cell by cell (formatPlan).
export function wholeAxis(range: SheetRange): WholeAxis | null {
	const toLastRow = range.endRow >= LAST_ROW
	const toLastCol = range.endCol >= LAST_COL

	if (toLastRow && toLastCol && range.startRow === 0 && range.startCol === 0) return "all"
	if (toLastRow) return "cols"

	return toLastCol ? "rows" : null
}

// Whether a whole-axis range starts inside the grid (A2:A1048576), so the cells before it keep their look.
export function partialAxis(range: SheetRange, kind: WholeAxis): boolean {
	return kind === "cols" ? range.startRow > 0 : kind === "rows" ? range.startCol > 0 : false
}

export interface CellTarget {
	readonly row: number
	readonly col: number
	readonly style: number
}

// What a whole-axis format changes besides cells, accumulated over an edit's ranges: the column list and row
// properties as they were before the edit's first change and as the edit left them.
export class AxisChange {
	cols: { readonly before: readonly XlsxColRange[]; after: readonly XlsxColRange[] } | null = null
	readonly rows = new Map<number, { readonly before: RowState; after: RowState }>()

	readonly sheet: SheetId

	constructor(sheet: SheetId) {
		this.sheet = sheet
	}

	get changed(): boolean {
		return this.cols !== null || this.rows.size > 0
	}

	get bytes(): number {
		return ((this.cols?.before.length ?? 0) + (this.cols?.after.length ?? 0)) * 64 + this.rows.size * 96
	}

	// Columns now as `after`, kept for undo.
	setCols(model: XlsxSheetModel, after: readonly XlsxColRange[]): void {
		this.cols = { before: this.cols?.before ?? model.cols, after }
		model.cols = after
	}

	setRows(store: SheetStore, rows: readonly { readonly row: number; readonly state: RowState }[]): void {
		for (const { row, state } of rows) this.rows.set(row, { before: this.rows.get(row)?.before ?? rowState(store, row), after: state })

		store.setRows(rows.map(({ row, state }) => ({ row, height: state.height, style: state.style, flags: state.flags })))
	}
}

// What a whole-axis range takes: cells (patched or new) and the axes' new formats.
export interface WholeAxisTargets {
	readonly cells: CellTarget[]
	readonly cols: readonly XlsxColRange[] | null
	readonly rows: readonly { readonly row: number; readonly state: RowState }[]
}

// The sheet's default column width, as a <col> without customWidth stores it.
function defaultWidth(model: XlsxSheetModel): number {
	return defaultColWidth(model.format)
}

function sameRange(a: XlsxColRange, b: XlsxColRange): boolean {
	return (
		a.last + 1 === b.first &&
		a.style === b.style &&
		a.flags === b.flags &&
		(Number.isNaN(a.width) ? Number.isNaN(b.width) : a.width === b.width)
	)
}

// Columns of one place class: the same column range (or none), all inside or all outside the formatted columns,
// and all or none of them its first (last) column. `rep` stands for them all.
interface ColClass {
	readonly first: number
	readonly last: number
	readonly rep: number
	readonly range: XlsxColRange | undefined
	// The column style before the edit (0: none) and after it.
	readonly old: number
	after: number
}

// Every column of the grid in classes, split at the formatted columns' first and last.
function colClasses(cols: readonly XlsxColRange[], first: number, last: number): ColClass[] {
	const sorted = cols.every((range, i) => i === 0 || (cols[i - 1]?.first ?? 0) <= range.first)
		? cols
		: [...cols].sort((a, b) => a.first - b.first)
	const out: ColClass[] = []
	const push = (from: number, to: number, range: XlsxColRange | undefined): void => {
		const old = range === undefined || range.style < 0 ? 0 : range.style

		for (const [a, b] of [
			[from, Math.min(to, first - 1)],
			[Math.max(from, first), Math.min(to, first)],
			[Math.max(from, first + 1), Math.min(to, last - 1)],
			[Math.max(from, last, first + 1), Math.min(to, last)],
			[Math.max(from, last + 1), to]
		] as const) {
			if (a <= b) out.push({ first: a, last: b, rep: a, range, old, after: old })
		}
	}
	let at = 0

	for (const range of sorted) {
		if (range.last < at) continue
		if (range.first > at) push(at, Math.min(range.first - 1, LAST_COL), undefined)

		const end = Math.min(range.last, LAST_COL)

		push(Math.max(at, range.first), end, range)
		at = end + 1

		if (at > LAST_COL) break
	}

	if (at <= LAST_COL) push(at, LAST_COL, undefined)

	return out
}

// The column list with each class's new style; ranges are cut and joined as the styles fall.
function restyledCols(model: XlsxSheetModel, classes: readonly ColClass[], first: number, last: number): XlsxColRange[] {
	const out: XlsxColRange[] = []
	const width = defaultWidth(model)

	for (const range of model.cols) {
		if (range.last < first || range.first > last) out.push(range)
		else if (range.first < first) out.push({ ...range, last: first - 1 })

		if (range.first <= last && range.last > last) out.push({ ...range, first: last + 1 })
	}

	const inside: XlsxColRange[] = []

	for (const entry of classes) {
		if (entry.last < first || entry.first > last) continue

		const own = entry.range
		let range: XlsxColRange | null

		if (entry.after === entry.old) range = own === undefined ? null : { ...own, first: entry.first, last: entry.last }
		else if (own === undefined) range = { first: entry.first, last: entry.last, width, style: entry.after, flags: 0 }
		else range = { ...own, first: entry.first, last: entry.last, style: entry.after === 0 ? -1 : entry.after }

		// A range left with nothing of its own (no style, width or flag) goes.
		if (range !== null && range.style < 0 && range.flags === 0 && (Number.isNaN(range.width) || range.width === width)) range = null

		if (range === null) continue

		const previous = inside[inside.length - 1]

		if (previous !== undefined && sameRange(previous, range)) inside[inside.length - 1] = { ...previous, last: range.last }
		else inside.push(range)
	}

	return [...out, ...inside].sort((a, b) => a.first - b.first)
}

function styledRows(store: SheetStore): Map<number, number> {
	const rows = new Map<number, number>()

	for (const block of store.blocks) {
		for (let i = 0; i < block.rows.length; i++) {
			if (((block.rowFlags[i] ?? 0) & RowFlag.CUSTOM_FORMAT) === 0) continue

			const style = block.rowStyles[i] ?? -1

			if (style >= 0) rows.set(block.firstRow + (block.rows[i] ?? 0), style)
		}
	}

	return rows
}

// Work done between the host's turns (cells looked at, classes weighed).
const STEP = 512
// Rows of existing cells looked at per scan.
const BAND = 64

// The cells and axis properties a whole-axis range takes, giving the host turns. Null when it would take more cells
// than `budget`. Throws StyleLimitError (from `patch`).
export function* wholeAxisTargets(
	model: XlsxSheetModel,
	range: SheetRange,
	kind: WholeAxis,
	patch: PositionedPatch,
	budget: number
): Generator<void, WholeAxisTargets | null> {
	const store = model.store as SheetStore
	const targets: CellTarget[] = []
	const taken = new Set<number>()
	let work = 0
	const due = (units: number): boolean => {
		work += units

		if (work < STEP) return false

		work = 0

		return true
	}
	const add = (row: number, col: number, style: number): boolean => {
		targets.push({ row, col, style })
		taken.add(row * SHEET_MAX_COLS + col)

		return targets.length <= budget
	}
	const firstCol = range.startCol
	const lastCol = Math.min(range.endCol, LAST_COL)
	const firstRow = range.startRow
	const lastRow = Math.min(range.endRow, LAST_ROW)
	// A row and a column inside the range, standing for every one not on its edge.
	const insideRow = Math.min(firstRow + 1, lastRow)
	const insideCol = Math.min(firstCol + 1, lastCol)

	// Cells of their own: patched where they are.
	const endRow = Math.min(lastRow, store.rowExtent - 1)
	const endCol = Math.min(lastCol, store.colExtent - 1)

	for (let band = firstRow; band <= endRow; band += BAND) {
		let fits = true
		let seen = 0

		store.forEachCell({ startRow: band, startCol: firstCol, endRow: Math.min(endRow, band + BAND - 1), endCol }, (row, col, slot) => {
			if (!fits) return

			const xf = store.style(slot)
			const next = patch(xf, row, col, true)

			seen++

			if (next !== xf) fits = add(row, col, next)
		})

		if (!fits) return null
		if (due(seen + 1)) yield
	}

	const classes = colClasses(model.cols, firstCol, lastCol)

	// The columns' new styles.
	if (kind !== "rows") {
		for (const entry of classes) {
			if (entry.first >= firstCol && entry.last <= lastCol) entry.after = patch(entry.old, insideRow, entry.rep, false)
			if (due(1)) yield
		}
	}

	const styled = styledRows(store)
	const rowStates: { row: number; state: RowState }[] = []
	// The column style most columns have, which a row without a style of its own starts from.
	let widest = 0
	let dominant = 0

	for (const entry of classes) {
		if (entry.last - entry.first + 1 > widest) {
			widest = entry.last - entry.first + 1
			dominant = entry.old
		}
	}

	// Cells row `row` needs so each of its empty cells shows what it should while the row shows `rowStyle` (null: no
	// style of its own, so its cells show their columns'): an empty cell shows its old look (`own`, else its
	// column's), patched inside the range. `count`: only how many columns, nothing added. False past the budget.
	const inRows = (row: number): boolean => row >= firstRow && row <= lastRow
	const rowCells = (row: number, own: number | null, rowStyle: number | null, count: boolean): number | false => {
		let cells = 0

		for (const entry of classes) {
			if (kind === "cols" && (entry.last < firstCol || entry.first > lastCol)) continue

			const old = own ?? entry.old
			const desired = inRows(row) && entry.rep >= firstCol && entry.rep <= lastCol ? patch(old, row, entry.rep, true) : old
			const shown = rowStyle ?? entry.after

			if (desired === shown) continue

			cells += entry.last - entry.first + 1

			if (count) continue

			for (let col = entry.first; col <= entry.last; col++) {
				if (store.locate(row, col) >= 0 || taken.has(row * SHEET_MAX_COLS + col)) continue
				if (!add(row, col, desired)) return false
			}
		}

		return cells
	}

	// Rows to look at: every row of whole rows; with columns, the styled rows, the edge rows and the rows above.
	const rows: number[] = []

	if (kind === "rows") {
		if (lastRow - firstRow + 1 > budget) return null

		for (let row = firstRow; row <= lastRow; row++) rows.push(row)
	} else {
		const changed = classes.some(entry => entry.after !== entry.old)

		for (const row of styled.keys()) if (kind === "all" || row >= firstRow) rows.push(row)
		for (const row of [firstRow, lastRow]) if (!styled.has(row)) rows.push(row)

		// Above a range starting below row 1: the rows without a style of their own keep their columns' old look.
		if (kind === "cols" && changed && firstRow > 0) {
			let plain = firstRow

			for (const row of styled.keys()) if (row < firstRow) plain--

			const per = classes.reduce(
				(sum, entry) =>
					entry.after !== entry.old && entry.first >= firstCol && entry.last <= lastCol
						? sum + entry.last - entry.first + 1
						: sum,
				0
			)

			if (plain * per > budget - targets.length) return null

			for (let row = 0; row < firstRow; row++) if (!styled.has(row)) rows.push(row)
		}

		rows.sort((a, b) => a - b)
	}

	let previous = -1

	for (const row of rows) {
		if (row === previous) continue

		previous = row

		const state = rowState(store, row)
		const own = styled.get(row) ?? null
		let rowStyle = own

		// Whole rows restyle every row, select-all its styled and edge rows; columns leave rows as they are.
		if (kind !== "cols") {
			if (own !== null) {
				const next = patch(own, row, insideCol, false)

				rowStyle = next === 0 ? null : next
			} else {
				// A row without a style of its own: a style of its own, or none (its cells then keep their columns'),
				// whichever leaves fewer cells to write.
				const next = patch(dominant, row, insideCol, false)
				const styledCount = next === 0 ? false : rowCells(row, null, next, true)
				const plainCount = rowCells(row, null, null, true)

				rowStyle = styledCount !== false && plainCount !== false && styledCount < plainCount ? next : null
			}
		}

		if (rowCells(row, own, rowStyle, false) === false) return null

		if (rowStyle !== own) {
			const flags = state === NO_ROW ? 0 : state.flags

			// The default format is no row format at all.
			rowStates.push({
				row,
				state:
					rowStyle === null
						? { height: state.height, style: -1, flags: flags & ~RowFlag.CUSTOM_FORMAT }
						: { height: state.height, style: rowStyle, flags: flags | RowFlag.CUSTOM_FORMAT }
			})
		}

		if (due(classes.length)) yield
	}

	const cols = kind === "rows" ? null : restyledCols(model, classes, firstCol, lastCol)

	return { cells: targets, cols, rows: rowStates }
}

// Puts the column list and rows as `which` side of the change has them (an undo or redo).
export function restoreAxisChange(document: XlsxDocument, change: AxisChange, which: "before" | "after", sink: SheetChangeSink): void {
	const model = document.model(change.sheet)

	if (model === null || !change.changed) return

	if (change.cols !== null) {
		model.cols = which === "before" ? change.cols.before : change.cols.after
		document.workbook.changes.sheetSection(change.sheet, "cols")
	}

	if (change.rows.size > 0) {
		const store = model.store as SheetStore

		const rows: RowProperties[] = []

		for (const [row, states] of change.rows) {
			const state = which === "before" ? states.before : states.after

			rows.push({ row, height: state.height, style: state.style, flags: state.flags })
		}

		store.setRows(rows)

		document.workbook.changes.sheetRows(change.sheet)
	}

	sink.layout(change.sheet)
}
