import { SHEET_MAX_COLS, SHEET_MAX_ROWS } from "../../sheet/protocol"
import { FormulaTable } from "../../xlsx/model/formulas"
import { SheetStore, StoreBuilder } from "../../xlsx/model/store"
import { SheetStrings } from "../../xlsx/model/strings"
import { encodeXstring } from "../../xlsx/edit/document"
import { cellKeyOf } from "../../xlsx/edit/history"
import type { SharedStrings } from "../../xlsx/model/sharedStrings"
import { MAX_SHEET_LINKS } from "../../xlsx/hyperlinks/edit"
import {
	CellKind,
	ColFlag,
	FormulaKind,
	NO_SPAN,
	RowFlag,
	type XlsxColRange,
	type XlsxRichRun,
	type XlsxRowRun
} from "../../xlsx/model/types"
import type { OdsCellStyles } from "../styles/cell"
import type { OdsIssues } from "../issues"
import { OdsReadError } from "./errors"
import { StyleRuns, StyleRunsBuilder } from "./runs"

// One table:table read into the xlsx sheet model's parts, row run by row run: content cells are materialised (a
// repeated one once per copy, up to the caps), formats of empty cells go to the run index unless short, and rows
// holding only properties wait for the table's end, where the commonest height becomes the default and costs nothing.

// Copies of one content cell a column repeat may make, and content rows a row repeat may make: past them the rest of
// the run is named (issue "repeat") and not shown.
const MAX_CELL_COPIES = SHEET_MAX_COLS
const MAX_ROW_CELLS = 4_000_000
// A style-only run of at most this many cells in an unrepeated row is cheaper as empty store cells than as a run.
const SHORT_RUN = 16
const MAX_MERGES = 1_000_000

export interface OdsCellInput {
	col: number
	repeat: number
	covered: boolean
	// Effective cell style name (own, else the row's default, else undefined for the column's).
	style: string | undefined
	kind: CellKind
	number: number
	text: string
	// The xlsx-grammar formula, "" for none.
	formula: string
	colSpan: number
	rowSpan: number
	matrixCols: number
	matrixRows: number
	validation: string | undefined
	link: string | undefined
	note: Uint8Array | undefined
	// Text LibreOffice reads as an edit cell: shown as written (a shared string with runs), with its spans' text
	// styles from where each starts.
	rich: boolean
	runs: readonly { readonly start: number; readonly style: string | undefined }[] | undefined
	// The engine calculates the result again once built (XlsxSheetModel.recalculate), the cached one shown till then.
	recalculate: boolean
}

// Field by field: Object.assign takes SpiderMonkey's generic path, several times this per cell.
function copyInput(cell: OdsCellInput, input: OdsCellInput): void {
	cell.col = input.col
	cell.repeat = input.repeat
	cell.covered = input.covered
	cell.style = input.style
	cell.kind = input.kind
	cell.number = input.number
	cell.text = input.text
	cell.formula = input.formula
	cell.colSpan = input.colSpan
	cell.rowSpan = input.rowSpan
	cell.matrixCols = input.matrixCols
	cell.matrixRows = input.matrixRows
	cell.validation = input.validation
	cell.link = input.link
	cell.note = input.note
	cell.rich = input.rich
	cell.runs = input.runs
	cell.recalculate = input.recalculate
}

export interface OdsRowInput {
	repeat: number
	style: string | undefined
	defaultCellStyle: string | undefined
	hidden: boolean
	outline: number
	collapsed: boolean
}

export interface OdsColumnRun {
	readonly first: number
	// Exclusive.
	readonly end: number
	readonly style: string | undefined
	readonly defaultXf: number
	readonly hidden: boolean
	readonly outline: number
	readonly collapsed: boolean
}

// A note or link of one cell, or of a block of its copies (a repeated cell): rows and columns inclusive. The copies
// share the one annotation as read, never one per copy.
export interface OdsCellNote {
	readonly row: number
	readonly col: number
	endRow: number
	readonly endCol: number
	readonly markup: Uint8Array
}

export interface OdsCellLink {
	readonly row: number
	readonly col: number
	endRow: number
	readonly endCol: number
	readonly href: string
}

export interface OdsBuiltSheet {
	readonly store: SheetStore
	readonly cols: XlsxColRange[]
	readonly merges: Uint32Array
	readonly styleRuns: StyleRuns | undefined
	// Rows of properties alone (a height, hidden), as runs: the store holds none of them.
	readonly rowRuns: XlsxRowRun[]
	readonly defaultRowHeight: number
	// Most rows of the default height have it set by their style (customHeight, ECMA-376 Part 1 §18.3.1.81).
	readonly defaultRowSet: boolean
	readonly defaultColPx: number
	// Cells naming each content validation, as rectangles [startRow, startCol, endRow, endCol].
	readonly validations: ReadonlyMap<string, readonly (readonly [number, number, number, number])[]>
	readonly notes: readonly OdsCellNote[]
	readonly links: readonly OdsCellLink[]
	// table:table-header-rows and -columns: [first, last] or null.
	readonly headerRows: readonly [number, number] | null
	readonly headerCols: readonly [number, number] | null
	readonly rowBreaks: readonly number[]
	readonly colBreaks: readonly number[]
	readonly recalculate: ReadonlySet<number>
	// The rows the table's row elements give (their repeats), and the columns its column elements declare.
	readonly rowCount: number
	readonly columnCount: number
}

// Text as the store keeps it (ST_Xstring escapes), at no cost for text that needs none: control characters XML
// cannot hold, U+FFFE, U+FFFF, surrogates (encodeXstring keeps a valid pair) and a literal "_xHHHH_".
function xstring(text: string): string {
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)

		if ((code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) || (code >= 0xd800 && code <= 0xdfff) || code >= 0xfffe) {
			return encodeXstring(text)
		}

		if (code === 0x5f && text.charCodeAt(i + 1) === 0x78) return encodeXstring(text)
	}

	return text
}

interface PropertyRun {
	readonly row: number
	repeat: number
	readonly heightPt: number
	readonly flags: number
	// The rows' style gave no height (the default's taken).
	readonly implicit: boolean
}

// LibreOffice's default row height (0.178 in) and column width (2.258 cm), for a file whose styles give none.
export const DEFAULT_ROW_PT = 12.8
const DEFAULT_COL_PX = 85

export class OdsSheetBuilder {
	readonly name: string
	readonly strings = new SheetStrings()
	readonly formulas = new FormulaTable()
	private readonly styles: OdsCellStyles
	private readonly issues: OdsIssues
	private readonly builder = new StoreBuilder()
	private readonly runs = new StyleRunsBuilder()
	private readonly columns: OdsColumnRun[] = []
	private nextColumn = 0
	private row = 0
	private readonly merges: number[] = []
	private readonly validations = new Map<string, [number, number, number, number][]>()
	private readonly notes: OdsCellNote[] = []
	private readonly links: OdsCellLink[] = []
	// The last note and link block starting at each column, which the same cell's copy in the next row extends.
	private readonly noteAt = new Map<number, OdsCellNote>()
	private readonly linkAt = new Map<number, OdsCellLink>()
	private readonly recalculate = new Set<number>()
	private readonly propertyRuns: PropertyRun[] = []
	// Rows by height in points, for the default.
	private readonly heights = new Map<number, number>()
	// Rows of each height whose style sets it (not LibreOffice's optimal height).
	private readonly setHeights = new Map<number, number>()
	private readonly rowBreaks: number[] = []
	private readonly colBreaks: number[] = []
	headerRows: [number, number] | null = null
	headerCols: [number, number] | null = null
	// The row under way: its cells in column order, materialised (content, or a short style-only run); entries are
	// reused from row to row.
	private readonly pool: OdsCellInput[] = []
	private cellsInRow = 0
	private rowCol = 0
	cellCount = 0

	private readonly sharedStrings: SharedStrings
	private readonly budget: { left: number }

	constructor(name: string, styles: OdsCellStyles, issues: OdsIssues, sharedStrings: SharedStrings, budget: { left: number }) {
		this.budget = budget
		this.name = name
		this.styles = styles
		this.issues = issues
		this.sharedStrings = sharedStrings
	}

	// An edit cell's text as a shared string with runs (each the span's font, none for the cell's).
	private richString(cell: OdsCellInput): number {
		const text = xstring(cell.text)
		const runs: XlsxRichRun[] = []
		const given = cell.runs ?? [{ start: 0, style: undefined }]
		const first = given[0]
		// Text before the first mark (empty paragraphs' line ends) is the cell's own: the first run starts at 0.
		const marks =
			first === undefined || first.start <= 0
				? given
				: first.style === undefined
					? [{ start: 0, style: undefined }, ...given.slice(1)]
					: [{ start: 0, style: undefined }, ...given]

		for (let k = 0; k < marks.length; k++) {
			const mark = marks[k]
			const start = mark?.start ?? 0
			const end = marks[k + 1]?.start ?? cell.text.length

			if (end <= start) continue

			runs.push({
				text: xstring(cell.text.slice(start, end)),
				font: mark?.style === undefined ? null : this.styles.textFont(mark.style)
			})
		}

		return this.sharedStrings.append(text, runs.length > 0 ? runs : [{ text, font: null }])
	}

	get currentRow(): number {
		return this.row
	}

	get currentCol(): number {
		return this.rowCol
	}

	// --- Columns ---

	column(
		repeat: number,
		style: string | undefined,
		defaultCellStyle: string | undefined,
		hidden: boolean,
		outline: number,
		collapsed: boolean,
		header: boolean
	): void {
		const first = this.nextColumn

		if (first >= SHEET_MAX_COLS) {
			this.issues.add("repeat", this.name, "columns past the last")

			return
		}

		const end = Math.min(SHEET_MAX_COLS, first + repeat)

		this.columns.push({
			first,
			end,
			style,
			defaultXf: defaultCellStyle === undefined ? 0 : this.styles.xf(defaultCellStyle),
			hidden,
			outline,
			collapsed
		})
		this.nextColumn = end

		if (header) this.headerCols = this.headerCols === null ? [first, end - 1] : [this.headerCols[0], end - 1]
		if (this.styles.column(style).pageBreak && first > 0) this.colBreaks.push(first)
	}

	// The column run holding `col` (its index), -1 past the declared columns.
	private columnAt(col: number): number {
		let low = 0
		let high = this.columns.length - 1

		while (low <= high) {
			const mid = (low + high) >>> 1

			if ((this.columns[mid]?.first ?? 0) <= col) low = mid + 1
			else high = mid - 1
		}

		return high >= 0 && col < (this.columns[high]?.end ?? 0) ? high : -1
	}

	columnXf(col: number): number {
		return this.columns[this.columnAt(col)]?.defaultXf ?? 0
	}

	// --- Rows and cells ---

	beginRow(): void {
		this.cellsInRow = 0
		this.rowCol = 0
	}

	// A cell (or a run of equal ones) of the row under way, in column order.
	cell(input: OdsCellInput, hasContent: boolean, rowInput: OdsRowInput): void {
		const col = this.rowCol
		const repeat = Math.min(input.repeat, SHEET_MAX_COLS - col)

		this.rowCol = col + input.repeat

		if (repeat <= 0) {
			if (hasContent) this.issues.add("repeat", this.name, "cells past the last column")

			return
		}

		if (input.validation !== undefined)
			this.validated(
				input.validation,
				this.row,
				col,
				this.row + Math.min(rowInput.repeat, SHEET_MAX_ROWS - this.row) - 1,
				col + repeat - 1
			)

		if (hasContent) {
			const copies = Math.min(repeat, MAX_CELL_COPIES)

			if (copies < repeat) this.issues.add("repeat", this.name, "cell copies", repeat - copies)

			this.take(input, col, copies)

			return
		}

		// Style alone: nothing where it equals the column's default, a short run as empty cells, a long one as a run.
		const xf = input.style === undefined ? -1 : this.styles.xf(input.style)
		let at = col
		const stop = col + repeat

		while (at < stop) {
			const index = this.columnAt(at)
			const columnEnd = index < 0 ? SHEET_MAX_COLS : (this.columns[index]?.end ?? SHEET_MAX_COLS)
			const end = Math.min(stop, columnEnd)
			const columnXf = this.columns[index]?.defaultXf ?? 0
			const own = xf < 0 ? columnXf : xf

			if (own !== columnXf) {
				if (rowInput.repeat === 1 && end - at <= SHORT_RUN) {
					const cell = this.take(input, at, end - at)

					cell.kind = CellKind.EMPTY
					cell.number = 0
					cell.formula = ""
				} else {
					this.runs.add(at, end, own)
				}
			}

			at = end
		}
	}

	// The row's next cell entry, a copy of `input` at `col` for `repeat` columns.
	private take(input: OdsCellInput, col: number, repeat: number): OdsCellInput {
		let cell = this.pool[this.cellsInRow]

		if (cell === undefined) {
			cell = { ...input }
			this.pool.push(cell)
		} else {
			copyInput(cell, input)
		}

		cell.col = col
		cell.repeat = repeat
		this.cellsInRow++

		return cell
	}

	private validated(name: string, startRow: number, startCol: number, endRow: number, endCol: number): void {
		let list = this.validations.get(name)

		if (list === undefined) {
			list = []
			this.validations.set(name, list)
		}

		// Extends the rectangle above when it spans the same columns.
		for (let k = list.length - 1; k >= Math.max(0, list.length - 8); k--) {
			const rect = list[k]

			if (rect !== undefined && rect[1] === startCol && rect[3] === endCol && rect[2] + 1 === startRow) {
				rect[2] = endRow

				return
			}
		}

		const last = list[list.length - 1]

		if (last !== undefined && last[0] === startRow && last[2] === endRow && last[3] + 1 === startCol) {
			last[3] = endCol

			return
		}

		list.push([startRow, startCol, endRow, endCol])
	}

	endRow(input: OdsRowInput): void {
		const first = this.row
		const repeat = Math.min(input.repeat, SHEET_MAX_ROWS - first)

		this.row = first + input.repeat

		if (repeat <= 0) {
			if (this.cellsInRow > 0) this.issues.add("repeat", this.name, "rows past the last")

			return
		}

		const style = this.styles.row(input.style)

		if (style.pageBreak && first > 0) this.rowBreaks.push(first)

		// A row whose style gives no height has LibreOffice's default (ODF 1.3 Part 3 §17.17.2 leaves it to the
		// application), whatever the commonest height of the others.
		const height = Number.isNaN(style.heightPt) ? DEFAULT_ROW_PT : style.heightPt
		const flags =
			(input.hidden ? RowFlag.HIDDEN : 0) |
			(style.custom ? RowFlag.CUSTOM_HEIGHT : 0) |
			(input.collapsed ? RowFlag.COLLAPSED : 0) |
			(Math.min(7, input.outline) << RowFlag.OUTLINE_SHIFT)

		if (!Number.isNaN(height)) this.heights.set(height, (this.heights.get(height) ?? 0) + repeat)
		if (!Number.isNaN(height) && style.custom) this.setHeights.set(height, (this.setHeights.get(height) ?? 0) + repeat)

		if (this.runs.pending) this.runs.endRows(first, first + repeat)

		if (this.cellsInRow === 0) {
			if (flags !== 0 || !Number.isNaN(height)) this.propertyRun(first, repeat, height, flags, Number.isNaN(style.heightPt))

			return
		}

		let width = 0

		for (let k = 0; k < this.cellsInRow; k++) width += this.pool[k]?.repeat ?? 0

		const rows = Math.max(1, Math.min(repeat, Math.floor(MAX_ROW_CELLS / Math.max(1, width))))

		if (rows < repeat) {
			this.issues.add("repeat", this.name, "row copies", repeat - rows)
			// The rest keeps the row's properties.
			this.propertyRun(first + rows, repeat - rows, height, flags, Number.isNaN(style.heightPt))
		}

		for (let r = 0; r < rows; r++) {
			const row = first + r

			this.builder.beginRow()

			for (let k = 0; k < this.cellsInRow; k++) this.place(this.pool[k] as OdsCellInput, row)

			this.builder.endRow(row, NO_SPAN, NO_SPAN, flags, height, -1)
		}
	}

	// Rows of properties alone, joined to the run before when alike and next to it. Each run is charged to the
	// workbook's cell limit as one: it is one entry, however many rows it covers.
	private propertyRun(row: number, repeat: number, heightPt: number, flags: number, implicit: boolean): void {
		const last = this.propertyRuns[this.propertyRuns.length - 1]

		if (
			last !== undefined &&
			last.row + last.repeat === row &&
			last.flags === flags &&
			last.implicit === implicit &&
			(last.heightPt === heightPt || (Number.isNaN(last.heightPt) && Number.isNaN(heightPt)))
		) {
			last.repeat += repeat

			return
		}

		if (--this.budget.left < 0) throw new OdsReadError("tooLarge", "The workbook holds more cells than the limit")

		this.propertyRuns.push({ row, repeat, heightPt, flags, implicit })
	}

	private place(cell: OdsCellInput, row: number): void {
		const xf = cell.style === undefined ? this.columnXf(cell.col) : this.styles.xf(cell.style)
		let value = cell.number
		let kind: CellKind = cell.kind

		if (cell.kind === CellKind.INLINE_STRING && cell.rich) {
			kind = CellKind.SHARED_STRING
			value = this.richString(cell)
		} else if (cell.kind === CellKind.INLINE_STRING || cell.kind === CellKind.ERROR) {
			value = this.strings.intern(cell.kind === CellKind.INLINE_STRING ? xstring(cell.text) : cell.text)
		}

		for (let k = 0; k < cell.repeat; k++) {
			const col = cell.col + k
			let formula = -1

			if (cell.formula !== "") {
				const matrix = cell.matrixCols > 0 || cell.matrixRows > 0

				formula = this.formulas.add({
					kind: matrix ? FormulaKind.ARRAY : FormulaKind.NORMAL,
					text: cell.formula,
					sharedIndex: -1,
					ref: matrix
						? {
								startRow: row,
								startCol: col,
								endRow: row + Math.max(1, cell.matrixRows) - 1,
								endCol: col + Math.max(1, cell.matrixCols) - 1
							}
						: null,
					flags: 0,
					cellMetadata: -1
				})
			}

			if (--this.budget.left < 0) throw new OdsReadError("tooLarge", "The workbook holds more cells than the limit")

			this.builder.cell(col, kind, value, xf, formula)
			this.cellCount++

			if ((cell.colSpan > 1 || cell.rowSpan > 1) && !cell.covered) {
				if (this.merges.length < MAX_MERGES * 4) {
					this.merges.push(
						row,
						col,
						Math.min(SHEET_MAX_ROWS - 1, row + cell.rowSpan - 1),
						Math.min(SHEET_MAX_COLS - 1, col + cell.colSpan - 1)
					)
				} else {
					this.issues.add("repeat", this.name, "merges")
				}
			}

			if (cell.recalculate && formula >= 0) this.recalculate.add(cellKeyOf(row, col))
		}

		const endCol = cell.col + cell.repeat - 1

		if (cell.note !== undefined) {
			const last = this.noteAt.get(cell.col)

			if (last !== undefined && last.markup === cell.note && last.endCol === endCol && last.endRow + 1 === row) {
				last.endRow = row
			} else {
				const note: OdsCellNote = { row, col: cell.col, endRow: row, endCol, markup: cell.note }

				this.notes.push(note)
				this.noteAt.set(cell.col, note)
			}
		}

		if (cell.link !== undefined) {
			const last = this.linkAt.get(cell.col)

			if (last !== undefined && last.href === cell.link && last.endCol === endCol && last.endRow + 1 === row) {
				last.endRow = row
			} else {
				if (this.links.length >= MAX_SHEET_LINKS) throw new OdsReadError("tooLarge", "A sheet holds more links than the limit")

				const link: OdsCellLink = { row, col: cell.col, endRow: row, endCol, href: cell.link }

				this.links.push(link)
				this.linkAt.set(cell.col, link)
			}
		}
	}

	headerRowsStart(): void {
		this.headerRowsFirst = this.row
	}

	headerRowsEnd(): void {
		if (this.row > this.headerRowsFirst) this.headerRows = [this.headerRowsFirst, this.row - 1]
	}

	private headerRowsFirst = 0

	// --- The table's end ---

	finish(): OdsBuiltSheet {
		let defaultHeight = Number.NaN
		let most = -1

		for (const [height, count] of this.heights) {
			if (count > most) {
				most = count
				defaultHeight = height
			}
		}

		if (Number.isNaN(defaultHeight)) defaultHeight = DEFAULT_ROW_PT

		const rowRuns: XlsxRowRun[] = []

		for (const run of this.propertyRuns) {
			// A user-set height (customHeight) changes nothing shown where it equals the default.
			if ((run.flags & ~RowFlag.CUSTOM_HEIGHT) === 0 && (Number.isNaN(run.heightPt) || run.heightPt === defaultHeight)) continue

			// No height of their own, the default's: none.
			rowRuns.push({
				first: run.row,
				last: run.row + run.repeat - 1,
				flags: run.flags,
				heightPt: run.implicit && run.heightPt === defaultHeight ? Number.NaN : run.heightPt
			})
		}

		const store = new SheetStore(this.builder.finish(), this.strings, this.formulas)
		const widths = new Map<number, number>()

		for (const column of this.columns) {
			const width = this.styles.column(column.style).widthPx

			if (!Number.isNaN(width)) widths.set(width, (widths.get(width) ?? 0) + column.end - column.first)
		}

		let defaultColPx = DEFAULT_COL_PX
		let widest = -1

		for (const [width, count] of widths) {
			if (count > widest) {
				widest = count
				defaultColPx = width
			}
		}

		const cols: XlsxColRange[] = []

		for (const column of this.columns) {
			const width = this.styles.column(column.style).widthPx
			const flags =
				(column.hidden ? ColFlag.HIDDEN : 0) |
				(column.collapsed ? ColFlag.COLLAPSED : 0) |
				(Math.min(7, column.outline) << ColFlag.OUTLINE_SHIFT) |
				(Number.isNaN(width) || width === defaultColPx ? 0 : ColFlag.CUSTOM_WIDTH)

			if (flags === 0 && column.defaultXf === 0 && (Number.isNaN(width) || width === defaultColPx)) continue

			const px = Number.isNaN(width) ? defaultColPx : width

			cols.push({
				first: column.first,
				last: column.end - 1,
				width: Number.NaN,
				px,
				style: column.defaultXf === 0 ? -1 : column.defaultXf,
				flags
			})
		}

		return {
			store,
			cols,
			merges: Uint32Array.from(this.merges),
			styleRuns: this.runs.finish(),
			rowRuns,
			defaultRowHeight: defaultHeight,
			defaultRowSet: (this.setHeights.get(defaultHeight) ?? 0) * 2 > (this.heights.get(defaultHeight) ?? 0),
			defaultColPx,
			validations: this.validations,
			notes: this.notes,
			links: this.links,
			headerRows: this.headerRows,
			headerCols: this.headerCols,
			rowBreaks: this.rowBreaks,
			colBreaks: this.colBreaks,
			recalculate: this.recalculate,
			rowCount: this.row,
			columnCount: this.nextColumn
		}
	}
}
