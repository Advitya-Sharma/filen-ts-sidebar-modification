import { dateToSerial, formatValue, type Cell, type CellStyle, type CellValue, type Sheet, type Workbook } from "hucre"
import {
	cellKey,
	colWidthToPx,
	rowHeightToPx,
	shownFormula,
	type CellView,
	type SheetView,
	type SpreadsheetDoc
} from "@filen/office/core/sheet"
import { StyleTable, styleView } from "@/features/spreadsheet/lib/styleTable"

const DEFAULT_FONT_SIZE = 11

// hucre's key for a cell in a sheet's cell map.
export function cellIdOf(row: number, col: number): string {
	return `${String(row)},${String(col)}`
}

export function parseCellId(cellId: string): [number, number] {
	const comma = cellId.indexOf(",")

	return [Number(cellId.slice(0, comma)), Number(cellId.slice(comma + 1))]
}

function pad(value: number): string {
	return String(value).padStart(2, "0")
}

// A date without a date format, shown the way the file's own locale-free default would be.
function isoDate(date: Date): string {
	return `${String(date.getUTCFullYear())}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`
}

// A date as typed: its time too when it has one ("2024-03-05 14:30"), the time alone for one on the
// workbook's epoch day (a time-only cell, "14:30"), to the second.
function typedDate(date: Date, is1904: boolean): string {
	const shown = new Date(Math.round(date.getTime() / 1000) * 1000)
	const seconds = shown.getUTCSeconds()
	const time = `${pad(shown.getUTCHours())}:${pad(shown.getUTCMinutes())}${seconds === 0 ? "" : `:${pad(seconds)}`}`
	const serial = dateToSerial(shown, is1904)

	if (serial >= 0 && serial < 1) {
		return time
	}

	return shown.getUTCHours() === 0 && shown.getUTCMinutes() === 0 && seconds === 0 ? isoDate(shown) : `${isoDate(shown)} ${time}`
}

// What a value shows through `numFmt`, the formatter's own rendering except where it has none. A number
// under a date format counts its days from the workbook's epoch.
export function displayText(value: CellValue | undefined, numFmt: string | undefined, is1904 = false): string {
	if (value === null || value === undefined) {
		return ""
	}

	if (value instanceof Date && (numFmt === undefined || numFmt === "General")) {
		return isoDate(value)
	}

	if (typeof value === "boolean") {
		return value ? "TRUE" : "FALSE"
	}

	return formatValue(value, numFmt ?? "General", { is1904 })
}

// The text editing a cell starts from: its formula, or the value as typed rather than as formatted.
export function inputText(value: CellValue | undefined, cell: Cell | undefined, is1904 = false): string {
	if (cell?.formula !== undefined) {
		return `=${shownFormula(cell.formula)}`
	}

	if (value === null || value === undefined) {
		return ""
	}

	if (value instanceof Date) {
		return typedDate(value, is1904)
	}

	if (typeof value === "boolean") {
		return value ? "TRUE" : "FALSE"
	}

	return String(value)
}

// Builds views of one workbook's cells, keeping one style table for all of them, so the ids a first view
// hands out stay valid for every later patch of the same workbook.
export class WorkbookViews {
	readonly styles = new StyleTable()
	// The reader hands each cell a style object of its own, sharing the font, fill and alignment objects of
	// its format: those, by identity, and the number format key the drawn part of a style.
	private readonly partIds = new WeakMap<object, number>()
	private nextPartId = 1
	private readonly structuralIds = new Map<string, number | undefined>()
	// Neighbouring cells mostly share a format: the last one looked up is checked first.
	private last: { style: CellStyle; id: number | undefined } | null = null
	private readonly themeColors: readonly string[] | undefined
	// Serial dates count from 1904 (older Mac files).
	private readonly is1904: boolean

	constructor(themeColors: readonly string[] | undefined, is1904 = false) {
		this.themeColors = themeColors
		this.is1904 = is1904
	}

	private partId(part: object | undefined): number {
		if (part === undefined) {
			return 0
		}

		let id = this.partIds.get(part)

		if (id === undefined) {
			id = this.nextPartId++
			this.partIds.set(part, id)
		}

		return id
	}

	private styleId(style: CellStyle | undefined): number | undefined {
		if (style === undefined) {
			return undefined
		}

		const last = this.last

		if (
			last !== null &&
			last.style.font === style.font &&
			last.style.fill === style.fill &&
			last.style.alignment === style.alignment &&
			last.style.numFmt === style.numFmt
		) {
			return last.id
		}

		const key = `${String(this.partId(style.font))},${String(this.partId(style.fill))},${String(this.partId(style.alignment))},${style.numFmt ?? ""}`

		let id: number | undefined

		if (this.structuralIds.has(key)) {
			id = this.structuralIds.get(key)
		} else {
			id = this.styles.add(styleView(style, this.themeColors, DEFAULT_FONT_SIZE))
			this.structuralIds.set(key, id)
		}

		this.last = { style, id }

		return id
	}

	// One cell's view, or null for a cell with nothing to show.
	cell(sheet: Sheet, row: number, col: number): CellView | null {
		return this.view(sheet.rows[row]?.[col], sheet.cells?.get(cellIdOf(row, col)))
	}

	view(value: CellValue | undefined, cell: Cell | undefined): CellView | null {
		if ((value === null || value === undefined) && cell === undefined) {
			return null
		}

		const shown = cell?.formula !== undefined ? (cell.formulaResult ?? value) : value
		const text =
			cell?.richText !== undefined
				? cell.richText.map(run => run.text).join("")
				: displayText(shown, cell?.style?.numFmt, this.is1904)
		const input = inputText(cell?.formula !== undefined ? shown : value, cell, this.is1904)
		const styleId = this.styleId(cell?.style)
		const view: CellView = { text }

		if (input !== text) view.input = input
		if (styleId !== undefined) view.style = styleId
		if (typeof shown === "number" || typeof shown === "boolean" || shown instanceof Date) view.numeric = true
		if (cell?.type === "error" || cell?.formulaResultError === true) {
			view.error = true
		}

		return text === "" && view.style === undefined && view.input === undefined ? null : view
	}

	sheet(sheet: Sheet, lockStructure: boolean): SheetView {
		return sheetView(sheet, this, lockStructure)
	}
}

// A workbook's cells, one view per non-empty or formatted cell.
// Whether a row or column has the Text ("@") format.
function hasTextFormat(sheet: Sheet): boolean {
	if (sheet.columns?.some(column => column.style?.numFmt === "@") === true) return true

	for (const def of sheet.rowDefs?.values() ?? []) {
		if (def.style?.numFmt === "@") return true
	}

	return false
}

function sheetView(sheet: Sheet, views: WorkbookViews, lockStructure: boolean): SheetView {
	const cells = new Map<number, CellView>()
	let colCount = 0
	// Cells with details are read from their map once, rather than looked up by a string key at every
	// position; the positions they cover are skipped when the plain values are read.
	const detailed = new Set<number>()
	// Cells with details that show nothing: their own (absent) format, not their row's or column's, says
	// how what is typed into them reads.
	const unshown: number[] = []
	const textFormatted = hasTextFormat(sheet)

	for (const [cellId, cell] of sheet.cells ?? []) {
		const [row, col] = parseCellId(cellId)
		const values = sheet.rows[row]
		const at = cellKey(row, col)
		const view = values === undefined || col >= values.length ? null : views.view(values[col], cell)

		detailed.add(at)

		if (view !== null) {
			cells.set(at, view)
		} else if (textFormatted) {
			unshown.push(at)
		}
	}

	for (let row = 0; row < sheet.rows.length; row++) {
		const values = sheet.rows[row] ?? []

		colCount = Math.max(colCount, values.length)

		for (let col = 0; col < values.length; col++) {
			const value = values[col]

			if (value === null || value === undefined) {
				continue
			}

			const at = cellKey(row, col)

			if (detailed.size > 0 && detailed.has(at)) {
				continue
			}

			const view = views.view(value, undefined)

			if (view !== null) {
				cells.set(at, view)
			}
		}
	}

	const merges = (sheet.merges ?? []).map(merge => ({
		startRow: merge.startRow,
		startCol: merge.startCol,
		endRow: merge.endRow,
		endCol: merge.endCol
	}))
	const extent = sheetExtent(sheet)

	const colWidths = new Map<number, number>()
	const hiddenCols: number[] = []
	const textCols: number[] = []
	const textRows: number[] = []
	const formattedRows: number[] = []

	sheet.columns?.forEach((column, index) => {
		if (column.style?.numFmt === "@") {
			textCols.push(index)
		}

		if (column.width !== undefined) {
			colWidths.set(index, colWidthToPx(column.width))
		}

		if (column.hidden === true) {
			hiddenCols.push(index)
		}
	})

	const rowHeights = new Map<number, number>()
	const hiddenRows: number[] = []

	sheet.rowDefs?.forEach((def, index) => {
		if (def.style?.numFmt === "@") {
			textRows.push(index)
		} else if (def.style !== undefined) {
			formattedRows.push(index)
		}

		if (def.height !== undefined) {
			rowHeights.set(index, rowHeightToPx(def.height))
		}

		if (def.hidden === true) {
			hiddenRows.push(index)
		}
	})

	return {
		name: sheet.name,
		rowCount: extent.rowCount,
		colCount: Math.max(colCount, extent.colCount),
		cells,
		merges,
		colWidths,
		rowHeights,
		hiddenCols,
		hiddenRows,
		frozenRows: sheet.freezePane?.rows ?? 0,
		frozenCols: sheet.freezePane?.columns ?? 0,
		structureLocked: lockStructure || structureLocked(sheet),
		...(textRows.length > 0 || textCols.length > 0
			? { textFormat: { rows: textRows, cols: textCols, otherRows: textCols.length > 0 ? formattedRows : [], otherCells: unshown } }
			: {})
	}
}

// The used area: the value rectangle and every merge.
export function sheetExtent(sheet: Sheet): { rowCount: number; colCount: number } {
	let rowCount = sheet.rows.length
	let colCount = sheet.rows[0]?.length ?? 0

	for (const merge of sheet.merges ?? []) {
		rowCount = Math.max(rowCount, merge.endRow + 1)
		colCount = Math.max(colCount, merge.endCol + 1)
	}

	return { rowCount, colCount }
}

// Whether anything in the sheet names cell ranges the workbook model would not shift with an inserted or
// deleted row or column.
export function structureLocked(sheet: Sheet): boolean {
	return (
		(sheet.tables?.length ?? 0) > 0 ||
		(sheet.conditionalRules?.length ?? 0) > 0 ||
		(sheet.dataValidations?.length ?? 0) > 0 ||
		(sheet.images?.length ?? 0) > 0 ||
		sheet.autoFilter !== undefined ||
		// Kept by range as read (a number stored as text Excel is told not to flag).
		(sheet.ignoredErrors?.length ?? 0) > 0 ||
		(sheet.sparklines?.length ?? 0) > 0 ||
		(sheet.charts?.length ?? 0) > 0 ||
		// Kept in their own part, which saving copies as read while the notes they pair with move.
		(sheet.threadedComments?.length ?? 0) > 0 ||
		// Print ranges (kept in the page setup) are written back as read.
		sheet.pageSetup?.printArea !== undefined ||
		sheet.pageSetup?.printTitlesRow !== undefined ||
		sheet.pageSetup?.printTitlesColumn !== undefined ||
		hasArrayFormulas(sheet)
	)
}

const arraySheets = new WeakMap<Sheet, boolean>()

// Whether the sheet holds array formulas (legacy or spilling), whose ranges saving writes as read: they
// cannot move with inserted or deleted rows. Looked up once per sheet, at its first view: edits never add
// one, and a sheet that loses its last stays locked.
export function hasArrayFormulas(sheet: Sheet): boolean {
	let found = arraySheets.get(sheet)

	if (found === undefined) {
		found = false

		for (const cell of sheet.cells?.values() ?? []) {
			if (cell.formulaType === "array") {
				found = true

				break
			}
		}

		arraySheets.set(sheet, found)
	}

	return found
}

// Whether something names ranges on any sheet, so no sheet's rows or columns may move: defined names, and
// charts, pivot tables, slicers and timelines, whose source ranges can sit on any sheet.
export function workbookStructureLocked(workbook: Workbook): boolean {
	return (
		(workbook.namedRanges?.length ?? 0) > 0 ||
		workbook.sheets.some(
			sheet =>
				!isWorksheet(sheet) ||
				(sheet.charts?.length ?? 0) > 0 ||
				(sheet.pivotTables?.length ?? 0) > 0 ||
				(sheet.slicers?.length ?? 0) > 0 ||
				(sheet.timelines?.length ?? 0) > 0
		)
	)
}

// A sheet not built yet: its name and used area, held structure-locked until its view comes.
export function deferredSheetView(sheet: Sheet): SheetView {
	return {
		name: sheet.name,
		...sheetExtent(sheet),
		cells: new Map(),
		merges: [],
		colWidths: new Map(),
		rowHeights: new Map(),
		hiddenCols: [],
		hiddenRows: [],
		frozenRows: 0,
		frozenCols: 0,
		structureLocked: true,
		deferred: true
	}
}

// The grid index of the workbook's active sheet: the file names its active tab among all of them, the grid
// counts worksheets only.
export function activeWorksheet(workbook: Workbook): number {
	const worksheets = workbook.sheets.filter(isWorksheet).length
	const active = workbook.sheets.slice(0, (workbook.activeSheet ?? 0) + 1).filter(isWorksheet).length - 1

	return Math.min(Math.max(active, 0), Math.max(worksheets - 1, 0))
}

// Every sheet's view, or (`activeOnly`) the active sheet's alone with the rest deferred, so an open costs
// what the sheet on show holds.
export function workbookDoc(workbook: Workbook, views: WorkbookViews, writable: boolean, activeOnly = false): SpreadsheetDoc {
	const lockStructure = workbookStructureLocked(workbook)
	const activeSheet = activeWorksheet(workbook)
	const sheets = workbook.sheets
		.filter(isWorksheet)
		.map((sheet, index) => (activeOnly && index !== activeSheet ? deferredSheetView(sheet) : views.sheet(sheet, lockStructure)))

	return { kind: "xlsx", sheets, activeSheet, styles: views.styles.styles, writable }
}

// Chart sheets and the like show as empty tabs in the workbook model; the grid lists worksheets only.
export function isWorksheet(sheet: Sheet): boolean {
	return sheet.kind === undefined || sheet.kind === "worksheet"
}
