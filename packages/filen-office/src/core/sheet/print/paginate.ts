import type { SheetRange } from "../protocol"
import type { SheetPageSetup, SheetPrintSpan } from "./setup"

// A sheet laid out on pages, as Excel does it: each range printed (a print area's ranges one after another, else the
// used range) is cut into bands of rows and of columns that fit the paper inside the margins at the print scale,
// starting a band at each manual break; the title rows and columns are repeated on every page that starts past them;
// pages go down then over (or over then down). Fit to W by H pages finds the largest whole percent, at most 100 and
// at least 10, at which every range takes at most that many pages across and down; it ignores manual breaks, as
// Excel's Fit To option does ("Insert, move, or delete page breaks in a worksheet", support.microsoft.com).
//
// Sizes are points at 100 %. Nothing here walks more than the rows and columns printed.

export interface PrintAxis {
	readonly count: number
	// Points at 100 %; 0 for a hidden row or column.
	size(index: number): number
}

export interface PrintSheet {
	readonly rows: PrintAxis
	readonly cols: PrintAxis
	readonly setup: SheetPageSetup
	// What prints, in order (a print area's ranges, the selection's, or the used range).
	readonly ranges: readonly SheetRange[]
}

export interface PrintPage {
	// Which of the sheet's ranges the page belongs to.
	readonly range: number
	readonly rows: SheetPrintSpan
	readonly cols: SheetPrintSpan
	// The title rows and columns repeated above and left of the page's own (null: none).
	readonly titleRows: SheetPrintSpan | null
	readonly titleCols: SheetPrintSpan | null
}

export interface SheetPagination {
	// The scale the pages print at (1: 100 %).
	readonly scale: number
	// The paper as it lies (landscape turned), in points.
	readonly paper: { readonly width: number; readonly height: number }
	// The area inside the margins, in points.
	readonly area: { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
	// The headings' size at 100 % (0 without headings): the row numbers' width and the column letters' height.
	readonly headingWidth: number
	readonly headingHeight: number
	readonly pages: readonly PrintPage[]
	// Laid out up to the cap it was given: more pages would follow.
	readonly more: boolean
}

const POINTS_PER_INCH = 72
// The headings: Excel's column letters take a default row's height, the row numbers the width of the widest number.
const HEADING_HEIGHT = 15
const DIGIT_POINTS = 5.625
const HEADING_PADDING = 9
const HEADING_MIN = 24
// Past this many pages a sheet is not laid out further (a print of a whole empty grid at 10 %).
export const MAX_PRINT_PAGES = 10_000

export function headingWidth(lastRow: number): number {
	return Math.max(HEADING_MIN, Math.ceil(String(lastRow + 1).length * DIGIT_POINTS + HEADING_PADDING))
}

export function orientedPaper(setup: SheetPageSetup): { readonly width: number; readonly height: number } {
	const { width, height } = setup.paper

	return setup.orientation === "landscape" ? { width: Math.max(width, height), height: Math.min(width, height) } : { width, height }
}

// The ranges' bands along one axis: [first, last] pairs, each fitting `room` (the room at `start` depends on how many
// titles repeat there).
function bands(
	axis: PrintAxis,
	first: number,
	last: number,
	room: (start: number) => number,
	breaks: readonly number[],
	limit: number
): SheetPrintSpan[] {
	const out: SheetPrintSpan[] = []
	let start = first
	let next = 0

	while (start <= last && out.length < limit) {
		while (next < breaks.length && (breaks[next] ?? 0) <= start) next++

		const stop = Math.min(last, next < breaks.length ? (breaks[next] ?? 0) - 1 : last)
		const capacity = room(start)
		let used = 0
		let end = start - 1
		let placed = false

		for (let index = start; index <= stop; index++) {
			const size = axis.size(index)

			if (size > 0 && placed && used + size > capacity + 1e-9) break

			used += size
			end = index

			if (size > 0) placed = true
		}

		// A band of hidden indices alone prints nothing.
		if (placed) out.push({ first: start, last: end })

		start = end + 1
	}

	return out
}

// The title span repeated on a page starting at `start`: the titles before it.
function repeated(titles: SheetPrintSpan | null, start: number): SheetPrintSpan | null {
	if (titles === null || start <= titles.first) return null

	return { first: titles.first, last: Math.min(titles.last, start - 1) }
}

function spanSize(axis: PrintAxis, span: SheetPrintSpan | null): number {
	if (span === null) return 0

	let total = 0

	for (let index = span.first; index <= span.last; index++) total += axis.size(index)

	return total
}

interface Laid {
	readonly rows: SheetPrintSpan[]
	readonly cols: SheetPrintSpan[]
}

function layOut(
	sheet: PrintSheet,
	range: SheetRange,
	scale: number,
	area: { width: number; height: number },
	manual: boolean,
	limit: number
): Laid {
	const { setup, rows, cols } = sheet
	const heading = setup.headings ? headingWidth(range.endRow) : 0
	const headingTop = setup.headings ? HEADING_HEIGHT : 0
	const titleRows = setup.titleRows
	const titleCols = setup.titleCols
	// Repeated titles and their sizes, by the band's start (most bands start past them all).
	const rowRoom = (start: number): number => area.height / scale - headingTop - spanSize(rows, repeated(titleRows, start))
	const colRoom = (start: number): number => area.width / scale - heading - spanSize(cols, repeated(titleCols, start))

	return {
		rows: bands(rows, range.startRow, Math.min(range.endRow, rows.count - 1), rowRoom, manual ? setup.rowBreaks : [], limit),
		cols: bands(cols, range.startCol, Math.min(range.endCol, cols.count - 1), colRoom, manual ? setup.colBreaks : [], limit)
	}
}

function printArea(sheet: PrintSheet): { paper: { width: number; height: number }; area: SheetPagination["area"] } {
	const paper = orientedPaper(sheet.setup)
	const { margins } = sheet.setup
	const x = margins.left * POINTS_PER_INCH
	const y = margins.top * POINTS_PER_INCH

	return {
		paper,
		area: {
			x,
			y,
			width: Math.max(POINTS_PER_INCH / 2, paper.width - x - margins.right * POINTS_PER_INCH),
			height: Math.max(POINTS_PER_INCH / 2, paper.height - y - margins.bottom * POINTS_PER_INCH)
		}
	}
}

// A chartsheet's pagination: one page, its paper and the area inside its margins, at 100 % (a chart sheet's print
// has no scaling: Microsoft Support, "Print a chart").
export function chartPagination(setup: SheetPageSetup): SheetPagination {
	const { paper, area } = printArea({ setup, ranges: [], rows: { count: 0, size: () => 0 }, cols: { count: 0, size: () => 0 } })
	const span = { first: 0, last: 0 }

	return {
		scale: 1,
		paper,
		area,
		headingWidth: 0,
		headingHeight: 0,
		pages: [{ range: 0, rows: span, cols: span, titleRows: null, titleCols: null }],
		more: false
	}
}

// Whether every range fits `wide` by `tall` pages at `percent` (0: any number on that axis).
function fits(sheet: PrintSheet, percent: number, area: { width: number; height: number }, wide: number, tall: number): boolean {
	for (const range of sheet.ranges) {
		// One band past the limit is enough to know it does not fit.
		const laid = layOut(sheet, range, percent / 100, area, false, Math.max(wide, tall) + 1)

		if (wide > 0 && laid.cols.length > wide) return false
		if (tall > 0 && laid.rows.length > tall) return false
	}

	return true
}

// The fit-to scale in whole percent.
export function fitScale(sheet: PrintSheet): number {
	const { setup } = sheet
	const { area } = printArea(sheet)
	const wide = setup.fitToWidth
	const tall = setup.fitToHeight

	if (wide === 0 && tall === 0) return 100
	if (fits(sheet, 100, area, wide, tall)) return 100

	let low = 10
	let high = 100

	// The largest percent that fits: fewer pages never need a larger scale.
	while (low < high) {
		const middle = Math.ceil((low + high) / 2)

		if (fits(sheet, middle, area, wide, tall)) low = middle
		else high = middle - 1
	}

	return low
}

// `maxPages`: the pages laid out at most (bands past them are not walked; `more` says some were left).
export function paginate(sheet: PrintSheet, maxPages: number = MAX_PRINT_PAGES): SheetPagination {
	const { setup } = sheet
	const { paper, area } = printArea(sheet)
	const percent = setup.fitToPage ? fitScale(sheet) : setup.scale
	const scale = percent / 100
	const cap = Math.max(0, Math.min(MAX_PRINT_PAGES, Math.floor(maxPages)))
	const pages: PrintPage[] = []
	let lastRow = 0
	let more = false

	sheet.ranges.forEach((range, index) => {
		if (pages.length >= cap) {
			if (!more) {
				const probe = layOut(sheet, range, scale, area, !setup.fitToPage, 1)

				more = probe.rows.length > 0 && probe.cols.length > 0
			}

			return
		}

		// One band past what is left tells whether more would follow.
		const laid = layOut(sheet, range, scale, area, !setup.fitToPage, cap - pages.length + 1)
		const page = (rows: SheetPrintSpan, cols: SheetPrintSpan): void => {
			if (pages.length >= cap) {
				more = true
				return
			}

			pages.push({
				range: index,
				rows,
				cols,
				titleRows: repeated(setup.titleRows, rows.first),
				titleCols: repeated(setup.titleCols, cols.first)
			})
		}

		lastRow = Math.max(lastRow, range.endRow)

		if (setup.pageOrder === "overThenDown") {
			for (const rows of laid.rows) for (const cols of laid.cols) page(rows, cols)
		} else {
			for (const cols of laid.cols) for (const rows of laid.rows) page(rows, cols)
		}
	})

	return {
		scale,
		paper,
		area,
		headingWidth: setup.headings ? headingWidth(lastRow) : 0,
		headingHeight: setup.headings ? HEADING_HEIGHT : 0,
		pages,
		more
	}
}
