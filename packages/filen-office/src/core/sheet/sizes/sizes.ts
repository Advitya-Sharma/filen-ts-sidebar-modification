import { DEFAULT_ROW_HEIGHT, type CellRange } from "../model/model"

// Column and row sizes: the grid's pixels, the xlsx's own units, the limits both share, and the edits
// resizing makes to a sheet's size maps. Column widths are stored as characters of the default font's widest
// digit with 5 px of padding counted in (ECMA-376 Part 1 §18.3.1.13), row heights in points.
// The widest digit of the workbook's default font (its Normal style's) at 96 dpi, the unit of column widths: Calibri
// 11's is 7 px (workbookDigitPx).
export const DEFAULT_DIGIT_PX = 7
const CHAR_PADDING = 5
export const PX_PER_POINT = 4 / 3

// The widest advance of the digits 0–9, in ems, of the bundled faces whose widths equal the fonts documents name
// (core/fonts/substitutes.ts): Carlito for Calibri, Arimo for Arial and Helvetica, Tinos for Times New Roman, Cousine
// for Courier New, Gelasio for Georgia. Read from the faces' hmtx (assets/fonts; a test reads them again).
const DIGIT_EMS: ReadonlyMap<string, number> = new Map([
	["calibri", 1038 / 2048],
	["arial", 1139 / 2048],
	["helvetica", 1139 / 2048],
	["times new roman", 1024 / 2048],
	["courier new", 1229 / 2048],
	["georgia", 1257 / 2048]
])
const DEFAULT_DIGIT_EM = 1038 / 2048
const DEFAULT_FONT_POINTS = 11

// Font sizes go to 409 points (Excel's limit): the widest digit any font of a known width has there.
export const MAX_DIGIT_PX = Math.ceil(409 * PX_PER_POINT * Math.max(DEFAULT_DIGIT_EM, ...DIGIT_EMS.values()))

// The unit Excel itself used at 96 dpi, read back from the widths it wrote (§18.3.1.13: a width is
// Truncate(px / digit × 256) / 256 for whole pixels px, so its fraction tells the digit). Per Normal font and size,
// from the corpus's files whose docProps/app.xml names Microsoft Excel: accepted where at least two independent
// sources agree on the smallest digit their widths allow, and that digit is within a pixel of the face-based estimate
// below (files saved at 125% or 150% display scaling show larger digits: Calibri 11 as 9 and 11, Arial 11 only as 12,
// left out).
//   Calibri 11: 7 (49 of 76 files; the rest 9, 11, 15+)   Arial 10: 7 (53 of 69)   Arial 12: 9 (8 of 10)
//   Aptos Narrow 11: 7 (2 of 10, both gov.uk; the rest 8+) Verdana 10: 8 (2 of 3; the third 11)
//   Calibri 8: 6 (2 of 2, two Office templates; Excel's digits are hinted at whole pixels, a pixel wider here than
//   the advances, and LibreOffice, give)
// Left out for want of a second source: Calibri 9, 7 in four copies of one template (LibreOffice and the advances
// give 6); Cambria 11, 8 in one file.
const EXCEL_DIGITS: ReadonlyMap<string, number> = new Map([
	["calibri 11", 7],
	["calibri 8", 6],
	["arial 10", 7],
	["arial 12", 9],
	["aptos narrow 11", 7],
	["verdana 10", 8]
])

// The workbook's column-width unit from its default font, the same on every device whatever fonts it has: Excel's
// own where its files show it (EXCEL_DIGITS), else the widest digit's advance at the font's size in CSS pixels
// (points × 4/3, unrounded), to whole pixels, as LibreOffice 7.5 lays its columns out. A font with neither Excel's
// evidence nor a face of equal widths here is counted in Calibri's digits at its own size: a Cambria or an Aptos
// workbook's columns keep their character counts' proportions, unmeasured.
export function workbookDigitPx(font: string | undefined, points: number | undefined): number {
	const family = (font ?? "Calibri").trim().toLowerCase()
	const em = DIGIT_EMS.get(family) ?? DEFAULT_DIGIT_EM
	const size = points !== undefined && Number.isFinite(points) && points > 0 ? Math.min(points, 409) : DEFAULT_FONT_POINTS
	const excel = EXCEL_DIGITS.get(`${family} ${String(size)}`)

	if (excel !== undefined) return excel

	return Math.max(1, Math.min(MAX_DIGIT_PX, Math.round(em * size * PX_PER_POINT)))
}

export type SizeAxis = "rows" | "cols"

// An index's new pixel size; null takes it back to the file's (or the default) size.
export type SizeEntry = readonly [index: number, px: number | null]

// Excel's own limits (a stored width of 255, MS-OI29500 2.1.598; 409 points), in pixels. Never 0: resizing does
// not hide.
export const MIN_SIZE = 8
export const MAX_COL_WIDTH = 1785
export const MAX_ROW_HEIGHT = 545

// The most indices one resize names. A whole selection larger than this (select-all on a large sheet)
// resizes only the dragged one, so no gesture builds a size map the size of the sheet.
export const MAX_RESIZE_TARGETS = 10_000

// A cell's horizontal padding (px-1.5 each side) and its gridline.
export const CELL_PADDING = 13
// What autofit adds to measured content: a cell's padding and 1px of slack; a row's line box to the
// default row's height.
const CELL_FIT_PADDING = CELL_PADDING + 1
const ROW_FIT_PADDING = 7

export function clampSize(axis: SizeAxis, px: number): number {
	return Math.round(Math.min(axis === "cols" ? MAX_COL_WIDTH : MAX_ROW_HEIGHT, Math.max(MIN_SIZE, px)))
}

// A stored width's pixels: Truncate(((256 * width + Truncate(128 / digit)) / 256) * digit) (§18.3.1.13).
export function colWidthToPx(width: number, digit = DEFAULT_DIGIT_PX): number {
	return Math.trunc(((256 * width + Math.trunc(128 / digit)) * digit) / 256)
}

// The width stored for pixels, to 1/256 of a character (§18.3.1.13, with (px - 5) / digit characters).
export function pxToColWidth(px: number, digit = DEFAULT_DIGIT_PX): number {
	return Math.trunc((px * 256) / digit) / 256
}

export interface ColWidthFormat {
	// Stored width, NaN when absent.
	readonly defaultColWidth: number
	// Characters, padding not counted.
	readonly baseColWidth: number
}

// The default column's pixels: defaultColWidth's, else baseColWidth characters with their padding (§18.3.1.81),
// which Excel rounds up to whole 8 px (MS-OI29500 2.1.598 f): 64 px for 8, the "8.43 characters" Excel shows.
export function defaultColPx(format: ColWidthFormat, digit = DEFAULT_DIGIT_PX): number {
	if (!Number.isNaN(format.defaultColWidth)) return colWidthToPx(format.defaultColWidth, digit)

	return Math.ceil((format.baseColWidth * digit + CHAR_PADDING) / 8) * 8
}

// The width a <col> of the default column's size stores.
export function defaultColWidth(format: ColWidthFormat, digit = DEFAULT_DIGIT_PX): number {
	return Number.isNaN(format.defaultColWidth) ? pxToColWidth(defaultColPx(format, digit), digit) : format.defaultColWidth
}

export function rowHeightToPx(points: number): number {
	return Math.round(points * PX_PER_POINT)
}

export function pxToRowHeight(px: number): number {
	return px / PX_PER_POINT
}

export function withSizes(sizes: ReadonlyMap<number, number>, entries: readonly SizeEntry[]): Map<number, number> {
	const next = new Map(sizes)

	for (const [index, px] of entries) {
		if (px === null) {
			next.delete(index)
		} else {
			next.set(index, px)
		}
	}

	return next
}

export function sheetWithSizes<T extends { colWidths: ReadonlyMap<number, number>; rowHeights: ReadonlyMap<number, number> }>(
	sheet: T,
	axis: SizeAxis,
	entries: readonly SizeEntry[]
): T {
	return axis === "cols"
		? { ...sheet, colWidths: withSizes(sheet.colWidths, entries) }
		: { ...sheet, rowHeights: withSizes(sheet.rowHeights, entries) }
}

// Where a row or column index lands once `count` of them are inserted or deleted at `at` (null: deleted).
export function shiftIndex(index: number, kind: "insert" | "delete", at: number, count: number): number | null {
	if (index < at) {
		return index
	}

	if (kind === "insert") {
		return index + count
	}

	return index < at + count ? null : index - count
}

export function shiftSizes(
	sizes: ReadonlyMap<number, number>,
	kind: "insert" | "delete",
	at: number,
	count: number
): { sizes: Map<number, number>; removed: [number, number][] } {
	const next = new Map<number, number>()
	const removed: [number, number][] = []

	for (const [index, px] of sizes) {
		const moved = shiftIndex(index, kind, at, count)

		if (moved === null) {
			removed.push([index - at, px])
		} else {
			next.set(moved, px)
		}
	}

	return { sizes: next, removed }
}

export function restoreSizes(sizes: ReadonlyMap<number, number>, at: number, removed: readonly [number, number][]): Map<number, number> {
	const next = new Map(sizes)

	for (const [offset, px] of removed) {
		next.set(at + offset, px)
	}

	return next
}

// Which columns or rows a resize of `dragged` applies to: every visible selected one when the selection
// covers whole columns or rows and holds the dragged one, else just the dragged one.
export function resizeTargets(
	axis: SizeAxis,
	dragged: number,
	range: CellRange,
	whole: boolean,
	isHidden: (index: number) => boolean
): number[] {
	const [start, end] = axis === "cols" ? [range.startCol, range.endCol] : [range.startRow, range.endRow]

	if (!whole || dragged < start || dragged > end || end - start + 1 > MAX_RESIZE_TARGETS) {
		return [dragged]
	}

	const targets: number[] = []

	for (let index = start; index <= end; index++) {
		if (!isHidden(index)) {
			targets.push(index)
		}
	}

	return targets
}

// Autofit from the measured content of an index's rendered cells: null (the default) with nothing to
// measure, or for a row whose content fits the default height.
export function fitSize(axis: SizeAxis, contents: readonly number[]): number | null {
	if (contents.length === 0) {
		return null
	}

	const largest = Math.ceil(Math.max(...contents))

	if (axis === "cols") {
		return clampSize("cols", largest + CELL_FIT_PADDING)
	}

	const height = clampSize("rows", largest + ROW_FIT_PADDING)

	return height <= DEFAULT_ROW_HEIGHT ? null : height
}

// What a Reset of the selected columns or rows applies to: the selection cut at the used area (select-all
// resets what is on the sheet, not a million rows), never a hidden one (its size is kept for when it is
// shown again), at most the cap.
export function resetTargets(start: number, end: number, used: number, isHidden: (index: number) => boolean): number[] {
	const last = Math.min(end, Math.max(start, used - 1))
	const targets: number[] = []

	for (let index = start; index <= last && targets.length < MAX_RESIZE_TARGETS; index++) {
		if (!isHidden(index)) {
			targets.push(index)
		}
	}

	return targets
}
