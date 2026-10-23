import { encodeUtf8, parseXml, type XmlDocument } from "../../xml/index"
import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetRange } from "../protocol"
import { columnIndex } from "../refs/cellRef"
import { LETTER, paperPoints, paperSizeKnown, universalMeasure } from "./paper"

// A worksheet's print settings as the file holds them (ECMA-376 Part 1 §18.3.1): pageSetup (§18.3.1.63), pageMargins
// (§18.3.1.62), printOptions (§18.3.1.70), headerFooter (§18.3.1.46), rowBreaks and colBreaks (§18.3.1.74,
// §18.3.1.14), sheetPr's pageSetUpPr fitToPage (§18.3.1.65), and the sheet's _xlnm.Print_Area and _xlnm.Print_Titles
// names (§18.2.6). An attribute a file leaves out takes its schema default.

export interface SheetPrintMargins {
	// Inches.
	readonly left: number
	readonly right: number
	readonly top: number
	readonly bottom: number
	readonly header: number
	readonly footer: number
}

// Excel's "Normal" margins, what a sheet without pageMargins prints with.
export const NORMAL_MARGINS: SheetPrintMargins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }

export interface SheetHeaderFooterTexts {
	readonly header: string
	readonly footer: string
}

export interface SheetHeaderFooter {
	readonly differentFirst: boolean
	readonly differentOddEven: boolean
	readonly scaleWithDoc: boolean
	readonly alignWithMargins: boolean
	readonly odd: SheetHeaderFooterTexts
	readonly even: SheetHeaderFooterTexts
	readonly first: SheetHeaderFooterTexts
}

export interface SheetPrintSpan {
	readonly first: number
	readonly last: number
}

export interface SheetPageSetup {
	readonly orientation: "portrait" | "landscape"
	// ST_PaperSize as written (§18.3.1.63), 1 when absent.
	readonly paperSize: number
	// The paper, portrait, in points: paperWidth and paperHeight when both read, else paperSize's.
	readonly paper: { readonly width: number; readonly height: number }
	// Custom paperWidth and paperHeight are set (paperSize is then ignored).
	readonly customPaper: boolean
	// Percent, 10 to 400.
	readonly scale: number
	// Fit to fitToWidth by fitToHeight pages instead of scaling (0: as many as it takes on that axis).
	readonly fitToPage: boolean
	readonly fitToWidth: number
	readonly fitToHeight: number
	readonly pageOrder: "downThenOver" | "overThenDown"
	// The first page's number, or null for automatic (1).
	readonly firstPageNumber: number | null
	readonly blackAndWhite: boolean
	readonly draft: boolean
	readonly errors: "displayed" | "blank" | "dash" | "NA"
	// Notes: not printed, printed where shown, or listed after the sheet (pageSetup cellComments, §18.18.5).
	readonly cellComments: "none" | "asDisplayed" | "atEnd"
	readonly margins: SheetPrintMargins
	readonly gridLines: boolean
	readonly headings: boolean
	readonly horizontalCentered: boolean
	readonly verticalCentered: boolean
	readonly headerFooter: SheetHeaderFooter
	// Manual breaks: the rows (columns) a new page starts at, ascending.
	readonly rowBreaks: readonly number[]
	readonly colBreaks: readonly number[]
	// The sheet's print area (one range per area, printed one after another), or null for the used range.
	readonly printArea: readonly SheetRange[] | null
	// Rows and columns repeated on every page (print titles).
	readonly titleRows: SheetPrintSpan | null
	readonly titleCols: SheetPrintSpan | null
}

const NO_TEXTS: SheetHeaderFooterTexts = { header: "", footer: "" }

export const DEFAULT_HEADER_FOOTER: SheetHeaderFooter = {
	differentFirst: false,
	differentOddEven: false,
	scaleWithDoc: true,
	alignWithMargins: true,
	odd: NO_TEXTS,
	even: NO_TEXTS,
	first: NO_TEXTS
}

export const DEFAULT_PAGE_SETUP: SheetPageSetup = {
	orientation: "portrait",
	paperSize: LETTER,
	paper: paperPoints(LETTER),
	customPaper: false,
	scale: 100,
	fitToPage: false,
	fitToWidth: 1,
	fitToHeight: 1,
	pageOrder: "downThenOver",
	firstPageNumber: null,
	blackAndWhite: false,
	draft: false,
	errors: "displayed",
	cellComments: "none",
	margins: NORMAL_MARGINS,
	gridLines: false,
	headings: false,
	horizontalCentered: false,
	verticalCentered: false,
	headerFooter: DEFAULT_HEADER_FOOTER,
	rowBreaks: [],
	colBreaks: [],
	printArea: null,
	titleRows: null,
	titleCols: null
}

// The worksheet markup a sheet's setup is read from (each section's occurrences, as they now are).
export interface SheetPrintMarkup {
	readonly pageSetup: string
	readonly pageMargins: string
	readonly printOptions: string
	readonly headerFooter: string
	readonly rowBreaks: string
	readonly colBreaks: string
	readonly sheetPr: string
}

function rootOf(markup: string): { readonly document: XmlDocument; readonly root: number } | null {
	if (markup.trim() === "") return null

	try {
		const document = parseXml(encodeUtf8(markup))

		return { document, root: document.documentElement() }
	} catch {
		return null
	}
}

function bool(value: string | undefined, fallback: boolean): boolean {
	if (value === undefined) return fallback

	return value === "1" || value === "true" ? true : value === "0" || value === "false" ? false : fallback
}

function unsigned(value: string | undefined, fallback: number): number {
	if (value === undefined || !/^\s*[0-9]+\s*$/.test(value)) return fallback

	return Number(value)
}

function double(value: string | undefined, fallback: number): number {
	if (value === undefined) return fallback

	const read = Number(value)

	return value.trim() !== "" && Number.isFinite(read) && read >= 0 ? read : fallback
}

function oneOf<T extends string>(value: string | undefined, allowed: readonly T[], fallback: T): T {
	return value !== undefined && (allowed as readonly string[]).includes(value) ? (value as T) : fallback
}

function readBreaks(markup: string, limit: number): number[] {
	const found = rootOf(markup)

	if (found === null) return []

	const { document, root } = found
	const out: number[] = []

	for (let brk = document.firstChildElement(root); brk >= 0; brk = document.nextSiblingElement(brk)) {
		if (document.localName(brk) !== "brk") continue
		// Breaks the application placed itself are worked out again; only manual ones are kept.
		if (!bool(document.attribute(brk, "man"), false)) continue

		const id = unsigned(document.attribute(brk, "id"), 0)

		// A break above the first row (left of the first column) starts no page.
		if (id > 0 && id < limit) out.push(id)
	}

	return [...new Set(out)].sort((a, b) => a - b)
}

function headerFooterOf(markup: string): SheetHeaderFooter {
	const found = rootOf(markup)

	if (found === null) return DEFAULT_HEADER_FOOTER

	const { document, root } = found
	const texts: Record<string, string> = {}

	for (let child = document.firstChildElement(root); child >= 0; child = document.nextSiblingElement(child))
		texts[document.localName(child)] = document.textContent(child)

	return {
		differentFirst: bool(document.attribute(root, "differentFirst"), false),
		differentOddEven: bool(document.attribute(root, "differentOddEven"), false),
		scaleWithDoc: bool(document.attribute(root, "scaleWithDoc"), true),
		alignWithMargins: bool(document.attribute(root, "alignWithMargins"), true),
		odd: { header: texts["oddHeader"] ?? "", footer: texts["oddFooter"] ?? "" },
		even: { header: texts["evenHeader"] ?? "", footer: texts["evenFooter"] ?? "" },
		first: { header: texts["firstHeader"] ?? "", footer: texts["firstFooter"] ?? "" }
	}
}

// sheetPr's pageSetUpPr fitToPage, as written.
export function readFitToPage(sheetPr: string): boolean {
	const found = rootOf(sheetPr)

	if (found === null) return false

	const { document, root } = found

	for (let child = document.firstChildElement(root); child >= 0; child = document.nextSiblingElement(child))
		if (document.localName(child) === "pageSetUpPr") return bool(document.attribute(child, "fitToPage"), false)

	return false
}

export interface SheetPrintNames {
	readonly printArea: string | null
	readonly printTitles: string | null
}

// A custom paper's sides, in points: a PDF page is at least 3 and at most 14,400 units a side (ISO 32000-1, Annex C.2),
// and a print is laid out on the paper the PDF can hold.
const MIN_PAPER = 3
const MAX_PAPER = 14_400

function paperSide(measure: number | null): number | null {
	return measure === null ? null : Math.min(MAX_PAPER, Math.max(MIN_PAPER, measure))
}

export function readPageSetup(markup: SheetPrintMarkup, names: SheetPrintNames, fitToPage?: boolean): SheetPageSetup {
	const page = rootOf(markup.pageSetup)
	const attribute = (name: string): string | undefined => (page === null ? undefined : page.document.attribute(page.root, name))
	const width = paperSide(universalMeasure(attribute("paperWidth")))
	const height = paperSide(universalMeasure(attribute("paperHeight")))
	const paperSize = unsigned(attribute("paperSize"), LETTER)
	const custom = width !== null && height !== null
	const margins = rootOf(markup.pageMargins)
	const margin = (name: keyof SheetPrintMargins): number =>
		margins === null ? NORMAL_MARGINS[name] : double(margins.document.attribute(margins.root, name), NORMAL_MARGINS[name])
	const options = rootOf(markup.printOptions)
	const option = (name: string): boolean => (options === null ? false : bool(options.document.attribute(options.root, name), false))
	const scale = unsigned(attribute("scale"), 100)
	const titles = names.printTitles === null ? null : readTitles(names.printTitles)

	return {
		orientation: attribute("orientation") === "landscape" ? "landscape" : "portrait",
		paperSize,
		paper: custom ? { width, height } : paperPoints(paperSizeKnown(paperSize) ? paperSize : LETTER),
		customPaper: custom,
		scale: Math.min(400, Math.max(10, scale)),
		fitToPage: fitToPage ?? readFitToPage(markup.sheetPr),
		fitToWidth: unsigned(attribute("fitToWidth"), 1),
		fitToHeight: unsigned(attribute("fitToHeight"), 1),
		pageOrder: oneOf(attribute("pageOrder"), ["downThenOver", "overThenDown"], "downThenOver"),
		firstPageNumber: bool(attribute("useFirstPageNumber"), false) ? unsigned(attribute("firstPageNumber"), 1) : null,
		blackAndWhite: bool(attribute("blackAndWhite"), false),
		draft: bool(attribute("draft"), false),
		errors: oneOf(attribute("errors"), ["displayed", "blank", "dash", "NA"], "displayed"),
		cellComments: oneOf(attribute("cellComments"), ["none", "asDisplayed", "atEnd"], "none"),
		margins: {
			left: margin("left"),
			right: margin("right"),
			top: margin("top"),
			bottom: margin("bottom"),
			header: margin("header"),
			footer: margin("footer")
		},
		gridLines: option("gridLines"),
		headings: option("headings"),
		horizontalCentered: option("horizontalCentered"),
		verticalCentered: option("verticalCentered"),
		headerFooter: headerFooterOf(markup.headerFooter),
		rowBreaks: readBreaks(markup.rowBreaks, SHEET_MAX_ROWS),
		colBreaks: readBreaks(markup.colBreaks, SHEET_MAX_COLS),
		printArea: names.printArea === null ? null : readAreas(names.printArea),
		titleRows: titles?.rows ?? null,
		titleCols: titles?.cols ?? null
	}
}

// --- Print area and titles ---

// A defined name's formula split at its top-level commas (a quoted sheet name may hold one), each part without its
// sheet.
function referenceParts(formula: string): string[] {
	const parts: string[] = []
	let current = ""
	let quoted = false

	for (const char of formula) {
		if (char === "'") quoted = !quoted
		if (char === "," && !quoted) {
			parts.push(current)
			current = ""

			continue
		}

		current += char
	}

	parts.push(current)

	return parts.map(part => {
		const bang = part.lastIndexOf("!")

		return (bang < 0 ? part : part.slice(bang + 1)).replace(/\$/g, "").trim()
	})
}

const CELLS = /^([A-Za-z]{1,3})([0-9]{1,7})(?::([A-Za-z]{1,3})([0-9]{1,7}))?$/
const ROWS = /^([0-9]{1,7}):([0-9]{1,7})$/
const COLUMNS = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/

// One reference as a range (whole rows and columns included), or null.
function readReference(text: string): SheetRange | null {
	const cells = CELLS.exec(text)

	if (cells !== null) {
		const startCol = columnIndex(cells[1] ?? "")
		const startRow = Number(cells[2]) - 1
		const endCol = cells[3] === undefined ? startCol : columnIndex(cells[3])
		const endRow = cells[4] === undefined ? startRow : Number(cells[4]) - 1

		return checked(startRow, startCol, endRow, endCol)
	}

	const rows = ROWS.exec(text)

	if (rows !== null) return checked(Number(rows[1]) - 1, 0, Number(rows[2]) - 1, SHEET_MAX_COLS - 1)

	const columns = COLUMNS.exec(text)

	if (columns !== null) return checked(0, columnIndex(columns[1] ?? ""), SHEET_MAX_ROWS - 1, columnIndex(columns[2] ?? ""))

	return null
}

function checked(startRow: number, startCol: number, endRow: number, endCol: number): SheetRange | null {
	const range = {
		startRow: Math.min(startRow, endRow),
		startCol: Math.min(startCol, endCol),
		endRow: Math.max(startRow, endRow),
		endCol: Math.max(startCol, endCol)
	}

	return range.startRow < 0 || range.startCol < 0 || range.endRow >= SHEET_MAX_ROWS || range.endCol >= SHEET_MAX_COLS ? null : range
}

// The print area's ranges; null when a part does not read (#REF! after its cells were deleted): the used range
// prints, as Excel prints a sheet whose print area is invalid.
export function readAreas(formula: string): SheetRange[] | null {
	const ranges: SheetRange[] = []

	for (const part of referenceParts(formula)) {
		const range = readReference(part)

		if (range === null) return null

		ranges.push(range)
	}

	return ranges.length === 0 ? null : ranges
}

// Print titles: whole rows, whole columns, or one of each (§18.2.5's _xlnm.Print_Titles).
export function readTitles(formula: string): { readonly rows: SheetPrintSpan | null; readonly cols: SheetPrintSpan | null } {
	let rows: SheetPrintSpan | null = null
	let cols: SheetPrintSpan | null = null

	for (const part of referenceParts(formula)) {
		const range = readReference(part)

		if (range === null) continue
		if (range.startCol === 0 && range.endCol === SHEET_MAX_COLS - 1) rows ??= { first: range.startRow, last: range.endRow }
		else if (range.startRow === 0 && range.endRow === SHEET_MAX_ROWS - 1) cols ??= { first: range.startCol, last: range.endCol }
	}

	return { rows, cols }
}
