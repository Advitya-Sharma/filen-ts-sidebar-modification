import type { SliceClock } from "../../slicing"
import type { SharedStrings } from "../../xlsx/model/sharedStrings"
import { SHEET_MAX_COLS } from "../../sheet/protocol"
import { CellKind } from "../../xlsx/model/types"
import { translateFormula, volatileFormula } from "../formula/toXlsx"
import type { OdsIssues } from "../issues"
import type { OdsCellStyles } from "../styles/cell"
import { OdsReadError } from "./errors"
import { A, E, OdsScanner, UNKNOWN, type ScanHandler } from "./scan"
import { OdsSheetBuilder, type OdsBuiltSheet, type OdsCellInput, type OdsRowInput } from "./sheet"
import { capturedElement } from "./captured"
import { dateValue, durationValue, nullDate, NULL_DATE_1899 } from "./values"
import type { OdfElement } from "../../odf/tree"

// office:body of content.xml read in one streamed pass (ODF 1.3 Part 3 §9.1 table:table and its children), into one
// built sheet per table plus the spreadsheet's own elements, which are small and kept as trees.

export interface OdsTableProtection {
	readonly key: string | undefined
	readonly algorithm: string | undefined
	// loext:table-protection / table:table-protection attributes (what stays allowed).
	readonly allows: ReadonlyMap<string, boolean>
}

export interface OdsTable {
	readonly name: string
	readonly styleName: string | undefined
	readonly protection: OdsTableProtection | null
	readonly printRanges: string | undefined
	readonly printed: boolean
	readonly built: OdsBuiltSheet
	readonly conditionalFormats: readonly OdfElement[]
	readonly names: readonly OdfElement[]
}

export interface OdsContent {
	readonly tables: readonly OdsTable[]
	// Days from 1970-01-01 to the null date (§19.639 table:null-date).
	readonly nullDate: number
	readonly calculation: OdfElement | undefined
	readonly validations: OdfElement | undefined
	readonly names: readonly OdfElement[]
	readonly databaseRanges: OdfElement | undefined
	readonly structureProtected: boolean
	readonly protectionKey: string | undefined
	readonly protectionAlgorithm: string | undefined
	readonly cells: number
}

const WHITE_SPACE = /[ \t\r\n]+/
// Whether text holds what §6.1.2 collapses: a tab or line end, two spaces, or a space at either end. A loop, not a
// regular expression: it runs once per text, and the test call costs more than the scan.
function collapses(text: string): boolean {
	const last = text.length - 1

	if (last < 0) return false
	if (text.charCodeAt(0) === 32 || text.charCodeAt(last) === 32) return true

	for (let i = 0; i <= last; i++) {
		const code = text.charCodeAt(i)

		if (code === 9 || code === 10 || code === 13 || (code === 32 && text.charCodeAt(i + 1) === 32)) return true
	}

	return false
}

const VALUE_TYPES = ["float", "percentage", "currency", "string", "date", "time", "boolean", "void"]
// An error result as LibreOffice writes its text: "#NAME?", "#VALUE!" (in the document's language), "Err:502".
const ERROR_TEXT = /^(?:Err:\d{3}|#[^\s#]{1,24})$/

// Value types (office:value-type, §19.389), with an error result (calcext:value-type="error").
const NONE = 0
const NUMBER = 1
const DATE = 2
const TIME = 3
const BOOLEAN = 4
const STRING = 5
const ERROR = 6

// What the reader is inside: the element stack's interesting frames.
const Where = {
	Body: 0,
	Table: 1,
	Row: 2,
	Cell: 3,
	Paragraph: 4
} as const

type Where = (typeof Where)[keyof typeof Where]

interface TableState {
	readonly builder: OdsSheetBuilder
	readonly name: string
	readonly styleName: string | undefined
	readonly protection: { key: string | undefined; algorithm: string | undefined; allows: Map<string, boolean> } | null
	readonly printRanges: string | undefined
	readonly printed: boolean
	readonly conditionalFormats: OdfElement[]
	readonly names: OdfElement[]
	rowOutline: number
	colOutline: number
	rowsCollapsed: number
	colsCollapsed: number
	headerColumns: boolean
}

export class ContentReader implements ScanHandler {
	readonly scanner: OdsScanner
	private readonly styles: OdsCellStyles
	private readonly issues: OdsIssues
	private readonly tables: OdsTable[] = []
	private table: TableState | null = null
	private readonly names: OdfElement[] = []
	private calculation: OdfElement | undefined
	private validations: OdfElement | undefined
	private databaseRanges: OdfElement | undefined
	private nullDateDays = NULL_DATE_1899
	private structureProtected = false
	private protectionKey: string | undefined
	private protectionAlgorithm: string | undefined
	private inSpreadsheet = false
	// Elements opened since the spreadsheet, with what each stands for: the reader acts on the innermost.
	private readonly stack: number[] = []
	private where: Where = Where.Body
	// Inside an element whose children the reader passes over (and how deep).
	private ignoring = 0
	// Closes already handled: of empty elements and of captured ones.
	private skipClose = 0
	// The row and cell under way.
	private readonly row: OdsRowInput = {
		repeat: 1,
		style: undefined,
		defaultCellStyle: undefined,
		hidden: false,
		outline: 0,
		collapsed: false
	}
	private readonly cell: OdsCellInput = {
		col: 0,
		repeat: 1,
		covered: false,
		style: undefined,
		kind: CellKind.EMPTY,
		number: 0,
		text: "",
		formula: "",
		colSpan: 1,
		rowSpan: 1,
		matrixCols: 0,
		matrixRows: 0,
		validation: undefined,
		link: undefined,
		note: undefined,
		rich: false,
		runs: undefined,
		recalculate: false
	}
	private valueType = NONE
	private stringValue: string | undefined
	private wantsText = false
	private paragraphs = 0
	private cellText = ""
	private hasTextContent = false
	private pendingSpace = false
	private hasFormula = false
	// The cell is what LibreOffice reads as an edit cell (paragraphs, spans, fields, line breaks): its text shows as
	// written, never through a number format's text section. Spans' text styles, and where each run starts.
	private editCell = false
	private readonly spans: (string | undefined)[] = []
	// The style of the run under way; null before the cell's first text.
	private runStyle: string | undefined | null = null
	// The cell's runs, the first runCount entries (the arrays are reused, never shrunk: a cell's runs cost no allocation).
	private readonly runStarts: number[] = []
	private readonly runStyles: (string | undefined)[] = []
	private runCount = 0
	private paragraphStart = true
	private cellCount = 0

	private readonly sharedStrings: SharedStrings
	// Cells every sheet may still materialise together (the workbook's cell limit).
	private readonly budget: { left: number }
	private readonly maxSheets: number

	constructor(
		styles: OdsCellStyles,
		issues: OdsIssues,
		sharedStrings: SharedStrings,
		prefixes: ReadonlyMap<string, string>,
		bodyStart: number,
		maxCells: number,
		maxSheets: number
	) {
		this.budget = { left: maxCells }
		this.maxSheets = maxSheets
		this.styles = styles
		this.issues = issues
		this.sharedStrings = sharedStrings
		this.scanner = new OdsScanner(this, prefixes, bodyStart)
	}

	// Feeds the stream through, giving the host a turn whenever the clock's slice is used.
	async read(
		chunks: AsyncIterable<Uint8Array>,
		clock: SliceClock | undefined,
		signal: { readonly aborted: boolean } | undefined
	): Promise<OdsContent> {
		for await (const chunk of chunks) {
			this.scanner.push(chunk)

			if (clock?.due === true) await clock.pause()
			if (signal?.aborted === true) throw new OdsReadError("aborted", "Reading was aborted")
		}

		this.scanner.end()

		return {
			tables: this.tables,
			nullDate: this.nullDateDays,
			calculation: this.calculation,
			validations: this.validations,
			names: this.names,
			databaseRanges: this.databaseRanges,
			structureProtected: this.structureProtected,
			protectionKey: this.protectionKey,
			protectionAlgorithm: this.protectionAlgorithm,
			cells: this.cellCount
		}
	}

	open(element: number, selfClosing: boolean): void {
		// The scanner closes an empty element at once: that close was handled here.
		if (selfClosing) this.skipClose++

		if (this.ignoring > 0) {
			if (!selfClosing) this.ignoring++

			return
		}

		if (!this.inSpreadsheet) {
			if (element === E["office:spreadsheet"]) {
				this.inSpreadsheet = true
				this.structureProtected = this.scanner.is(A["table:structure-protected"], "true")
				this.protectionKey = this.scanner.attribute(A["table:protection-key"])
				this.protectionAlgorithm = this.scanner.attribute(A["table:protection-key-digest-algorithm"])
			}

			return
		}

		switch (this.where) {
			case Where.Body:
				this.openInBody(element, selfClosing)
				break
			case Where.Table:
				this.openInTable(element, selfClosing)
				break
			case Where.Row:
				if (element === E["table:table-cell"] || element === E["table:covered-table-cell"])
					this.openCell(element === E["table:covered-table-cell"], selfClosing)
				else this.ignore(selfClosing)
				break
			case Where.Cell:
				this.openInCell(element, selfClosing)
				break
			case Where.Paragraph:
				this.openInParagraph(element, selfClosing)
				break
		}
	}

	close(element: number): void {
		if (this.skipClose > 0) {
			this.skipClose--

			return
		}

		if (this.ignoring > 0) {
			this.ignoring--

			return
		}

		if (!this.inSpreadsheet) return

		if (element === E["office:spreadsheet"] && this.where === Where.Body) {
			this.inSpreadsheet = false

			return
		}

		const frame = this.stack.pop()

		if (frame === undefined) return

		switch (frame) {
			case E["table:table"]:
				this.endTable()
				break
			case E["table:table-row"]:
				this.endRow()
				break
			case E["table:table-cell"]:
			case E["table:covered-table-cell"]:
				this.endCell()
				break
			case E["text:p"]:
			case E["text:h"]:
				// White space at the paragraph's end goes, as at its start (step 5).
				this.pendingSpace = false
				this.where = Where.Cell
				this.scanner.wantText = false
				break
			case E["text:span"]:
				this.spans.pop()
				break
			case E["table:table-row-group"]: {
				const table = this.table

				if (table !== null) table.rowOutline--
				break
			}
			case E["table:table-column-group"]: {
				const table = this.table

				if (table !== null) table.colOutline--
				break
			}
			case E["table:table-header-rows"]:
				this.table?.builder.headerRowsEnd()
				break
			case E["table:table-header-columns"]: {
				const table = this.table

				if (table !== null) table.headerColumns = false
				break
			}
		}
	}

	// Paragraph text with its white space collapsed (ODF 1.3 Part 3 §6.1.2): a run of spaces, tabs and line ends is one
	// space, and none at the paragraph's start or end; text:s, text:tab and text:line-break stand for themselves.
	// LibreOffice 7.5 keeps a paragraph's raw white space; the specification's algorithm is followed.
	text(text: string): void {
		if (this.ignoring > 0) return

		// Most text has no run of white space and none at its ends: as it is, without splitting.
		if (!collapses(text)) {
			if (text !== "") this.emit(text)

			return
		}

		const parts = text.split(WHITE_SPACE)

		for (let k = 0; k < parts.length; k++) {
			if (k > 0) this.pendingSpace = true

			const part = parts[k] ?? ""

			if (part !== "") this.emit(part)
		}
	}

	private emit(text: string): void {
		if (this.pendingSpace && !this.paragraphStart) this.cellText += " "

		this.pendingSpace = false
		this.paragraphStart = false

		const style = this.spans[this.spans.length - 1]

		if (style !== this.runStyle) {
			this.runStarts[this.runCount] = this.cellText.length
			this.runStyles[this.runCount] = style
			this.runCount++
			this.runStyle = style
		}

		this.cellText += text
		this.hasTextContent = true
	}

	captured(element: number, markup: Uint8Array): void {
		const prefixes = this.scanner.namespaces

		switch (element) {
			case E["table:calculation-settings"]: {
				const settings = capturedElement(markup, prefixes)

				this.calculation = settings

				const date = settings.children.find(child => child.name === "table:null-date")?.attributes.get("table:date-value")

				this.nullDateDays = nullDate(date)
				break
			}
			case E["table:content-validations"]:
				this.validations = capturedElement(markup, prefixes)
				break
			case E["table:named-expressions"]:
				if (this.table === null) this.names.push(capturedElement(markup, prefixes))
				else this.table.names.push(capturedElement(markup, prefixes))
				break
			case E["table:database-ranges"]:
				this.databaseRanges = capturedElement(markup, prefixes)
				break
			case E["calcext:conditional-formats"]:
				this.table?.conditionalFormats.push(capturedElement(markup, prefixes))
				break
			case E["office:annotation"]:
				if (this.cell.note === undefined) this.cell.note = markup
				break
			case E["table:shapes"]:
				this.issues.add("drawing", this.table?.name, "shapes", Math.max(1, countShapes(markup)))
				break
			case E["office:forms"]:
				if (hasChildElements(markup)) this.issues.add("form", this.table?.name)
				break
			default:
				break
		}
	}

	private ignore(selfClosing: boolean): void {
		if (!selfClosing) this.ignoring = 1
	}

	// The element's markup goes to captured(), then its close (handled there).
	private capture(selfClosing: boolean): void {
		this.scanner.captureElementMarkup()

		if (!selfClosing) this.skipClose++
	}

	private push(element: number, selfClosing: boolean): void {
		if (!selfClosing) this.stack.push(element)
	}

	private openInBody(element: number, selfClosing: boolean): void {
		switch (element) {
			case E["table:table"]:
				this.openTable(selfClosing)
				break
			case E["table:calculation-settings"]:
			case E["table:content-validations"]:
			case E["table:named-expressions"]:
			case E["table:database-ranges"]:
				this.capture(selfClosing)
				break
			case E["table:data-pilot-tables"]:
				this.issues.add("pivotTable", undefined)
				this.ignore(selfClosing)
				break
			case E["table:consolidation"]:
				this.issues.add("consolidation", undefined)
				this.ignore(selfClosing)
				break
			case E["table:dde-links"]:
				this.issues.add("ddeLink", undefined)
				this.ignore(selfClosing)
				break
			case E["table:label-ranges"]:
				if (!selfClosing) this.issues.add("labelRanges", undefined)
				this.ignore(selfClosing)
				break
			case E["table:tracked-changes"]:
				this.issues.add("trackedChanges", undefined)
				this.ignore(selfClosing)
				break
			default:
				this.ignore(selfClosing)
		}
	}

	private openTable(selfClosing: boolean): void {
		const scanner = this.scanner

		if (this.tables.length >= this.maxSheets) throw new OdsReadError("tooLarge", "Too many sheets")
		const name = scanner.attribute(A["table:name"]) ?? `Sheet${this.tables.length + 1}`
		const isProtected = scanner.is(A["table:protected"], "true")

		this.table = {
			builder: new OdsSheetBuilder(name, this.styles, this.issues, this.sharedStrings, this.budget),
			name,
			styleName: scanner.attribute(A["table:style-name"]),
			protection: isProtected
				? {
						key: scanner.attribute(A["table:protection-key"]),
						algorithm: scanner.attribute(A["table:protection-key-digest-algorithm"]),
						allows: new Map()
					}
				: null,
			printRanges: scanner.attribute(A["table:print-ranges"]),
			printed: !scanner.is(A["table:print"], "false"),
			conditionalFormats: [],
			names: [],
			rowOutline: 0,
			colOutline: 0,
			rowsCollapsed: 0,
			colsCollapsed: 0,
			headerColumns: false
		}

		if (scanner.is(A["table:is-sub-table"], "true")) this.issues.add("subTable", name)

		if (selfClosing) {
			this.endTable()

			return
		}

		this.stack.push(E["table:table"])
		this.where = Where.Table
	}

	private openInTable(element: number, selfClosing: boolean): void {
		const table = this.table

		if (table === null) return this.ignore(selfClosing)

		const scanner = this.scanner

		switch (element) {
			case E["table:table-column"]: {
				const hidden = !scanner.is(A["table:visibility"], "visible") && scanner.has(A["table:visibility"])

				table.builder.column(
					scanner.count(A["table:number-columns-repeated"], 1),
					scanner.attribute(A["table:style-name"]),
					scanner.attribute(A["table:default-cell-style-name"]),
					hidden,
					table.colOutline,
					table.colsCollapsed > 0,
					table.headerColumns
				)

				if (!selfClosing) this.ignore(false)
				break
			}
			case E["table:table-column-group"]:
				if (!selfClosing) {
					table.colOutline++
					this.stack.push(element)
				}
				break
			case E["table:table-header-columns"]:
				if (!selfClosing) {
					table.headerColumns = true
					this.stack.push(element)
				}
				break
			case E["table:table-columns"]:
			case E["table:table-rows"]:
				this.push(element, selfClosing)
				break
			case E["table:table-row-group"]:
				if (!selfClosing) {
					table.rowOutline++
					this.stack.push(element)
				}
				break
			case E["table:table-header-rows"]:
				if (!selfClosing) {
					table.builder.headerRowsStart()
					this.stack.push(element)
				}
				break
			case E["table:table-row"]:
				this.openRow(selfClosing)
				break
			case E["calcext:conditional-formats"]:
			case E["table:named-expressions"]:
			case E["table:shapes"]:
			case E["office:forms"]:
				this.capture(selfClosing)
				break
			case E["loext:table-protection"]:
			case E["table:table-protection"]:
				if (table.protection !== null) {
					for (const [id, name] of PROTECTION_ATTRIBUTES) {
						const value = scanner.attribute(id)

						if (value !== undefined) table.protection.allows.set(name, value === "true")
					}
				}
				this.ignore(selfClosing)
				break
			case E["table:table-source"]:
				this.issues.add("linkedSheet", table.name)
				this.ignore(selfClosing)
				break
			case E["table:scenario"]:
				this.issues.add("scenario", table.name)
				this.ignore(selfClosing)
				break
			default:
				this.ignore(selfClosing)
		}
	}

	private openRow(selfClosing: boolean): void {
		const scanner = this.scanner
		const row = this.row
		const visibility = scanner.symbol(A["table:visibility"])

		row.repeat = scanner.count(A["table:number-rows-repeated"], 1)
		row.style = scanner.symbol(A["table:style-name"])
		row.defaultCellStyle = scanner.symbol(A["table:default-cell-style-name"])
		row.hidden = visibility !== undefined && visibility !== "visible"
		row.outline = this.table?.rowOutline ?? 0
		row.collapsed = false
		this.table?.builder.beginRow()

		if (selfClosing) {
			this.endRow()

			return
		}

		this.stack.push(E["table:table-row"])
		this.where = Where.Row
	}

	private endRow(): void {
		this.table?.builder.endRow(this.row)
		this.where = Where.Table
	}

	private openCell(covered: boolean, selfClosing: boolean): void {
		const scanner = this.scanner
		const cell = this.cell

		cell.repeat = scanner.count(A["table:number-columns-repeated"], 1)
		cell.covered = covered
		cell.style = scanner.symbol(A["table:style-name"]) ?? this.row.defaultCellStyle
		cell.colSpan = covered ? 1 : scanner.count(A["table:number-columns-spanned"], 1)
		cell.rowSpan = covered ? 1 : scanner.count(A["table:number-rows-spanned"], 1)
		cell.matrixCols = scanner.count(A["table:number-matrix-columns-spanned"], 0)
		cell.matrixRows = scanner.count(A["table:number-matrix-rows-spanned"], 0)
		cell.validation = scanner.symbol(A["table:content-validation-name"])
		cell.link = undefined
		cell.note = undefined
		cell.kind = CellKind.EMPTY
		cell.number = 0
		cell.text = ""
		cell.formula = ""
		cell.recalculate = false
		this.paragraphs = 0
		this.cellText = ""
		this.hasTextContent = false
		this.stringValue = undefined
		this.editCell = false
		// Emptied only when used: setting an array's length is a slow path in SpiderMonkey, once per cell.
		if (this.spans.length > 0) this.spans.length = 0
		this.runStyle = null
		this.runCount = 0
		cell.rich = false
		cell.runs = undefined

		let valueType = NONE

		if (scanner.is(A["calcext:value-type"], "error")) {
			valueType = ERROR
		} else {
			switch (scanner.match(A["office:value-type"], VALUE_TYPES)) {
				case 0:
				case 1:
				case 2:
					valueType = NUMBER
					cell.number = scanner.number(A["office:value"])
					break
				case 3:
					valueType = STRING
					this.stringValue = scanner.attribute(A["office:string-value"])
					break
				case 4:
					valueType = DATE
					cell.number = dateValue(scanner.symbol(A["office:date-value"]) ?? "") - this.nullDateDays
					break
				case 5:
					valueType = TIME
					cell.number = durationValue(scanner.attribute(A["office:time-value"]) ?? "")
					break
				case 6:
					valueType = BOOLEAN
					cell.number = scanner.is(A["office:boolean-value"], "true") ? 1 : 0
					break
				case 7:
					break
				default:
					// Another type is text (LibreOffice reads it so).
					if (scanner.has(A["office:value-type"])) valueType = STRING
			}
		}

		this.valueType = valueType

		const formula = scanner.attribute(A["table:formula"])

		if (formula !== undefined) {
			const translated = translateFormula(formula)

			if ("text" in translated) {
				cell.formula = translated.text
				cell.recalculate = volatileFormula(translated.text)
			} else {
				this.issues.add("formula", this.table?.name, translated.refused)
			}
		}

		// The text is read where the value is text (or there is no value); a number's is the display LibreOffice wrote,
		// read for a formula only: an error result written as a number with its error text (before calcext:value-type).
		this.hasFormula = formula !== undefined
		this.wantsText =
			valueType === STRING || valueType === ERROR || valueType === NONE || (formula !== undefined && valueType === NUMBER)

		if (selfClosing) {
			this.endCell()

			return
		}

		this.stack.push(covered ? E["table:covered-table-cell"] : E["table:table-cell"])
		this.where = Where.Cell
		this.scanner.expectParagraph()
	}

	// A cell holding one plain paragraph, read at once by the scanner.
	paragraph(): void {
		this.paragraphs++

		if (!this.wantsText) return

		this.paragraphStart = true
		this.pendingSpace = false
		this.text(this.scanner.paragraphText())
		this.pendingSpace = false
	}

	private openInCell(element: number, selfClosing: boolean): void {
		switch (element) {
			case E["text:p"]:
			case E["text:h"]:
				if (this.paragraphs++ > 0) {
					this.cellText += "\n"
					this.editCell = true
				}

				this.paragraphStart = true
				this.pendingSpace = false

				if (!selfClosing) {
					this.stack.push(element)
					this.where = Where.Paragraph
					this.scanner.wantText = this.wantsText
				}
				break
			case E["office:annotation"]:
				this.capture(selfClosing)
				break
			case E["table:detective"]:
				this.issues.add("detective", this.table?.name)
				this.ignore(selfClosing)
				break
			case E["table:cell-range-source"]:
				this.issues.add("cellRangeSource", this.table?.name)
				this.ignore(selfClosing)
				break
			case E["office:annotation-end"]:
				this.ignore(selfClosing)
				break
			case E["table:table"]:
				this.issues.add("subTable", this.table?.name)
				this.ignore(selfClosing)
				break
			default:
				// draw:frame, draw:custom-shape and other shapes anchored to the cell.
				if (element === UNKNOWN) this.issues.add("drawing", this.table?.name, "cell")
				this.ignore(selfClosing)
		}
	}

	private openInParagraph(element: number, selfClosing: boolean): void {
		switch (element) {
			case E["text:s"]:
				if (this.wantsText) this.emit(" ".repeat(Math.min(SHEET_MAX_COLS, this.scanner.count(A["text:c"], 1))))
				break
			case E["text:tab"]:
				if (this.wantsText) this.emit("\t")
				break
			case E["text:line-break"]:
				if (this.wantsText) this.emit("\n")
				this.editCell = true
				break
			case E["text:span"]:
				this.editCell = true
				if (!selfClosing) this.spans.push(this.scanner.attribute(A["text:style-name"]))
				break
			case E["text:a"]:
				if (this.cell.link === undefined) this.cell.link = this.scanner.attribute(A["xlink:href"])
				else this.issues.add("link", this.table?.name)

				this.editCell = true
				break
			case E["office:annotation"]:
				this.capture(selfClosing)

				return
			case E["text:note"]:
			case E["text:ruby-text"]:
				this.ignore(selfClosing)

				return
			default:
				// A field (a date, a sheet name, a page number...): the cell is an edit cell showing its text.
				this.editCell = true
				break
		}

		// Every other inline element (spans, links, fields) is gone through: its text is the paragraph's.
		if (!selfClosing) this.stack.push(element)
	}

	private endCell(): void {
		const cell = this.cell
		const table = this.table

		this.where = Where.Row
		this.scanner.wantText = false

		if (table === null) return

		const valueType = this.valueType

		switch (valueType) {
			case NUMBER:
				// An error result written as a number with its error text, in the writer's language (older LibreOffice):
				// shown until the engine calculates it, which gives the error in its own words.
				if (this.hasFormula && ERROR_TEXT.test(this.cellText)) {
					cell.kind = CellKind.ERROR
					cell.text = this.cellText
					cell.recalculate = true
				} else {
					cell.kind = Number.isFinite(cell.number) ? CellKind.NUMBER : CellKind.EMPTY
				}
				break
			case DATE:
			case TIME:
				cell.kind = Number.isFinite(cell.number) ? CellKind.NUMBER : CellKind.EMPTY
				break
			case BOOLEAN:
				cell.kind = CellKind.BOOLEAN
				break
			case STRING:
				cell.kind = CellKind.INLINE_STRING
				cell.text = this.stringValue ?? this.cellText
				this.richText(cell)
				break
			case ERROR:
				cell.kind = CellKind.ERROR
				cell.text = this.cellText
				break
			default:
				if (this.paragraphs > 0 && (this.hasTextContent || this.cellText !== "")) {
					cell.kind = CellKind.INLINE_STRING
					cell.text = this.cellText
					this.richText(cell)
				}
		}

		const hasContent =
			cell.kind !== CellKind.EMPTY ||
			cell.formula !== "" ||
			cell.note !== undefined ||
			cell.link !== undefined ||
			cell.colSpan > 1 ||
			cell.rowSpan > 1

		table.builder.cell(cell, hasContent, this.row)
		this.cellCount += cell.repeat
	}

	// An edit cell's text, with its spans' fonts as runs where a span has a text style; a formula's result never is one.
	private richText(cell: OdsCellInput): void {
		if (!this.editCell || this.hasFormula || this.stringValue !== undefined) return

		cell.rich = true

		let styled = false

		for (let k = 0; k < this.runCount && !styled; k++) styled = this.runStyles[k] !== undefined

		if (!styled) return

		const runs: { start: number; style: string | undefined }[] = []

		for (let k = 0; k < this.runCount; k++) runs.push({ start: this.runStarts[k] ?? 0, style: this.runStyles[k] })

		cell.runs = runs
	}

	private endTable(): void {
		const table = this.table

		this.where = Where.Body

		if (table === null) return

		this.tables.push({
			name: table.name,
			styleName: table.styleName,
			protection: table.protection,
			printRanges: table.printRanges,
			printed: table.printed,
			built: table.builder.finish(),
			conditionalFormats: table.conditionalFormats,
			names: table.names
		})
		this.table = null
	}
}

const PROTECTION_ATTRIBUTES: readonly (readonly [number, string])[] = [
	[A["loext:select-protected-cells"], "selectLockedCells"],
	[A["loext:select-unprotected-cells"], "selectUnlockedCells"],
	[A["table:select-protected-cells"], "selectLockedCells"],
	[A["table:select-unprotected-cells"], "selectUnlockedCells"],
	[A["loext:insert-columns"], "insertColumns"],
	[A["loext:insert-rows"], "insertRows"],
	[A["loext:delete-columns"], "deleteColumns"],
	[A["loext:delete-rows"], "deleteRows"]
]

// Elements directly inside a captured element (beyond its own start tag): table:shapes' shapes, office:forms' forms.
function countShapes(markup: Uint8Array): number {
	let depth = 0
	let count = 0

	for (let at = markup.indexOf(60); at >= 0; at = markup.indexOf(60, at + 1)) {
		const next = markup[at + 1]

		if (next === 47) {
			depth--
		} else if (next !== 33 && next !== 63) {
			const close = markup.indexOf(62, at)

			if (depth === 1) count++
			if (close > 0 && markup[close - 1] !== 47) depth++
		}
	}

	return count
}

function hasChildElements(markup: Uint8Array): boolean {
	return countShapes(markup) > 0
}
