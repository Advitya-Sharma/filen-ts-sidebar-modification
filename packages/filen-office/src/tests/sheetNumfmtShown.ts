import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { bytesSource, OpcPackage, type Relationships } from "@filen/office/core/opc"
import { Namespaces, type XmlDocument } from "@filen/office/core/xml"
import { formatNumberInto, formatTextInto, NumFmtOutput, SectionKind, type NumFmt } from "../core/sheet/numfmt"
import { decompose } from "../core/sheet/numfmt/decimal"
import { SheetStyles, type SheetStyleSource } from "../core/sheet/style"
import type { XlsxColor, XlsxFill, XlsxTheme, XlsxXf } from "../core/xlsx/model/types"
import type { CellScalar } from "./oracles/sheet/xlsxCells"
import type { ShownWorkbook } from "./oracles/sheet/libreOffice"
import { ZIP_ENTRY_TIME } from "./zipTime"

// The number-format and fill oracles' side of the corpus: a plain reader of what each cell holds (its value as
// cached and its cell format), the styles the formatter and the style resolver need (numFmts, cellXfs and
// cellStyleXfs, fills, the palette, the theme's colours), and the comparison with LibreOffice's "as shown"
// exports, every difference classified.

export interface ShownCell {
	readonly row: number
	readonly column: number
	readonly style: number
	// undefined for a cell kept only for its format.
	readonly value: CellScalar | undefined
	// A shared string with runs of its own format: shown as it is, without the format's text section (LibreOffice
	// keeps rich text apart from number formats; the runs index into the text as stored).
	readonly rich: boolean
}

export interface ShownInput {
	readonly date1904: boolean
	readonly styles: SheetStyleSource
	readonly theme: XlsxTheme | null
	readonly sheets: readonly ShownSheetInput[]
}

export interface ShownSheetInput {
	readonly name: string
	readonly cells: readonly ShownCell[]
	// Ranges whose fill is not their cells' own ([top, left, bottom, right], 0-based): under conditional formatting,
	// merged into another cell, or under a drawing (LibreOffice's HTML export gives such a cell the shape's fill).
	readonly conditional: readonly (readonly [number, number, number, number])[]
	// The first hidden row and column (-1 for none): LibreOffice's HTML export leaves them out, so its cells are
	// aligned only before them.
	readonly firstHiddenRow: number
	readonly firstHiddenColumn: number
}

function* children(document: XmlDocument, node: number, local?: string): Generator<number> {
	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		if (local === undefined || document.localName(child) === local) yield child
}

function child(document: XmlDocument, node: number, local: string): number {
	for (const found of children(document, node, local)) return found

	return -1
}

function relationshipOf(relationships: Relationships, suffix: string): string | undefined {
	return relationships.all.find(relationship => relationship.type.endsWith(suffix) && relationship.partName !== undefined)?.partName
}

// A string's text: its runs' <t> (phonetic runs left out), _xHHHH_ escapes decoded (§22.9.2.19 ST_Xstring).
function stringText(document: XmlDocument, node: number): string {
	let out = ""

	for (const item of children(document, node)) {
		const local = document.localName(item)

		if (local === "t") out += document.textContent(item)
		else if (local === "r") for (const t of children(document, item, "t")) out += document.textContent(t)
	}

	return unescapeXstring(out)
}

function unescapeXstring(text: string): string {
	return text.replace(/_x([0-9a-fA-F]{4})_/gu, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
}

function number(document: XmlDocument, node: number, name: string): number | undefined {
	const value = document.attribute(node, name)

	return value === undefined || value === "" ? undefined : Number(value)
}

function truthy(value: string | undefined): boolean {
	return value === "1" || value === "true"
}

function color(document: XmlDocument, node: number): XlsxColor | undefined {
	if (node < 0) return undefined

	const rgb = document.attribute(node, "rgb")
	const theme = number(document, node, "theme")
	const indexed = number(document, node, "indexed")
	const tint = number(document, node, "tint")

	return {
		...(rgb !== undefined && /^[0-9a-fA-F]{6,8}$/u.test(rgb) && { argb: parseInt(rgb, 16) }),
		...(theme !== undefined && { theme }),
		...(indexed !== undefined && { indexed }),
		...(tint !== undefined && { tint }),
		...(truthy(document.attribute(node, "auto")) && { auto: true })
	}
}

function fill(document: XmlDocument, node: number): XlsxFill {
	const pattern = child(document, node, "patternFill")

	if (pattern >= 0) {
		const type = (document.attribute(pattern, "patternType") ?? "none") as Extract<XlsxFill, { kind: "pattern" }>["pattern"]
		const fg = color(document, child(document, pattern, "fgColor"))
		const bg = color(document, child(document, pattern, "bgColor"))

		return { kind: "pattern", pattern: type, ...(fg !== undefined && { fg }), ...(bg !== undefined && { bg }) }
	}

	const gradient = child(document, node, "gradientFill")

	if (gradient < 0) return { kind: "pattern", pattern: "none" }

	return {
		kind: "gradient",
		type: document.attribute(gradient, "type") === "path" ? "path" : "linear",
		degree: number(document, gradient, "degree") ?? 0,
		left: number(document, gradient, "left") ?? 0,
		right: number(document, gradient, "right") ?? 0,
		top: number(document, gradient, "top") ?? 0,
		bottom: number(document, gradient, "bottom") ?? 0,
		stops: [...children(document, gradient, "stop")].map(stop => ({
			position: number(document, stop, "position") ?? 0,
			color: color(document, child(document, stop, "color")) ?? {}
		}))
	}
}

function xfs(document: XmlDocument, node: number, cellXf: boolean): XlsxXf[] {
	if (node < 0) return []

	return [...children(document, node, "xf")].map(xf => ({
		numFmtId: number(document, xf, "numFmtId") ?? 0,
		fontId: number(document, xf, "fontId") ?? 0,
		fillId: number(document, xf, "fillId") ?? 0,
		borderId: number(document, xf, "borderId") ?? 0,
		xfId: cellXf ? (number(document, xf, "xfId") ?? 0) : -1,
		flags: truthy(document.attribute(xf, "applyFill")) ? 4 : 0
	}))
}

function readStyles(document: XmlDocument): SheetStyleSource {
	const top = document.documentElement()
	const numFmts = new Map<number, string>()
	const numFmtsNode = child(document, top, "numFmts")

	for (const item of numFmtsNode < 0 ? [] : children(document, numFmtsNode, "numFmt")) {
		const id = number(document, item, "numFmtId")
		const code = document.attribute(item, "formatCode")

		if (id !== undefined && code !== undefined) numFmts.set(id, code)
	}

	const fillsNode = child(document, top, "fills")
	const colorsNode = child(document, top, "colors")
	const indexedNode = colorsNode < 0 ? -1 : child(document, colorsNode, "indexedColors")
	const indexedColors =
		indexedNode < 0
			? null
			: [...children(document, indexedNode, "rgbColor")].map(item => parseInt(document.attribute(item, "rgb") ?? "0", 16))

	return {
		numFmts,
		fonts: [],
		fills: fillsNode < 0 ? [] : [...children(document, fillsNode, "fill")].map(item => fill(document, item)),
		borders: [],
		cellStyleXfs: xfs(document, child(document, top, "cellStyleXfs"), false),
		cellXfs: xfs(document, child(document, top, "cellXfs"), true),
		indexedColors
	}
}

function readTheme(document: XmlDocument): XlsxTheme | null {
	let scheme = -1
	const find = (node: number): void => {
		for (const item of children(document, node)) {
			if (scheme >= 0) return
			if (document.localName(item) === "clrScheme") scheme = item
			else find(item)
		}
	}

	find(document.documentElement())

	if (scheme < 0) return null

	const colors = [...children(document, scheme)].map(item => {
		const value = document.firstChildElement(item)

		if (value < 0) return 0

		const hex = document.localName(value) === "sysClr" ? document.attribute(value, "lastClr") : document.attribute(value, "val")

		return parseInt(hex ?? "0", 16)
	})

	return { colors, majorFont: "", minorFont: "" }
}

function cellValue(type: string, raw: string | undefined, strings: readonly string[], inline: string | undefined): CellScalar | undefined {
	switch (type) {
		case "s": {
			const value = raw === undefined || raw === "" ? undefined : strings[Number(raw)]

			return value === undefined ? undefined : { type: "s", value }
		}
		case "str":
			return raw === undefined ? undefined : { type: "s", value: unescapeXstring(raw) }
		case "inlineStr":
			return inline === undefined ? undefined : { type: "s", value: inline }
		case "b":
			return raw === undefined || raw === "" ? undefined : { type: "b", value: raw.trim() === "1" || raw.trim() === "true" }
		case "e":
			return raw === undefined || raw === "" ? undefined : { type: "e", value: raw.trim() }
		case "d":
			return undefined
		default: {
			if (raw === undefined || raw.trim() === "") return undefined

			const value = Number(raw)

			return Number.isNaN(value) ? undefined : { type: "n", value }
		}
	}
}

function column(ref: string): number {
	let index = 0

	for (let i = 0; i < ref.length; i++) {
		const c = ref.charCodeAt(i)

		if (c < 65 || c > 90) break

		index = index * 26 + (c - 64)
	}

	return index - 1
}

function range(ref: string): [number, number, number, number] {
	const [first = "", last = first] = ref.toUpperCase().replaceAll("$", "").split(":")
	const at = (a1: string): [number, number] => [Number(/\d+/u.exec(a1)?.[0] ?? "1") - 1, column(a1)]
	const [top, left] = at(first)
	const [bottom, right] = at(last)

	return [top, left, bottom, right]
}

// Conditional formatting ranges (the sheet's own and the x14 ones in extLst), and the merged cells but their
// top-left one (a merged range shows its first cell's format).
function readCovered(document: XmlDocument): [number, number, number, number][] {
	const ranges: [number, number, number, number][] = []
	const visit = (node: number): void => {
		for (const item of children(document, node)) {
			const local = document.localName(item)

			if (local === "conditionalFormatting") {
				const inner = child(document, item, "sqref")
				const sqref = document.attribute(item, "sqref") ?? (inner < 0 ? "" : document.textContent(inner))

				for (const ref of sqref.split(" ")) if (ref !== "") ranges.push(range(ref))
			} else if (local === "mergeCell") {
				const [top, left, bottom, right] = range(document.attribute(item, "ref") ?? "A1")

				if (bottom > top) ranges.push([top + 1, left, bottom, right])
				if (right > left) ranges.push([top, left + 1, top, right])
			} else if (local === "mergeCells" || local === "extLst" || local === "ext" || local === "conditionalFormattings") visit(item)
		}
	}

	visit(document.documentElement())

	return ranges
}

// The cells a drawing's anchors span (xdr:from and xdr:to of two-cell anchors, the from cell of the others).
function readAnchors(document: XmlDocument): [number, number, number, number][] {
	const ranges: [number, number, number, number][] = []
	const cell = (node: number): [number, number] | undefined =>
		node < 0
			? undefined
			: [Number(document.textContent(child(document, node, "row"))), Number(document.textContent(child(document, node, "col")))]

	for (const anchor of children(document, document.documentElement())) {
		const from = cell(child(document, anchor, "from"))
		const to = cell(child(document, anchor, "to")) ?? from

		if (from !== undefined && to !== undefined) ranges.push([from[0], from[1], to[0], to[1]])
	}

	return ranges
}

interface SheetCells {
	readonly cells: ShownCell[]
	readonly firstHiddenRow: number
	readonly firstHiddenColumn: number
}

function readCells(document: XmlDocument, strings: readonly string[], rich: ReadonlySet<number>): SheetCells {
	const cells: ShownCell[] = []
	const top = document.documentElement()
	const sheetData = child(document, top, "sheetData")
	const cols = child(document, top, "cols")
	let firstHiddenRow = -1
	let firstHiddenColumn = -1

	for (const col of cols < 0 ? [] : children(document, cols, "col"))
		if (truthy(document.attribute(col, "hidden"))) {
			const min = (number(document, col, "min") ?? 1) - 1

			if (firstHiddenColumn < 0 || min < firstHiddenColumn) firstHiddenColumn = min
		}

	if (sheetData < 0) return { cells, firstHiddenRow, firstHiddenColumn }

	let row = -1

	for (const rowNode of children(document, sheetData, "row")) {
		const r = document.attribute(rowNode, "r")

		row = r === undefined ? row + 1 : Number(r) - 1

		if (firstHiddenRow < 0 && truthy(document.attribute(rowNode, "hidden"))) firstHiddenRow = row

		let col = -1

		for (const cellNode of children(document, rowNode, "c")) {
			const ref = document.attribute(cellNode, "r")

			col = ref === undefined ? col + 1 : column(ref.toUpperCase())

			const type = document.attribute(cellNode, "t") ?? "n"
			const v = child(document, cellNode, "v")
			const raw = v < 0 ? undefined : document.textContent(v)
			const is = child(document, cellNode, "is")
			const value = cellValue(type, raw, strings, is < 0 ? undefined : stringText(document, is))
			const style = number(document, cellNode, "s") ?? 0

			if (value !== undefined || style !== 0)
				cells.push({ row, column: col, style, value, rich: type === "s" && raw !== undefined && rich.has(Number(raw)) })
		}
	}

	return { cells, firstHiddenRow, firstHiddenColumn }
}

export async function readShownInput(bytes: Uint8Array): Promise<ShownInput> {
	const pkg = await OpcPackage.open(bytesSource(bytes))
	const workbookPart = relationshipOf(await pkg.relationships(""), "/officeDocument")

	if (workbookPart === undefined) throw new Error("No workbook part")

	const workbook = await pkg.xml(workbookPart)
	const namespaces = new Namespaces(workbook, { lenient: true })
	const relationships = await pkg.relationships(workbookPart)
	const top = workbook.documentElement()
	const workbookPr = child(workbook, top, "workbookPr")
	const date1904 = workbookPr >= 0 && truthy(workbook.attribute(workbookPr, "date1904"))
	const stringsPart = relationshipOf(relationships, "/sharedStrings")
	const strings: string[] = []
	const rich = new Set<number>()

	if (stringsPart !== undefined && pkg.has(stringsPart)) {
		const sst = await pkg.xml(stringsPart)

		for (const si of children(sst, sst.documentElement(), "si")) {
			if (child(sst, si, "r") >= 0) rich.add(strings.length)

			strings.push(stringText(sst, si))
		}
	}

	const stylesPart = relationshipOf(relationships, "/styles")
	const styles =
		stylesPart !== undefined && pkg.has(stylesPart)
			? readStyles(await pkg.xml(stylesPart))
			: { numFmts: new Map(), fonts: [], fills: [], borders: [], cellStyleXfs: [], cellXfs: [], indexedColors: null }
	const themePart = relationshipOf(relationships, "/theme")
	const theme = themePart !== undefined && pkg.has(themePart) ? readTheme(await pkg.xml(themePart)) : null
	const sheets: ShownSheetInput[] = []
	const sheetsNode = child(workbook, top, "sheets")

	for (const sheet of sheetsNode < 0 ? [] : children(workbook, sheetsNode, "sheet")) {
		const id = [
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
			"http://purl.oclc.org/ooxml/officeDocument/relationships"
		]
			.map(uri => namespaces.attribute(sheet, uri, "id"))
			.find(value => value !== undefined)
		const relationship = id === undefined ? undefined : relationships.byId(id)
		const part = relationship?.partName

		if (part === undefined || !pkg.has(part) || !relationship?.type.endsWith("/worksheet")) continue

		const document = await pkg.xml(part)
		const covered = readCovered(document)
		const drawing = relationshipOf(await pkg.relationships(part), "/drawing")

		if (drawing !== undefined && pkg.has(drawing)) covered.push(...readAnchors(await pkg.xml(drawing)))

		sheets.push({ name: workbook.attribute(sheet, "name") ?? "", ...readCells(document, strings, rich), conditional: covered })
	}

	return { date1904, styles, theme, sheets }
}

// The workbook as constants: every formula removed with its cached result kept (a string result as an inline
// string), and its pivot tables and form controls unlinked, so LibreOffice shows the cached values through the
// formats instead of results it calculates, pivots it lays out again or values a control sets.
export function cachedValuesOnly(bytes: Uint8Array): Uint8Array {
	const files = unzipSync(bytes)

	for (const [name, data] of Object.entries(files)) {
		if (/^xl\/worksheets\/[^/]+\.xml$/iu.test(name)) {
			const xml = strFromU8(data)
				.replace(/<(?:\w+:)?f\b[^>]*\/>/gu, "")
				.replace(/<((?:\w+:)?f)\b[^>]*>[^<]*<\/\1>/gu, "")
				.replace(/<c\b([^>]*?)\st="str"([^>/]*)>([^]*?)<\/c>/gu, (_, before: string, after: string, inner: string) => {
					const value = /<v\b[^>]*>([^<]*)<\/v>/u.exec(inner)?.[1] ?? ""

					return `<c${before} t="inlineStr"${after}><is><t xml:space="preserve">${value}</t></is></c>`
				})

			files[name] = strToU8(xml)
		} else if (/^xl\/worksheets\/_rels\/[^/]+\.rels$/iu.test(name)) {
			const xml = strFromU8(data)

			files[name] = strToU8(xml.replace(/<Relationship\b[^>]*\/(?:pivotTable|ctrlProp|vmlDrawing)"[^>]*\/>/gu, ""))
		}
	}

	return zipSync(files, { mtime: ZIP_ENTRY_TIME })
}

// --- Comparison ---

// Why a cell's text differs from LibreOffice's, each with the source that decides it; "unclassified" is a bug
// until shown otherwise.
export type ShownReason =
	// General shows at most 11 characters, sign excluded (ECMA-376 Part 1 §18.8.30 "General" format); LibreOffice
	// shows up to 15 significant digits and a three-digit exponent.
	| "generalDigits"
	// Serials before 1900-03-01: 1 is 1900-01-01, 0 shows day 0 and 60 is 1900-02-29 (§18.17.4.1); LibreOffice
	// counts from 1899-12-30 without the 1900 leap day.
	| "serial1900"
	// A date or time beyond 9999-12-31 or negative fills the cell with "#" (§18.17.4.1; support.microsoft.com "How
	// to correct a ##### error"); LibreOffice shows it.
	| "dateRange"
	// Month and day names of a [$-xxxx] language the app did not inject (core holds en-US only).
	| "languageNames"
	// A/P shows the letter as written (support.microsoft.com "Number format codes": h:mm:ss A/P shows 4:36:03 P);
	// LibreOffice lowercases it.
	| "ampmCase"
	// Codes this formatter does not apply: other calendars and eras ([$-411]g, B2), digit shapes ([DBNum],
	// [$-xx0yyyy]); LibreOffice applies them.
	| "calendarOrDigits"
	// [$-x-sysdate] and [$-x-systime] (Excel's system date and time); LibreOffice does not know them.
	| "systemTag"
	// A boolean cell (§18.18.11 ST_CellType b) shows TRUE or FALSE; LibreOffice keeps booleans as numbers and shows
	// them through the number format.
	| "boolean"
	// An error cell shows its error; LibreOffice reads an error constant as a formula of its own.
	| "errorValue"
	// Text over several lines: LibreOffice keeps it as an edit cell and does not apply the format's text section.
	| "multilineText"
	// An empty string: LibreOffice imports it as an empty cell.
	| "emptyText"
	// A value exactly halfway at its 15 significant digits (what a cell holds: support.microsoft.com "Excel
	// specifications and limits", number precision) rounds away from zero; LibreOffice 7.5 rounds the binary
	// double below it (0.7444999999999999 through 0.0%: 74.5% here and in 26.8, 74.4% in 7.5).
	| "halfwayRounding"
	// A "?" placeholder's space: LibreOffice 26.8 writes it as U+2007 FIGURE SPACE (a digit's width, as Excel's
	// documentation describes "?"), 7.5 as a plain space.
	| "digitSpace"
	// LibreOffice's CSV changes control characters in text: a tab becomes U+0001, a carriage return is dropped.
	| "csvControl"
	| "unclassified"

export interface ShownMismatch {
	readonly sheet: string
	readonly row: number
	readonly column: number
	readonly code: string
	readonly value: CellScalar | undefined
	readonly ours: string
	readonly theirs: string
	readonly reason: ShownReason
}

const MAX_SERIAL_1900 = 2958465
// [$-xxxx] languages whose names are English.
const ENGLISH = new Set(["409", "809", "c09", "1009", "1409", "1809", "2009", "4009", "3c09", "f800", "f400"])

function classify(cell: ShownCell, format: NumFmt, code: string, date1904: boolean, ours: string, theirs: string): ShownReason {
	const { value } = cell

	if (value === undefined) return "unclassified"
	if (ours.includes("\u2007") && ours.replaceAll("\u2007", " ") === theirs) return "digitSpace"
	if (value.type === "b") return "boolean"
	if (value.type === "e") return "errorValue"

	if (value.type === "s") {
		if (value.value === "") return "emptyText"
		if (value.value.includes("\n") && theirs === value.value) return "multilineText"
		if (ours.replaceAll("\t", "\u0001").replaceAll("\r", "") === theirs) return "csvControl"

		return "unclassified"
	}

	const section =
		format.sections.length === 1 || value.value > 0 ? format.sections[0] : value.value < 0 ? format.sections[1] : format.sections[2]
	const kind = section?.kind ?? SectionKind.GENERAL
	const lower = code.toLowerCase()
	const bare = lower
		.replace(/"[^"]*"/gu, "")
		.replace(/\\./gu, "")
		.replace(/general/gu, "")

	if (lower.includes("-x-sysdate") || lower.includes("-x-systime")) return "systemTag"
	if (/\[\$[^\]-]*-[1-9a-f][0-9a-f]{4,7}\]/u.test(lower) || lower.includes("[dbnum")) return "calendarOrDigits"
	if (kind === SectionKind.DATE && /g|b[12]/u.test(bare)) return "calendarOrDigits"

	if (kind === SectionKind.NUMBER && section !== undefined) {
		const d = decompose(value.value)
		const keep = d.point + (section.percent > 0 ? 2 : 0) - 3 * section.thousands + 1 + section.fractionDigits

		if (d.digits.length === keep + 1 && d.digits.endsWith("5")) return "halfwayRounding"
	}

	if (kind === SectionKind.DATE || format.isDate) {
		const serial = value.value

		if (serial < 0 || Math.floor(serial) > (date1904 ? MAX_SERIAL_1900 - 1462 : MAX_SERIAL_1900)) return "dateRange"
		if (!date1904 && serial < 61) return "serial1900"
		if (/a\/p/iu.test(code) && ours.toLowerCase() === theirs.toLowerCase()) return "ampmCase"

		const lcid = /\[\$[^\]-]*-([0-9a-f]{3,4})\]/iu.exec(code)

		if (lcid !== null && !ENGLISH.has((lcid[1] ?? "").toLowerCase())) return "languageNames"
	}

	// Both show the value: LibreOffice with more digits, ours within General's 11 characters (12 with a three-digit
	// exponent) and 6 digits.
	if (kind === SectionKind.GENERAL) {
		const number = (shown: string): number => Number(shown.replace(/[^0-9eE.+-]/gu, ""))
		const close = (shown: number, digits: number): boolean =>
			Math.abs(shown - value.value) <= Math.abs(value.value) * Math.pow(10, -digits)

		const exponent = /E[+-]\d{3}$/u.test(ours) ? 1 : 0

		if (
			close(number(theirs), 14) &&
			close(number(ours), 5) &&
			ours.replace(/[^0-9E.+-]/gu, "").replace(/^-/u, "").length <= 11 + exponent
		)
			return "generalDigits"
	}

	return "unclassified"
}

export interface ShownComparison {
	readonly compared: number
	readonly equal: number
	readonly mismatches: readonly ShownMismatch[]
	// Sheets LibreOffice exported that could not be matched by name, and the reverse.
	readonly unmatchedSheets: number
}

// Compares every cell with a value against LibreOffice's text of it.
export function compareShown(input: ShownInput, shown: ShownWorkbook): ShownComparison {
	const styles = new SheetStyles(input.styles, input.theme)
	const options = { date1904: input.date1904, palette: styles.palette }
	const out = new NumFmtOutput()
	const mismatches: ShownMismatch[] = []
	let compared = 0
	let equal = 0
	let unmatchedSheets = 0

	for (const sheet of input.sheets) {
		const theirs =
			shown.sheets.length === 1 && input.sheets.length === 1 ? shown.sheets[0] : shown.sheets.find(item => item.name === sheet.name)

		if (theirs === undefined) {
			unmatchedSheets++
			continue
		}

		for (const cell of sheet.cells) {
			const text = theirs.text[cell.row]?.[cell.column]

			if (text === undefined || cell.value === undefined) continue

			const code = styles.numFmtCode(cell.style)
			const format = styles.numFmt(cell.style)
			let ours: string

			switch (cell.value.type) {
				case "n":
					formatNumberInto(format, cell.value.value, out, options)
					ours = out.text
					break
				case "s":
					if (cell.rich) ours = cell.value.value
					else {
						formatTextInto(format, cell.value.value, out, options)
						ours = out.text
					}
					break
				case "b":
					ours = cell.value.value ? "TRUE" : "FALSE"
					break
				default:
					ours = cell.value.value
			}

			compared++

			if (ours === text) {
				equal++
				continue
			}

			mismatches.push({
				sheet: sheet.name,
				row: cell.row,
				column: cell.column,
				code,
				value: cell.value,
				ours,
				theirs: text,
				reason: classify(cell, format, code, input.date1904, ours, text)
			})
		}
	}

	return { compared, equal, mismatches, unmatchedSheets }
}

export interface FillComparison {
	readonly compared: number
	readonly equal: number
	// Within 2 per channel (the tint's rounding).
	readonly close: number
	readonly mismatches: readonly {
		readonly sheet: string
		readonly row: number
		readonly column: number
		readonly ours: string
		readonly theirs: string
	}[]
}

function inside(ranges: ShownSheetInput["conditional"], row: number, column: number): boolean {
	return ranges.some(([top, left, bottom, right]) => row >= top && row <= bottom && column >= left && column <= right)
}

// Solid fills against the background LibreOffice's HTML export gives each cell, outside conditional formatting.
export function compareFills(input: ShownInput, shown: ShownWorkbook): FillComparison {
	const styles = new SheetStyles(input.styles, input.theme)
	const mismatches: { sheet: string; row: number; column: number; ours: string; theirs: string }[] = []
	let compared = 0
	let equal = 0
	let close = 0

	for (const sheet of input.sheets) {
		const theirs =
			shown.sheets.length === 1 && input.sheets.length === 1 ? shown.sheets[0] : shown.sheets.find(item => item.name === sheet.name)

		if (theirs?.fills === undefined) continue

		for (const cell of sheet.cells) {
			const fill = styles.renderStyle(cell.style).fill

			if (fill === undefined || inside(sheet.conditional, cell.row, cell.column)) continue
			if (
				(sheet.firstHiddenRow >= 0 && cell.row >= sheet.firstHiddenRow) ||
				(sheet.firstHiddenColumn >= 0 && cell.column >= sheet.firstHiddenColumn)
			)
				continue

			const lo = theirs.fills.get(`${cell.row},${cell.column}`)

			if (lo === undefined) continue

			const ours = `#${fill.toString(16).padStart(6, "0")}`.toUpperCase()
			const other = parseInt(lo.slice(1), 16)

			compared++

			if (ours === lo.toUpperCase()) {
				equal++
				continue
			}

			if ([16, 8, 0].every(shift => Math.abs(((fill >> shift) & 0xff) - ((other >> shift) & 0xff)) <= 2)) {
				close++
				continue
			}

			mismatches.push({ sheet: sheet.name, row: cell.row, column: cell.column, ours, theirs: lo.toUpperCase() })
		}
	}

	return { compared, equal, close, mismatches }
}
