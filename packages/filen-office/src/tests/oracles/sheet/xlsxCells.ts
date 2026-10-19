import { bytesSource, OpcPackage, type Relationships } from "@filen/office/core/opc"
import { Namespaces, type XmlDocument } from "@filen/office/core/xml"

// A plain reader of what the cells of an xlsx package hold as written: values, cached formula results and
// formula text, sheet by sheet (ECMA-376 Part 1 §18.3.1.4 c, §18.3.1.40 f, §18.4 shared strings). It is the
// oracles' independent view of a file, not the engine's: no number formats, no recalculation.

export type CellScalar =
	| { readonly type: "n"; readonly value: number }
	| { readonly type: "s"; readonly value: string }
	| { readonly type: "b"; readonly value: boolean }
	// The error's text as written ("#DIV/0!", "#N/A", ...).
	| { readonly type: "e"; readonly value: string }

export interface CellFormula {
	// As written; empty for a shared formula's dependent cells (the master's text, shifted, applies).
	readonly text: string
	readonly kind: "normal" | "shared" | "array" | "dataTable"
	readonly ref?: string
	readonly si?: string
}

export interface XlsxCell {
	// 0-based.
	readonly row: number
	readonly column: number
	// undefined: no cached value (no <v>, or an empty one on a non-string cell).
	readonly value: CellScalar | undefined
	readonly formula?: CellFormula
}

export interface XlsxSheet {
	readonly name: string
	readonly part: string
	readonly state: string
	readonly cells: readonly XlsxCell[]
}

export interface XlsxWorkbook {
	readonly application: string | undefined
	readonly appVersion: string | undefined
	readonly date1904: boolean
	// calcPr: calcId (the calculation engine that last calculated), calcMode, fullCalcOnLoad.
	readonly calcId: string | undefined
	readonly calcMode: string | undefined
	readonly fullCalcOnLoad: boolean
	readonly hasCalcChain: boolean
	readonly sheets: readonly XlsxSheet[]
}

const RELATIONSHIP_NAMESPACES = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"http://purl.oclc.org/ooxml/officeDocument/relationships"
]

// The first internal target of a relationship type ending in one of `suffixes` (Transitional or Strict spelling).
function relationshipOf(relationships: Relationships, ...suffixes: string[]): string | undefined {
	return relationships.all.find(
		relationship => suffixes.some(suffix => relationship.type.endsWith(suffix)) && relationship.partName !== undefined
	)?.partName
}

function* children(document: XmlDocument, node: number, local?: string): Generator<number> {
	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		if (local === undefined || document.localName(child) === local) yield child
}

function child(document: XmlDocument, node: number, local: string): number {
	for (const found of children(document, node, local)) return found

	return -1
}

// A shared string's or inline string's text: its runs' <t>, phonetic runs (<rPh>) left out.
function stringText(document: XmlDocument, node: number): string {
	let out = ""

	for (const item of children(document, node)) {
		const local = document.localName(item)

		if (local === "t") out += document.textContent(item)
		else if (local === "r") for (const t of children(document, item, "t")) out += document.textContent(t)
	}

	return out
}

const LETTERS = /^([A-Za-z]{1,3})(\d+)$/

// 0-based row and column of an A1 reference, or undefined when it is not one.
export function parseA1(ref: string): { row: number; column: number } | undefined {
	const match = LETTERS.exec(ref.replace(/\$/g, ""))

	if (match === null) return undefined

	let column = 0

	for (const char of (match[1] as string).toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64

	return { row: Number(match[2]) - 1, column: column - 1 }
}

export function formatA1(row: number, column: number): string {
	let name = ""

	for (let index = column + 1; index > 0; index = Math.floor((index - 1) / 26)) name = String.fromCharCode(65 + ((index - 1) % 26)) + name

	return `${name}${row + 1}`
}

function cellValue(type: string, raw: string | undefined, strings: readonly string[], inline: string | undefined): CellScalar | undefined {
	switch (type) {
		case "s": {
			const value = raw === undefined || raw === "" ? undefined : strings[Number(raw)]

			return value === undefined ? undefined : { type: "s", value }
		}
		case "str":
			return raw === undefined ? undefined : { type: "s", value: raw }
		case "inlineStr":
			return inline === undefined ? undefined : { type: "s", value: inline }
		case "b":
			return raw === undefined || raw === "" ? undefined : { type: "b", value: raw.trim() === "1" || raw.trim() === "true" }
		case "e":
			return raw === undefined || raw === "" ? undefined : { type: "e", value: raw.trim() }
		// ISO 8601 date cells keep their text: the reader does not convert dates.
		case "d":
			return raw === undefined || raw === "" ? undefined : { type: "s", value: raw }
		default: {
			if (raw === undefined || raw.trim() === "") return undefined

			const value = Number(raw)

			return Number.isNaN(value) ? undefined : { type: "n", value }
		}
	}
}

function readSheet(document: XmlDocument, strings: readonly string[], keep: (cell: XlsxCell) => boolean): XlsxCell[] {
	const cells: XlsxCell[] = []
	const sheetData = child(document, document.documentElement(), "sheetData")

	if (sheetData < 0) return cells

	let row = -1

	for (const rowNode of children(document, sheetData, "row")) {
		const r = document.attribute(rowNode, "r")

		row = r === undefined ? row + 1 : Number(r) - 1

		let column = -1

		for (const cellNode of children(document, rowNode, "c")) {
			const ref = document.attribute(cellNode, "r")
			const at = ref === undefined ? undefined : parseA1(ref)

			column = at === undefined ? column + 1 : at.column

			const cellRow = at === undefined ? row : at.row
			const v = child(document, cellNode, "v")
			const is = child(document, cellNode, "is")
			const f = child(document, cellNode, "f")
			const value = cellValue(
				document.attribute(cellNode, "t") ?? "n",
				v < 0 ? undefined : document.textContent(v),
				strings,
				is < 0 ? undefined : stringText(document, is)
			)
			let formula: CellFormula | undefined

			if (f >= 0) {
				const kind = document.attribute(f, "t") ?? "normal"
				const fRef = document.attribute(f, "ref")
				const si = document.attribute(f, "si")

				formula = {
					text: document.textContent(f),
					kind: kind === "shared" || kind === "array" || kind === "dataTable" ? kind : "normal",
					...(fRef !== undefined && { ref: fRef }),
					...(si !== undefined && { si })
				}
			}

			const cell: XlsxCell = { row: cellRow, column, value, ...(formula !== undefined && { formula }) }

			if (keep(cell)) cells.push(cell)
		}
	}

	return cells
}

// Every sheet's cells that `keep` accepts (all by default); chartsheets and dialogsheets have none.
export async function readXlsxCells(bytes: Uint8Array, keep: (cell: XlsxCell) => boolean = () => true): Promise<XlsxWorkbook> {
	const pkg = await OpcPackage.open(bytesSource(bytes))
	const root = await pkg.relationships("")
	const workbookPart = relationshipOf(root, "/officeDocument")

	if (workbookPart === undefined) throw new Error("No workbook part")

	const appPart = relationshipOf(root, "/extended-properties", "/extendedProperties")
	let application: string | undefined
	let appVersion: string | undefined

	if (appPart !== undefined && pkg.has(appPart)) {
		const app = await pkg.xml(appPart)
		const top = app.documentElement()
		const applicationNode = child(app, top, "Application")
		const versionNode = child(app, top, "AppVersion")

		application = applicationNode < 0 ? undefined : app.textContent(applicationNode).trim()
		appVersion = versionNode < 0 ? undefined : app.textContent(versionNode).trim()
	}

	const workbook = await pkg.xml(workbookPart)
	const namespaces = new Namespaces(workbook, { lenient: true })
	const relationships = await pkg.relationships(workbookPart)
	const top = workbook.documentElement()
	const workbookPr = child(workbook, top, "workbookPr")
	const calcPr = child(workbook, top, "calcPr")
	const date1904 = workbookPr >= 0 && ["1", "true"].includes(workbook.attribute(workbookPr, "date1904") ?? "")
	const stringsPart = relationshipOf(relationships, "/sharedStrings")
	const strings: string[] = []

	if (stringsPart !== undefined && pkg.has(stringsPart)) {
		const sst = await pkg.xml(stringsPart)

		for (const si of children(sst, sst.documentElement(), "si")) strings.push(stringText(sst, si))
	}

	const sheets: XlsxSheet[] = []
	const sheetsNode = child(workbook, top, "sheets")

	for (const sheet of sheetsNode < 0 ? [] : children(workbook, sheetsNode, "sheet")) {
		const id = RELATIONSHIP_NAMESPACES.map(uri => namespaces.attribute(sheet, uri, "id")).find(value => value !== undefined)
		const relationship = id === undefined ? undefined : relationships.byId(id)
		const part = relationship?.partName

		if (part === undefined || !pkg.has(part) || !relationship?.type.endsWith("/worksheet")) continue

		sheets.push({
			name: workbook.attribute(sheet, "name") ?? "",
			part,
			state: workbook.attribute(sheet, "state") ?? "visible",
			cells: readSheet(await pkg.xml(part), strings, keep)
		})
	}

	return {
		application,
		appVersion,
		date1904,
		calcId: calcPr < 0 ? undefined : workbook.attribute(calcPr, "calcId"),
		calcMode: calcPr < 0 ? undefined : workbook.attribute(calcPr, "calcMode"),
		fullCalcOnLoad: calcPr >= 0 && ["1", "true"].includes(workbook.attribute(calcPr, "fullCalcOnLoad") ?? ""),
		hasCalcChain: relationshipOf(relationships, "/calcChain") !== undefined,
		sheets
	}
}
