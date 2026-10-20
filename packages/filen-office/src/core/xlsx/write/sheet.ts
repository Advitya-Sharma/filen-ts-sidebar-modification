import { encodeUtf8 } from "../../xml/index"
import type { SheetRange } from "../../sheet/protocol"
import { runSliced, type SliceClock } from "../../slicing"
import type { XlsxSection, XlsxSheetChanges, XlsxSheetEntry, XlsxSheetModel, XlsxWorkbookModel } from "../model/types"
import { assembleSections } from "./assemble"
import { namespacesCompatible, usedRange, writeSheetDataSteps, type CarriedPart, type WrittenRow } from "./cells"
import { XlsxWriteError } from "./failure"
import { Pieces, rootTag, toBytes, utf8Part } from "./markup"
import { namespacesOf } from "./namespaces"
import { CHARTSHEET, DIALOGSHEET, WORKSHEET } from "./order.generated"
import {
	readDimension,
	writeCols,
	writeDimension,
	writeMergeCells,
	writeSheetFormatPr,
	writeSheetPr,
	writeSheetProtection,
	writeSheetViews
} from "./sections"

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

type OwnWriter = (sheet: XlsxSheetModel, section: Uint8Array | null, prefix: string) => Uint8Array | null

// The sections the writer serialises itself (an edit marks them without a writer of its own).
export const OWN_SHEET_SECTIONS: ReadonlyMap<string, OwnWriter> = new Map<string, OwnWriter>([
	["sheetPr", (sheet, section, prefix) => writeSheetPr(sheet, section, prefix)],
	["dimension", (sheet, section, prefix) => writeDimension(usedRange(sheet), section, prefix)],
	["sheetViews", (sheet, section, prefix) => writeSheetViews(sheet.view, section, prefix)],
	["sheetFormatPr", (sheet, section, prefix) => writeSheetFormatPr(sheet.format, section, prefix)],
	["cols", (sheet, section, prefix) => writeCols(sheet.cols, section, prefix)],
	["mergeCells", (sheet, section, prefix) => writeMergeCells(sheet.merges, section, prefix)],
	["sheetProtection", (sheet, section, prefix) => writeSheetProtection(sheet.protection, section, prefix)]
])

// When little of a large source part is copied into the new one (most rows rewritten or renumbered), those pieces
// become copies of their own, so the inflated source is not held through the check and the zip.
function releaseSource(pieces: Pieces, source: Uint8Array): void {
	if (source.length < RELEASE_AT) return

	let viewed = 0

	for (const item of pieces.items) if (item.buffer === source.buffer) viewed += item.length

	if (viewed * 4 > source.length) return

	pieces.items.forEach((item, k) => {
		if (item.buffer === source.buffer) pieces.items[k] = item.slice()
	})
}

const RELEASE_AT = 8 << 20

export function sheetOrder(entry: XlsxSheetEntry): readonly string[] {
	return entry.kind === "chartsheet" ? CHARTSHEET : entry.kind === "dialogsheet" ? DIALOGSHEET : WORKSHEET
}

export interface WrittenSheet {
	readonly entry: XlsxSheetEntry
	readonly part: string
	readonly pieces: Pieces
	readonly prefix: string
	// The root's start tag (namespace declarations for checking fragments).
	readonly rootOpen: Uint8Array
	readonly names: readonly string[]
	readonly sourceNames: readonly string[]
	// Sections written anew (not copied), by name; removed ones are absent.
	readonly written: ReadonlyMap<string, Uint8Array>
	// Rows rewritten or renumbered (null: sheetData was copied).
	readonly rows: readonly WrittenRow[] | null
	// A rewritten row added, changed or dropped a formula.
	readonly formulasChanged: boolean
}

function contains(outer: SheetRange, inner: SheetRange): boolean {
	return (
		outer.startRow <= inner.startRow && outer.startCol <= inner.startCol && outer.endRow >= inner.endRow && outer.endCol >= inner.endCol
	)
}

function sectionBytes(source: Uint8Array, section: XlsxSection): Uint8Array {
	return section.bytes ?? source.subarray(section.start, section.end)
}

// A worksheet part: `source` is the part's inflated bytes (null for a sheet added since open). Without changes
// to its rows, sheetData is copied whole.
export async function writeSheet(
	entry: XlsxSheetEntry,
	workbook: XlsxWorkbookModel,
	changes: XlsxSheetChanges | undefined,
	source: Uint8Array | null,
	clock?: SliceClock
): Promise<WrittenSheet> {
	const sheet = entry.model

	if (sheet === null) throw new XlsxWriteError("write", `Sheet ${entry.name} is not loaded`)

	let bytes = source
	let sections: readonly XlsxSection[] = sheet.sections
	const fresh = source === null

	if (bytes === null) {
		const namespaces = namespacesOf(workbook.strict)
		const root = `<${entry.kind === "chartsheet" ? "chartsheet" : "worksheet"} xmlns="${namespaces.main}" xmlns:r="${namespaces.relationships}">`
		const head = encodeUtf8(`${DECLARATION}${root}`)
		const data = encodeUtf8("<sheetData/>")
		const tail = encodeUtf8(`</${entry.kind === "chartsheet" ? "chartsheet" : "worksheet"}>`)

		bytes = new Uint8Array(head.length + data.length + tail.length)
		bytes.set(head, 0)
		bytes.set(data, head.length)
		bytes.set(tail, head.length + data.length)
		sections =
			entry.kind === "chartsheet" ? [] : [{ name: "sheetData", start: head.length, end: head.length + data.length, bytes: null }]
	}

	const part = bytes
	const { tag, prefix } = rootTag(part)
	const byName = new Map(sections.map(section => [section.name, section]))
	const replacements = new Map<string, Uint8Array>()
	const marked = new Map(changes?.sections ?? [])

	if (fresh) {
		for (const name of OWN_SHEET_SECTIONS.keys()) if (!marked.has(name) && entry.kind !== "chartsheet") marked.set(name, null)
	}

	for (const [name, writer] of marked) {
		const section = byName.get(name)
		const current = section === undefined ? null : sectionBytes(part, section)

		if (writer !== null) {
			replacements.set(name, toBytes(writer(sheet, workbook)))

			continue
		}

		const own = OWN_SHEET_SECTIONS.get(name)

		if (own === undefined) throw new XlsxWriteError("write", `No writer for the ${name} section of ${entry.name}`)

		const written = own(sheet, current, prefix)

		if (written !== null && written !== current) replacements.set(name, written)
	}

	// Content the edits added outside the recorded dimension widens it (a dimension that is too small makes a
	// reader drop cells).
	const dimension = byName.get("dimension")

	if (changes?.rows === true && !marked.has("dimension") && dimension !== undefined) {
		const recorded = readDimension(sectionBytes(part, dimension))
		const used = usedRange(sheet)

		if (used !== null && (recorded === null || !contains(recorded, used))) {
			const union =
				recorded === null
					? used
					: {
							startRow: Math.min(recorded.startRow, used.startRow),
							startCol: Math.min(recorded.startCol, used.startCol),
							endRow: Math.max(recorded.endRow, used.endRow),
							endCol: Math.max(recorded.endCol, used.endCol)
						}

			replacements.set("dimension", writeDimension(union, sectionBytes(part, dimension), prefix))
		}
	}

	// Filled when sheetData is written (an object: the assembly calls the writer back).
	const data: { rows: WrittenRow[] | null; formulasChanged: boolean } = { rows: null, formulasChanged: false }
	const inPlace = new Map<string, (pieces: Pieces) => void>()
	const sheetData = byName.get("sheetData")

	if ((changes?.rows === true || fresh) && sheetData !== undefined) {
		// Written ahead, in slices, and placed where the assembly reaches sheetData.
		const rows = new Pieces()
		// Other sheets' parts pastes carried cells from, read for their markup.
		const carriedParts = new Map<string, CarriedPart>()

		for (const other of sheet.store.carriedParts()) {
			const entryOfPart = workbook.parts.get(other)?.source

			if (other === entry.part || part === null || entryOfPart === undefined || entryOfPart === null) continue

			const bytes = utf8Part(await workbook.parts.read(other, clock))

			carriedParts.set(other, { bytes, compatible: namespacesCompatible(part, bytes) })
		}

		const written = await runSliced(writeSheetDataSteps(rows, sheet, prefix, part, sheetData, carriedParts), clock)

		data.rows = written.rows
		data.formulasChanged = written.formulasChanged
		inPlace.set("sheetData", pieces => {
			for (const item of rows.items) pieces.push(item)
		})
	}

	const order = sheetOrder(entry)
	const assembled = assembleSections({ bytes: part, sections, order, replacements, inPlace })

	releaseSource(assembled.pieces, part)

	const written = new Map<string, Uint8Array>()

	for (const [name, markup] of replacements) if (markup.length > 0) written.set(name, markup)

	return {
		entry,
		part: entry.part,
		pieces: assembled.pieces,
		prefix,
		rootOpen: part.slice(tag.start, tag.end),
		names: assembled.names,
		sourceNames: sections.map(section => section.name),
		written,
		rows: data.rows,
		formulasChanged: data.formulasChanged
	}
}
