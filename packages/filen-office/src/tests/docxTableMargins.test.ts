import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { unzipSync, zipSync } from "fflate"
import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { docxSchema, type DocxEditorCommand } from "@filen/office/core/docx/edit"
import { tableBorderOf, type PagedEditor } from "@filen/office/core/docx/paged"
import { tableRightToLeft } from "../core/docx/edit/snapshot/table"
import { NO_SOURCE } from "../core/docx/edit/schema"
import { drawingState } from "../core/docx/objects/edit"
import { CORPUS_DIR } from "./corpusDir"
import { STYLES_PART } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { layOut, layoutDigest } from "./layoutFixtures"
import { pagedEditor } from "./pagedFixtures"

// A table's cell margins (w:tblCellMar, else its style hierarchy's, else the standard's 115 twips a side) bound how
// narrow a column gets and how wide what goes into a cell may be, and layout pads cells by the same; a table's
// direction from its style is each document's own.

const nodes = docxSchema.nodes
const TDF131203 = join(CORPUS_DIR, "docx", "lo__tdf131203.docx")
const EMU_PER_TWIP = 635
const STYLES_RELATIONSHIP = [{ id: "rIdStyles", type: "styles", target: "styles.xml" }] as const

const table = (tblPr: string, widths: readonly number[]): string =>
	`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${tblPr}</w:tblPr><w:tblGrid>${widths.map(w => `<w:gridCol w:w="${String(w)}"/>`).join("")}</w:tblGrid>` +
	`<w:tr>${widths.map((w, k) => `<w:tc><w:tcPr><w:tcW w:w="${String(w)}" w:type="dxa"/></w:tcPr>${p(`c${String(k)}`)}</w:tc>`).join("")}</w:tr></w:tbl>`
const MARGINS_400 = '<w:tblCellMar><w:left w:w="400" w:type="dxa"/><w:right w:w="400" w:type="dxa"/></w:tblCellMar>'

function u32(value: number): number[] {
	return [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255]
}

// A PNG header of a width x height picture: placing it needs no pixels.
function png(width: number, height: number): Uint8Array {
	const ihdr = [...u32(width), ...u32(height), 8, 6, 0, 0, 0]

	return new Uint8Array([
		0x89,
		0x50,
		0x4e,
		0x47,
		0x0d,
		0x0a,
		0x1a,
		0x0a,
		...u32(13),
		73,
		72,
		68,
		82,
		...ihdr,
		0,
		0,
		0,
		0,
		...u32(0),
		73,
		69,
		78,
		68,
		0,
		0,
		0,
		0
	])
}

function caretIn(editor: PagedEditor, text: string): void {
	let at = -1

	editor.state.doc.descendants((node, pos) => {
		if (at < 0 && node.type === nodes.paragraph && node.textContent === text) at = pos + 1

		return at < 0
	})
	editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)))
}

function tables(doc: Node): { node: Node; pos: number }[] {
	const out: { node: Node; pos: number }[] = []

	doc.descendants((node, pos) => {
		if (node.type === nodes.table) out.push({ node, pos })

		return true
	})

	return out
}

async function run(editor: PagedEditor, command: DocxEditorCommand): Promise<boolean> {
	return await editor.execute(command)
}

async function opened(tblPr: string, styles?: string): Promise<PagedEditor> {
	const body = p("Before") + table(tblPr, [4000, 4000]) + p("After")

	return (
		await pagedEditor(
			buildDocx(styles === undefined ? { body } : { body, parts: [STYLES_PART(styles)], relationships: STYLES_RELATIONSHIP })
		)
	).editor
}

describe("table cell margins", () => {
	it("keeps a dragged column as wide as its cell margins: the table's, its style hierarchy's per side, else 115 twips a side", async () => {
		const NORMAL_TABLE =
			'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/>' +
			'<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
		const NARROW =
			'<w:style w:type="table" w:styleId="Base"><w:name w:val="Base"/><w:tblPr><w:tblCellMar><w:right w:w="300" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
			'<w:style w:type="table" w:styleId="Narrow"><w:name w:val="Narrow"/><w:basedOn w:val="Base"/><w:tblPr><w:tblCellMar><w:left w:w="50" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'

		for (const [tblPr, styles, min] of [
			[MARGINS_400, undefined, 800],
			["", undefined, 230],
			["", NORMAL_TABLE, 216],
			['<w:tblStyle w:val="Narrow"/>', NARROW, 350],
			['<w:tblStyle w:val="Base"/>', NARROW, 415]
		] as const) {
			const editor = await opened(tblPr, styles)
			const [first] = tables(editor.state.doc)

			if (first === undefined) throw new Error("no table")

			const border = tableBorderOf(editor, { kind: "tableBorder", story: editor.activeStory, pos: first.pos, part: "column:0" })

			expect(border?.minTwips).toBe(min)

			caretIn(editor, "c0")
			expect(await run(editor, { type: "table.column.width", index: 0, twips: 10 })).toBe(true)
			expect(editor.snapshot().context.table?.widths).toEqual([min, 8000 - min])
			editor.dispose()
		}
	})

	it("fits a table inserted in a cell and a picture inserted in one within the cell less its table's margins", async () => {
		for (const [tblPr, margins] of [
			[MARGINS_400, 800],
			["", 230]
		] as const) {
			const editor = await opened(tblPr)

			caretIn(editor, "c0")
			expect(await run(editor, { type: "table.insert", rows: 1, cols: 2 })).toBe(true)

			const nested = tables(editor.state.doc)[1]?.node

			expect(nested?.attrs["grid"]).toEqual([(4000 - margins) / 2, (4000 - margins) / 2])
			expect(await run(editor, { type: "history.undo" })).toBe(true)

			caretIn(editor, "c0")

			const media = editor.addMedia(png(4000, 4000))

			expect(await run(editor, { type: "image.insert", media: media as number })).toBe(true)

			let width: number | undefined

			editor.state.doc.descendants(node => {
				if (node.type === nodes.drawing) width = drawingState(editor.session.context, node).cx

				return width === undefined
			})
			expect(width).toBe((4000 - margins) * EMU_PER_TWIP)
			editor.dispose()
		}
	})
})

describe("cell margins in layout", () => {
	const padded = (margins: string): string =>
		p("Before") +
		`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${margins}</w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>` +
		`<w:tr>${[0, 1].map(k => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>${p(`cell ${String(k)} with words that wrap in a narrow column`)}</w:tc>`).join("")}</w:tr></w:tbl>` +
		p("After")
	const explicit = (twips: number): string =>
		`<w:tblCellMar><w:left w:w="${String(twips)}" w:type="dxa"/><w:right w:w="${String(twips)}" w:type="dxa"/></w:tblCellMar>`

	it("pads a cell nothing gives margins 115 twips a side, as the commands take it", async () => {
		const none = layoutDigest(await layOut(buildDocx({ body: padded("") })))

		expect(none).toEqual(layoutDigest(await layOut(buildDocx({ body: padded(explicit(115)) }))))
		expect(none).not.toEqual(layoutDigest(await layOut(buildDocx({ body: padded(explicit(0)) }))))
	})

	it.skipIf(!existsSync(TDF131203))(
		"lays out a document without table styles as with 115 twips margins written on its tables",
		async () => {
			const bytes = readFileSync(TDF131203)
			const files = unzipSync(bytes)
			const xml = new TextDecoder().decode(files["word/document.xml"])

			expect(xml).not.toContain("tblCellMar")
			expect(files["word/styles.xml"]).toBeUndefined()

			const written = zipSync({
				...files,
				"word/document.xml": new TextEncoder().encode(xml.replace(/<\/w:tblPr>/g, `${explicit(115)}</w:tblPr>`))
			})

			expect(layoutDigest(await layOut(bytes))).toEqual(layoutDigest(await layOut(written)))
		}
	)
})

describe("table direction from a style", () => {
	it("reads each document's own style for a table the editor made", async () => {
		const styles = (rtl: boolean): string =>
			`<w:style w:type="table" w:styleId="TS"><w:name w:val="TS"/><w:tblPr>${rtl ? "<w:bidiVisual/>" : ""}</w:tblPr></w:style>`

		for (const rtl of [false, true, false]) {
			const { editor } = await pagedEditor(
				buildDocx({ body: p("x"), parts: [STYLES_PART(styles(rtl))], relationships: STYLES_RELATIONSHIP })
			)
			const session = editor.session
			const made = nodes.table.create(
				{ src: NO_SOURCE, styleId: "TS" },
				nodes.table_row.create({ src: NO_SOURCE }, nodes.table_cell.create({ src: NO_SOURCE }, nodes.paragraph.create()))
			)

			expect(tableRightToLeft(session.story.part, session.context, made)).toBe(rtl)
			editor.dispose()
		}
	})
})
