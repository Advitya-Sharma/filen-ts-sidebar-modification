import { unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema, type BlockProof } from "@filen/office/core/docx/edit"
import { freshSlice } from "../core/docx/edit/invariants/core"
import { tableEdge } from "../core/docx/edit/schema"
import { rowOverrides } from "../core/docx/edit/tables/map"
import { RevisionFlag } from "../core/docx/model/index"
import { resolveChanges, storyReviewIndex } from "../core/docx/review/accept"
import { paragraphChange } from "../core/docx/review/track"
import { saveDocx } from "@filen/office/core/docx/write"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"

// Word and LibreOffice show two tables with no paragraph between them as one table (both LibreOffice versions,
// also when the first ends a content control): a table inserted right after a table
// keeps a paragraph between them, also after a table the proof locked and after a control ending with one.

const TABLE =
	'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
const SDT_OPEN = '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>'
const SDT_CLOSE = "</w:sdtContent></w:sdt>"
// Locks every table (as the writer's proof does for a table it cannot write again).
const LOCK_TABLES: BlockProof = {
	check: node => (node.type === docxSchema.nodes.table ? "proof" : undefined)
}

async function insertAfter(body: string, proof?: BlockProof): Promise<{ names: string[]; xml: string }> {
	const opened = await openBytes(buildDocx({ body }))
	const session = new DocxEditorSession(opened.context, proof === undefined ? {} : { proof })
	let at = -1

	session.state.doc.descendants((node, pos) => {
		if (at < 0 && node.isTextblock && node.textContent === "After") at = pos + 1
	})
	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at)))
	expect(session.execute({ type: "table.insert", rows: 1, cols: 2 })).toBe(true)

	const names: string[] = []

	session.state.doc.descendants(node => {
		if (node.isBlock && node.type !== docxSchema.nodes.table_row && node.type !== docxSchema.nodes.table_cell) {
			names.push(node.type === docxSchema.nodes.locked_block ? `locked ${String(node.attrs["kind"])}` : node.type.name)

			return node.type === docxSchema.nodes.sdt
		}

		return false
	})

	const result = await saveDocx({
		pkg: opened.pkg,
		context: opened.context,
		sessions: [session]
	})
	const part = result.bytes === undefined ? undefined : unzipSync(result.bytes)["word/document.xml"]

	return {
		names,
		xml: part === undefined ? "" : new TextDecoder().decode(part)
	}
}

describe("tables side by side", () => {
	it("tells a table's edges: a table, a locked one, a control ending or starting with one", () => {
		const table = docxSchema.nodes.table.createAndFill()
		const locked = docxSchema.nodes.locked_block.create({ kind: "table" })
		const lockedParagraph = docxSchema.nodes.locked_block.create({
			kind: "paragraph"
		})
		const paragraph = docxSchema.nodes.paragraph.create()

		expect(table).not.toBeNull()

		if (table === null) return

		const control = docxSchema.nodes.sdt.create(null, [paragraph, table])

		expect(tableEdge(table, true)).toBe(true)
		expect(tableEdge(locked, false)).toBe(true)
		expect(tableEdge(lockedParagraph, true)).toBe(false)
		expect(tableEdge(control, true)).toBe(true)
		expect(tableEdge(control, false)).toBe(false)
		expect(tableEdge(null, true)).toBe(false)
	})

	it("keeps a paragraph between a locked table and a table inserted right after it", async () => {
		const { names } = await insertAfter(TABLE + p("After"), LOCK_TABLES)

		expect(names).toEqual(["locked table", "paragraph", "table", "paragraph"])
	})

	it("keeps a paragraph between a control ending with a table and a table inserted right after it", async () => {
		const { names, xml } = await insertAfter(SDT_OPEN + TABLE + SDT_CLOSE + p("After"))

		expect(names).toEqual(["sdt", "table", "paragraph", "table", "paragraph"])
		expect(xml.replace(/\s+</g, "<")).not.toMatch(/<\/w:tbl>(?:<\/w:sdtContent><\/w:sdt>)?<w:tbl>/)
	})

	it("keeps a paragraph before a table inserted at the start of a control right after a table", async () => {
		const { names } = await insertAfter(TABLE + SDT_OPEN + p("After") + SDT_CLOSE)

		expect(names).toEqual(["table", "sdt", "paragraph", "table", "paragraph"])
	})
})

describe("host edits that put tables side by side", () => {
	async function session(body: string): Promise<DocxEditorSession> {
		const opened = await openBytes(buildDocx({ body }))

		return new DocxEditorSession(opened.context)
	}

	// The top-level blocks, and their positions.
	function blocks(state: DocxEditorSession["state"]): { name: string; text: string; pos: number; size: number }[] {
		const out: { name: string; text: string; pos: number; size: number }[] = []

		state.doc.forEach((node, pos) =>
			out.push({
				name: node.type.name,
				text: node.textContent,
				pos,
				size: node.nodeSize
			})
		)

		return out
	}

	const names = (editor: DocxEditorSession): string[] =>
		blocks(editor.state).map(block => (block.name === "paragraph" ? `p:${block.text}` : block.name))

	it("keeps a paragraph between two tables when the one between them is deleted", async () => {
		const editor = await session(TABLE + p("Gone") + TABLE + p("After"))
		const [, between] = blocks(editor.state)

		if (between === undefined) throw new Error("no blocks")

		editor.dispatch(editor.state.tr.delete(between.pos, between.pos + between.size))
		expect(names(editor)).toEqual(["table", "p:", "table", "p:After"])
	})

	it("puts a paragraph between a table and one pasted or dropped right after it", async () => {
		const editor = await session(TABLE + p("Between") + p("After") + TABLE)
		const [, between, , last] = blocks(editor.state)
		const table = last === undefined ? null : editor.state.doc.nodeAt(last.pos)

		if (between === undefined || last === undefined || table === null) throw new Error("no blocks")

		// Dropped at the start of "Between": right after the first table.
		editor.dispatch(editor.state.tr.delete(last.pos, last.pos + last.size).insert(between.pos, table))
		expect(names(editor)).toEqual(["table", "p:", "table", "p:Between", "p:After"])
	})

	it("leaves tables already side by side in the file, and typing, alone", async () => {
		const editor = await session(TABLE + TABLE + p("After") + p("More"))
		const after = blocks(editor.state)[2]

		if (after === undefined) throw new Error("no blocks")

		editor.dispatch(editor.state.tr.insertText("x", after.pos + 1))

		const more = blocks(editor.state)[3]

		if (more === undefined) throw new Error("no blocks")

		// A block-level step next to them: the last paragraph joined to the one before.
		editor.dispatch(editor.state.tr.delete(more.pos - 1, more.pos + 1))
		expect(names(editor)).toEqual(["table", "table", "p:xAfterMore"])
	})

	it("records the paragraph a table pasted while tracking gets as part of the paste, which rejecting takes away", async () => {
		const opened = await openBytes(buildDocx({ body: TABLE + p("Between") + p("After") }))
		const editor = new DocxEditorSession(opened.context)
		const [table, between] = blocks(editor.state)
		const copy = table === undefined ? null : editor.state.doc.nodeAt(table.pos)

		if (between === undefined || table === undefined || copy === null) throw new Error("no blocks")

		editor.context.review.tracking = true
		editor.context.identity.author = "Ann"
		editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, between.pos + 1)))
		// As a paste puts it in, at the start of the paragraph after the table.
		editor.dispatch(editor.state.tr.replaceSelection(freshSlice(editor.state.doc.slice(table.pos, table.pos + table.size))))
		expect(names(editor)).toEqual(["table", "p:", "table", "p:Between", "p:After"])

		// Inserted by the paste's change: its mark carries the insertion (w:ins in its paragraph mark's w:rPr).
		const separator = editor.state.doc.child(1)
		const pasted = editor.state.doc.child(2).firstChild

		expect((separator.attrs["revision"] as number) & RevisionFlag.Inserted).toBe(RevisionFlag.Inserted)
		expect(paragraphChange(separator)?.inserted).toEqual(pasted === null ? undefined : rowOverrides(pasted).ins)

		const saved = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [editor] })
		const part = saved.bytes === undefined ? undefined : unzipSync(saved.bytes)["word/document.xml"]
		const xml = part === undefined ? "" : new TextDecoder().decode(part)

		expect(xml).toMatch(/<\/w:tbl><w:p><w:pPr><w:rPr><w:ins [^>]*w:author="Ann"/)

		const { tr } = resolveChanges(editor, storyReviewIndex(editor), { all: true }, false)

		if (tr !== null) editor.dispatch(tr)

		expect(names(editor)).toEqual(["table", "p:Between", "p:After"])
	})

	it("keeps an empty paragraph between two tables once the tracked deletion of the one between them is accepted", async () => {
		const opened = await openBytes(buildDocx({ body: TABLE + p("Gone") + TABLE + p("After") }))
		const editor = new DocxEditorSession(opened.context)
		const [, between] = blocks(editor.state)

		if (between === undefined) throw new Error("no blocks")

		editor.context.review.tracking = true
		editor.context.identity.author = "Ann"
		editor.dispatch(editor.state.tr.delete(between.pos, between.pos + between.size))
		expect(names(editor)).toEqual(["table", "p:Gone", "table", "p:After"])

		// Deliberately not Word's result (Word merges the tables once that mark goes): the paragraph stays, empty and
		// untracked, so the file keeps the two tables these pages show.
		const { tr } = resolveChanges(editor, storyReviewIndex(editor), { all: true }, true)

		if (tr !== null) editor.dispatch(tr)

		expect(names(editor)).toEqual(["table", "p:", "table", "p:After"])
		expect(editor.state.doc.child(1).attrs["revision"]).toBe(0)
		expect(paragraphChange(editor.state.doc.child(1))).toBeUndefined()
	})

	it("gives a story an edit left ending with a table its last paragraph back, and leaves one that did alone", async () => {
		const editor = await session(p("Before") + TABLE + p("Last"))
		const last = blocks(editor.state)[2]

		if (last === undefined) throw new Error("no blocks")

		editor.dispatch(editor.state.tr.delete(last.pos, last.pos + last.size))
		expect(names(editor)).toEqual(["p:Before", "table", "p:"])

		const ending = await session(p("Before") + TABLE)
		const first = blocks(ending.state)[0]

		if (first === undefined) throw new Error("no blocks")

		ending.dispatch(ending.state.tr.insertText("x", first.pos + 1))
		expect(names(ending)).toEqual(["p:xBefore", "table"])
	})
})
