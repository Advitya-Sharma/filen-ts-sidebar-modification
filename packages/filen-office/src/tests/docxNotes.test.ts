import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { OpcPackage, bytesSource } from "@filen/office/core/opc"
import type { DocxPackage } from "@filen/office/core/docx"
import { docxSchema, type DocxEditContext, type DocxEditorCommand } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { DocxSaveError, saveDocx } from "@filen/office/core/docx/write"
import { markInternal } from "../core/docx/edit/invariants/index"
import { noteReferences } from "../core/docx/stories/notes"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"

const decoder = new TextDecoder()
const nodes = docxSchema.nodes

const SEPARATORS = (kind: string): string =>
	`<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
	`<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`
const NOTE = (kind: string, id: string, text: string): string =>
	`<w:${kind} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${kind}>`
const REFERENCE = (kind: string, id: string): string =>
	`<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>`

// A body with two footnotes, the first referred to in the first paragraph, the second in the second.
const WITH_NOTES = buildDocx({
	body:
		`<w:p><w:r><w:t>One</w:t></w:r>${REFERENCE("footnote", "1")}<w:r><w:t xml:space="preserve"> two</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Three</w:t></w:r>${REFERENCE("footnote", "2")}</w:p>`,
	footnotes: SEPARATORS("footnote") + NOTE("footnote", "1", "first") + NOTE("footnote", "2", "second")
})

async function read(bytes: Uint8Array, name: string): Promise<string> {
	return decoder.decode(await (await OpcPackage.open(bytesSource(bytes))).read(name))
}

interface Opened {
	readonly editor: PagedEditor
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
}

function caret(editor: PagedEditor, pos: number): void {
	const session = editor.session

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, pos)))
}

async function run(editor: PagedEditor, command: DocxEditorCommand): Promise<boolean> {
	return await editor.execute(command)
}

function type(editor: PagedEditor, text: string): void {
	editor.session.dispatch(editor.session.state.tr.insertText(text))
}

// The saved file, after checking that the pages shown are a fresh layout of it.
async function saved({ editor, pkg, context }: Opened): Promise<Uint8Array> {
	await settleAll(editor)

	const result = await saveDocx({ pkg, context, sessions: editor.sessions })

	if (result.bytes === undefined) throw new Error("No bytes")

	expect(pageDigest(editor.layout)).toEqual(await savedDigest(result.bytes))

	return result.bytes
}

async function notesOf(bytes: Uint8Array, kind: "footnote" | "endnote"): Promise<Record<string, string>> {
	const opened = await openBytes(bytes)
	const out: Record<string, string> = {}

	for (const [id, story] of kind === "footnote" ? opened.model.footnotes : opened.model.endnotes) {
		if (story.noteType !== "normal") continue

		const part = opened.model.parts[story.part]

		out[id] = part === undefined ? "" : decoder.decode(part.document.bytes.subarray(part.start(story.item), part.byteEnd(story.item)))
	}

	return out
}

describe("docx footnotes and endnotes", () => {
	it("inserts a footnote where the document has no notes part, with Word's styles and separators", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Alpha beta") + p("Second") }))
		const { editor } = opened

		caret(editor, 6)
		expect(editor.snapshot().enabled["note.insertFootnote"]).toBe(true)
		expect(await run(editor, { type: "note.insertFootnote" })).toBe(true)
		expect(editor.session.story.story.kind).toBe("footnote")
		expect(editor.snapshot().context.story.note).toEqual({ kind: "footnote", id: "1" })
		type(editor, "the note")

		const bytes = await saved(opened)
		const document = await read(bytes, "word/document.xml")
		const footnotes = await read(bytes, "word/footnotes.xml")
		const styles = await read(bytes, "word/styles.xml")

		expect(document).toContain('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>')
		expect(footnotes).toContain('<w:footnote w:type="separator" w:id="-1">')
		expect(footnotes).toContain(
			'<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> the note</w:t></w:r></w:p></w:footnote>'
		)
		expect(styles).toContain('w:styleId="FootnoteText"')
		expect(styles).toContain('w:styleId="FootnoteReference"')
		expect(Object.keys(await notesOf(bytes, "footnote"))).toEqual(["1"])
		editor.dispose()
	})

	it("inserts an endnote with a custom mark after the selection, numbered after the document's own", async () => {
		const opened = await pagedEditor(WITH_NOTES)
		const { editor } = opened
		const doc = editor.session.state.doc

		editor.session.dispatch(editor.session.state.tr.setSelection(TextSelection.create(doc, 1, 3)))
		expect(await run(editor, { type: "note.insertEndnote", mark: "*" })).toBe(true)
		type(editor, "marked")

		const bytes = await saved(opened)
		const document = await read(bytes, "word/document.xml")

		expect(document).toContain(
			'<w:t>On</w:t></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:customMarkFollows="1" w:id="1"/><w:t>*</w:t></w:r>'
		)
		expect(Object.values(await notesOf(bytes, "endnote"))[0]).toContain("<w:t>*</w:t>")

		// A footnote is numbered after the document's (ids max + 1).
		editor.exitStory()
		caret(editor, 2)
		expect(await run(editor, { type: "note.insertFootnote" })).toBe(true)
		expect(editor.snapshot().context.story.note?.id).toBe("3")
		editor.dispose()
	})

	it("deletes a note with a custom mark together with its mark", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Alpha beta") }))
		const { editor } = opened

		caret(editor, 6)
		expect(await run(editor, { type: "note.insertFootnote", mark: "**" })).toBe(true)
		editor.exitStory()
		caret(editor, 6)
		expect(await run(editor, { type: "note.delete" })).toBe(true)
		expect(editor.session.state.doc.textContent).toBe("Alpha beta")
		expect(await notesOf(await saved(opened), "footnote")).toEqual({})
		editor.dispose()
	})

	it("refuses notes outside the main text, as Word does", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Body"), headers: { rIdH: p("Head") } }))
		const { editor } = opened

		expect(await run(editor, { type: "header.edit" })).toBe(true)
		expect(editor.session.story.story.kind).toBe("header")
		expect(editor.snapshot().enabled["note.insertFootnote"]).toBe(false)
		expect(await run(editor, { type: "note.insertFootnote" })).toBe(false)
		editor.dispose()
	})

	it("moves between a note and its reference, and edits a note in place", async () => {
		const opened = await pagedEditor(WITH_NOTES)
		const { editor } = opened

		caret(editor, 4)
		expect(await run(editor, { type: "note.goto" })).toBe(true)
		expect(editor.session.story.story).toMatchObject({ kind: "footnote", id: "1" })
		type(editor, " edited")
		expect(await run(editor, { type: "note.goto" })).toBe(true)
		expect(editor.session).toBe(editor.body.session)
		expect(editor.session.state.selection.from).toBe(5)

		const bytes = await saved(opened)

		expect((await notesOf(bytes, "footnote"))["1"]).toContain("edited")
		editor.dispose()
	})

	// The note is laid out while its new font is still loading: the page it is on is laid out again once loaded.
	it("lays a note out again with a font it needs that was still loading", async () => {
		const opened = await pagedEditor(WITH_NOTES)
		const { editor } = opened

		caret(editor, 4)
		expect(await run(editor, { type: "note.goto" })).toBe(true)

		const session = editor.session

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, 1, session.state.doc.content.size - 1)))
		expect(await run(editor, { type: "font.family", family: "Calibri" })).toBe(true)
		await saved(opened)
		editor.dispose()
	})

	it("deletes a note with its reference, and keeps it while an undo brings the reference back", async () => {
		const opened = await pagedEditor(WITH_NOTES)
		const { editor } = opened

		caret(editor, 4)
		expect(await run(editor, { type: "note.delete" })).toBe(true)
		expect(Object.keys(await notesOf(await saved(opened), "footnote"))).toEqual(["2"])

		// A plain delete of the reference too.
		const reference = noteReferences(editor.session.state.doc)[0]

		if (reference === undefined) throw new Error("no reference")

		editor.session.dispatch(editor.session.state.tr.delete(reference.pos, reference.pos + 1))

		const none = await saved(opened)

		expect(await notesOf(none, "footnote")).toEqual({})
		// The separators stay.
		expect(await read(none, "word/footnotes.xml")).toContain('w:type="continuationSeparator"')

		await run(editor, { type: "history.undo" })
		await run(editor, { type: "history.undo" })
		expect(Object.keys(await notesOf(await saved(opened), "footnote"))).toEqual(["1", "2"])
		editor.dispose()
	})

	it("converts a note to the other kind, and every note at once", async () => {
		const opened = await pagedEditor(WITH_NOTES)
		const { editor } = opened

		caret(editor, 4)
		expect(await run(editor, { type: "note.convert" })).toBe(true)

		const once = await saved(opened)

		expect(Object.keys(await notesOf(once, "footnote"))).toEqual(["2"])
		expect(Object.values(await notesOf(once, "endnote"))[0]).toContain(" first")
		expect(Object.values(await notesOf(once, "endnote"))[0]).toContain("<w:endnoteRef/>")

		expect(await run(editor, { type: "note.convertAll", to: "endnote" })).toBe(true)

		const all = await saved(opened)

		expect(await notesOf(all, "footnote")).toEqual({})
		expect(Object.keys(await notesOf(all, "endnote"))).toHaveLength(2)

		await run(editor, { type: "history.undo" })
		await run(editor, { type: "history.undo" })
		expect(Object.keys(await notesOf(await saved(opened), "footnote"))).toEqual(["1", "2"])
		editor.dispose()
	})

	it("tracks a note inserted while changes are tracked, and rejecting it takes the note out", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Alpha beta") }))
		const { editor, context } = opened

		context.identity.author = "Tester"
		context.review.tracking = true
		caret(editor, 6)
		expect(await run(editor, { type: "note.insertFootnote" })).toBe(true)
		type(editor, "tracked")

		const bytes = await saved(opened)

		expect(await read(bytes, "word/document.xml")).toMatch(
			/<w:ins w:id="\d+" w:author="Tester" w:date="[^"]+"><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="1"\/>/
		)
		expect(await read(bytes, "word/footnotes.xml")).toContain('w:author="Tester"')

		editor.exitStory()
		expect(await run(editor, { type: "review.rejectAll" })).toBe(true)
		expect(noteReferences(editor.session.state.doc)).toEqual([])

		const rejected = await saved(opened)

		expect(opened.pkg.has("word/footnotes.xml")).toBe(false)
		expect(rejected.length).toBeGreaterThan(0)
		await expect(read(rejected, "word/footnotes.xml")).rejects.toThrow()
		editor.dispose()
	})

	it("numbers notes as the Footnote and Endnote options set, for a section or the document", async () => {
		const opened = await pagedEditor(WITH_NOTES)
		const { editor } = opened

		caret(editor, 2)
		expect(await run(editor, { type: "note.options", kind: "footnote", format: "lowerRoman", start: 3, scope: "document" })).toBe(true)

		const bytes = await saved(opened)

		expect(await read(bytes, "word/document.xml")).toContain(
			'<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/></w:footnotePr>'
		)
		// Laid out: the first note is iii.
		expect(pageDigest(editor.layout).join("\n")).toContain('"iii first"')
		editor.dispose()
	})

	it("refuses a save with two references to one new note", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Alpha beta") }))
		const { editor, pkg, context } = opened

		caret(editor, 3)
		expect(await run(editor, { type: "note.insertFootnote" })).toBe(true)
		editor.exitStory()

		const body = editor.body.session

		body.dispatch(markInternal(body.state.tr.insert(8, nodes.note_ref.create({ kind: "footnote", id: "1" }))))

		const error = await saveDocx({ pkg, context, sessions: editor.sessions }).catch((caught: unknown) => caught)

		expect(error).toBeInstanceOf(DocxSaveError)
		expect((error as DocxSaveError).reason).toBe("duplicateId")
		editor.dispose()
	})
})
