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

	it("leaves out a pasted copy of the reference to a note it made, as copies of the file's references", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Alpha beta") + p("Second") }))
		const { editor } = opened

		caret(editor, 6)
		expect(await run(editor, { type: "note.insertFootnote" })).toBe(true)
		type(editor, "the note")
		editor.exitStory()

		// The first paragraph's text, reference and all, copied into the second as a drag-copy without the sanitiser does.
		const body = editor.body.session
		const first = body.state.doc.child(0)

		body.dispatch(body.state.tr.setSelection(TextSelection.create(body.state.doc, first.nodeSize + 2)))
		expect(body.dispatch(body.state.tr.replaceSelection(body.state.doc.slice(1, first.nodeSize - 1)))).toBe(true)
		expect(body.state.doc.textContent.match(/Alpha beta/g)).toHaveLength(2)
		expect(noteReferences(body.state.doc)).toHaveLength(1)

		const bytes = await saved(opened)

		expect((await read(bytes, "word/document.xml")).match(/<w:footnoteReference /g)).toHaveLength(1)
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

	it("places notes document-wide, in settings, as Word reads w:pos; footnotes beneath the text follow its last line", async () => {
		const opened = await pagedEditor(WITH_NOTES)
		const { editor } = opened
		const top = async (): Promise<number> => {
			await settleAll(editor)

			const line =
				pageDigest(editor.layout)[0]
					?.split("\n")
					.find(entry => entry.includes('"1 first"')) ?? ""

			return Number(/^(-?[\d.]+) /.exec(line)?.[1] ?? Number.NaN)
		}
		const bottom = await top()

		caret(editor, 2)
		expect(await run(editor, { type: "note.options", kind: "footnote", position: "beneathText", scope: "section" })).toBe(true)
		expect(await top()).toBeLessThan(bottom - 200)

		const bytes = await saved(opened)

		expect(await read(bytes, "word/document.xml")).not.toContain("<w:pos ")
		expect(await read(bytes, "word/settings.xml")).toContain('<w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr>')
		// Already so: nothing to change.
		expect(await run(editor, { type: "note.options", kind: "footnote", position: "beneathText", scope: "document" })).toBe(false)
		editor.dispose()

		// Spliced into the settings' own w:footnotePr, before its separator note ids.
		const listed = await pagedEditor(
			buildDocx({
				body: `<w:p><w:r><w:t>One</w:t></w:r>${REFERENCE("footnote", "1")}</w:p>`,
				footnotes: SEPARATORS("footnote") + NOTE("footnote", "1", "first"),
				settings: '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
			})
		)

		caret(listed.editor, 2)
		expect(await run(listed.editor, { type: "note.options", kind: "footnote", position: "beneathText", scope: "document" })).toBe(true)
		expect(await read(await saved(listed), "word/settings.xml")).toContain(
			'<w:footnotePr><w:pos w:val="beneathText"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
		)
		listed.editor.dispose()

		// A section's own w:pos is ignored.
		const ignored = await pagedEditor(
			buildDocx({
				body: `<w:p><w:r><w:t>One</w:t></w:r>${REFERENCE("footnote", "1")}</w:p><w:sectPr><w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr></w:sectPr>`,
				footnotes: SEPARATORS("footnote") + NOTE("footnote", "1", "first")
			})
		)

		await settleAll(ignored.editor)
		expect(pageDigest(ignored.editor.layout)[0]).toBe(
			pageDigest(
				(
					await pagedEditor(
						buildDocx({
							body: `<w:p><w:r><w:t>One</w:t></w:r>${REFERENCE("footnote", "1")}</w:p><w:sectPr/>`,
							footnotes: SEPARATORS("footnote") + NOTE("footnote", "1", "first")
						})
					)
				).editor.layout
			)[0]
		)
		ignored.editor.dispose()
	})

	it("converts a note with a custom mark, its mark's text taking the new kind's reference style", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Alpha beta") }))
		const { editor } = opened

		caret(editor, 3)
		expect(await run(editor, { type: "note.insertFootnote", mark: "*" })).toBe(true)
		editor.session.dispatch(editor.session.state.tr.insertText("star"))
		editor.exitStory()
		caret(editor, 3)
		expect(await run(editor, { type: "note.convert" })).toBe(true)

		const notes = await read(await saved(opened), "word/endnotes.xml")
		const note = notes.slice(notes.lastIndexOf("<w:endnote "))

		expect(note).toContain('<w:rStyle w:val="EndnoteReference"/></w:rPr><w:t>*</w:t>')
		expect(note).not.toContain("FootnoteReference")
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

	// Letter pages with one-inch margins.
	const PAGE =
		'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
	const lines = (count: number, name: string): string => Array.from({ length: count }, (_, k) => p(`${name} ${String(k)}`)).join("")
	const shows = (editor: PagedEditor, text: string): number => pageDigest(editor.layout).findIndex(page => page.includes(text))

	it("shows a footnote with a line taller than the page, and lays out edits to it", async () => {
		const opened = await pagedEditor(
			buildDocx({
				body: `<w:p><w:r><w:t>Giant</w:t></w:r>${REFERENCE("footnote", "1")}</w:p>${lines(5, "after")}${PAGE}`,
				footnotes:
					SEPARATORS("footnote") +
					'<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:rPr><w:sz w:val="1400"/></w:rPr><w:t>i</w:t></w:r></w:p></w:footnote>'
			})
		)
		const { editor } = opened

		// Shown on the page of its reference, overflowing it, rather than left for a page that could hold it (the
		// text after goes on the next page).
		expect(shows(editor, '"1i"')).toBe(0)
		expect(editor.layout.pages.length).toBe(2)

		caret(editor, 6)
		expect(await run(editor, { type: "note.goto" })).toBe(true)
		type(editor, "j")
		await saved(opened)
		expect(shows(editor, '"1ij"')).toBe(0)
		editor.dispose()
	})

	it("lays out the paragraph referring to a note again on the page before when the note shrinks", async () => {
		const words = "alpha beta gamma delta epsilon zeta eta theta ".repeat(3)
		const opened = await pagedEditor(
			buildDocx({
				// The note does not fit under the paragraph referring to it, and splits: the paragraph's last line goes
				// to the next page, and with it (widow control) the paragraph, the note after it.
				body: `${lines(48, "line")}<w:p><w:pPr><w:widowControl/></w:pPr><w:r><w:t xml:space="preserve">refers </w:t></w:r>${REFERENCE("footnote", "1")}<w:r><w:t xml:space="preserve"> ${words}</w:t></w:r></w:p>${lines(10, "after")}${PAGE}`,
				footnotes:
					SEPARATORS("footnote") +
					'<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> small</w:t></w:r></w:p><w:p><w:r><w:rPr><w:sz w:val="96"/></w:rPr><w:t>big</w:t></w:r></w:p></w:footnote>'
			})
		)
		const { editor } = opened
		const story = editor.edits.model.stories.find(entry => entry?.kind === "footnote" && entry.id === "1")

		if (story === undefined) throw new Error("no note")

		expect(shows(editor, '"refers')).toBe(1)
		expect(editor.enterStory(story.index, 1)).toBe(true)

		// Its second paragraph emptied (not the one opening the page): the note fits under the paragraph now.
		const note = editor.session
		const second = note.state.doc.child(0).nodeSize

		note.dispatch(note.state.tr.delete(second + 1, second + 4))
		await saved(opened)
		expect(shows(editor, '"refers')).toBe(0)
		editor.dispose()
	})

	it("lays out a keystroke in a footnote on the pages near it, not to the end of the document", async () => {
		const opened = await pagedEditor(
			buildDocx({
				body: `<w:p><w:r><w:t>Noted</w:t></w:r>${REFERENCE("footnote", "1")}</w:p>${lines(600, "after")}${PAGE}`,
				footnotes: SEPARATORS("footnote") + NOTE("footnote", "1", "first")
			})
		)
		const { editor } = opened
		let laid = 0

		expect(editor.layout.pages.length).toBeGreaterThan(10)
		editor.layout.subscribe(event => {
			if (!event.final) laid += event.to - event.from
		})

		caret(editor, 6)
		expect(await run(editor, { type: "note.goto" })).toBe(true)
		type(editor, " edited")
		await settleAll(editor)

		expect(laid).toBeLessThan(4)
		await saved(opened)
		editor.dispose()
	})
})
