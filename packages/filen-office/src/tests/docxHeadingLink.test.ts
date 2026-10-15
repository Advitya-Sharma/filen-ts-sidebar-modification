import type { Node as ProseMirrorNode } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { docxSchema, parseDocxEditorCommand, type DocxEditContext } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { hitObject, type PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxPackage } from "@filen/office/core/docx"
import { sessionBookmarks } from "../core/docx/fields/bookmarks"
import { STYLES_PART } from "./docxEditFixtures"
import { buildDocx, p, type SyntheticDocx } from "./docxSynthetic"
import { pagedEditor, settleAll } from "./pagedFixtures"

// A link to a heading goes, as in Word, to a hidden _Toc bookmark around the heading (its own, or one made in
// the same undo step); a press on a note reference is found for Go to note.

const decoder = new TextDecoder()
const marks = docxSchema.marks

interface Opened {
	readonly editor: PagedEditor
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
}

const STYLES = STYLES_PART(
	'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
		'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
		'<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>'
)

function spec(body: string, extra: Partial<SyntheticDocx> = {}): SyntheticDocx {
	return { body, parts: [STYLES], relationships: [{ id: "rIdS", type: "styles", target: "styles.xml" }], ...extra }
}

async function open(source: SyntheticDocx | Uint8Array): Promise<Opened> {
	const opened = await pagedEditor(source instanceof Uint8Array ? source : buildDocx(source))

	await settleAll(opened.editor)

	return opened
}

async function save(opened: Opened): Promise<Uint8Array> {
	const bytes = (await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: opened.editor.sessions })).bytes

	if (bytes === undefined) throw new Error("not saved")

	return bytes
}

async function documentXml(bytes: Uint8Array): Promise<string> {
	return decoder.decode(await (await OpcPackage.open(bytesSource(bytes))).read("word/document.xml"))
}

function find(doc: ProseMirrorNode, text: string): number {
	let found = -1

	doc.descendants((node, pos) => {
		if (found < 0 && node.isText && (node.text ?? "").includes(text)) found = pos + (node.text ?? "").indexOf(text)

		return found < 0
	})

	return found
}

function caret(editor: PagedEditor, pos: number): void {
	const session = editor.session

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, pos)))
}

function link(editor: PagedEditor, heading: number): boolean {
	const command = parseDocxEditorCommand({ type: "link.set", heading })

	if (command === undefined) throw new Error("not parsed")

	return editor.execute(command) === true
}

function linkAnchor(doc: ProseMirrorNode, pos: number): unknown {
	return doc.nodeAt(pos)?.marks.find(mark => mark.type === marks.link)?.attrs["anchor"]
}

const HEADING = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Results</w:t></w:r></w:p>'
const BODY = p("See ") + HEADING + p("Text") + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>'

describe("link.set to a heading", () => {
	it("parses one target only, and keeps hidden names out of bookmark.add", () => {
		expect(parseDocxEditorCommand({ type: "link.set", heading: 4 })).toEqual({ type: "link.set", heading: 4 })
		expect(parseDocxEditorCommand({ type: "link.set", heading: 4, anchor: "A" })).toBeUndefined()
		expect(parseDocxEditorCommand({ type: "link.set", heading: -1 })).toBeUndefined()
		expect(parseDocxEditorCommand({ type: "link.set", heading: "4" })).toBeUndefined()
		expect(parseDocxEditorCommand({ type: "bookmark.add", name: "_Toc123456789" })).toBeUndefined()
	})

	it("links through a new hidden _Toc bookmark around the heading, one undo step, saved and reopened", async () => {
		const opened = await open(spec(BODY))
		const editor = opened.editor
		const session = editor.session
		const heading = session.state.doc.child(0).nodeSize

		caret(editor, find(session.state.doc, "See ") + 4)
		expect(link(editor, heading)).toBe(true)

		const [bookmark] = sessionBookmarks(session).list

		expect(bookmark?.name).toMatch(/^_Toc\d{9}$/)
		expect(bookmark?.hidden).toBe(true)
		expect(session.state.doc.textBetween(bookmark?.from ?? 0, bookmark?.to ?? 0)).toBe("Results")
		// The link shows the heading's text.
		expect(session.state.doc.child(0).textContent).toBe("See Results")
		expect(linkAnchor(session.state.doc, find(session.state.doc, "Results"))).toBe(bookmark?.name)

		const xml = await documentXml(await save(opened))

		expect(xml).toContain(`<w:hyperlink w:anchor="${bookmark?.name ?? ""}"`)
		expect(xml).toMatch(
			new RegExp(`<w:bookmarkStart w:id="\\d+" w:name="${bookmark?.name ?? ""}"/><w:r><w:t>Results</w:t></w:r><w:bookmarkEnd`)
		)

		const reopened = await open(await save(opened))

		expect(sessionBookmarks(reopened.editor.session).list.map(entry => entry.name)).toEqual([bookmark?.name])

		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(sessionBookmarks(session).list).toEqual([])
		expect(session.state.doc.child(0).textContent).toBe("See ")
	})

	it("reuses the heading's own _Toc bookmark", async () => {
		const opened = await open(
			spec(
				p("See ") +
					'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="7" w:name="_Toc000000042"/><w:r><w:t>Results</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>'
			)
		)
		const editor = opened.editor
		const session = editor.session

		caret(editor, find(session.state.doc, "See ") + 4)
		expect(link(editor, session.state.doc.child(0).nodeSize)).toBe(true)
		expect(sessionBookmarks(session).list.map(entry => entry.name)).toEqual(["_Toc000000042"])
		expect(linkAnchor(session.state.doc, find(session.state.doc, "Results"))).toBe("_Toc000000042")
		expect(await documentXml(await save(opened))).toContain('<w:hyperlink w:anchor="_Toc000000042"')
	})

	it("refuses a position that is no paragraph, and links while changes are tracked", async () => {
		const opened = await open(spec(BODY))
		const editor = opened.editor
		const session = editor.session

		caret(editor, find(session.state.doc, "See ") + 4)
		expect(link(editor, 1)).toBe(false)
		// An empty heading: no text to show, so nothing is linked at the caret (a selection or a text would be).
		expect(
			link(editor, session.state.doc.child(0).nodeSize + session.state.doc.child(1).nodeSize + session.state.doc.child(2).nodeSize)
		).toBe(false)
		expect(sessionBookmarks(session).list).toHaveLength(0)

		opened.context.identity.author = "Ann Lee"
		opened.context.review.tracking = true
		expect(link(editor, session.state.doc.child(0).nodeSize)).toBe(true)
		expect(sessionBookmarks(session).list).toHaveLength(1)

		const xml = await documentXml(await save(opened))

		expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d{9}"\/>/)
		expect(xml).toContain('<w:hyperlink w:anchor="_Toc')
	})
})

describe("note reference hits", () => {
	it("finds a footnote reference under a press, and nothing on the text beside it", async () => {
		const opened = await open(
			spec(
				'<w:p><w:r><w:t xml:space="preserve">Body text </w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>',
				{ footnotes: '<w:footnote w:id="1"><w:p><w:r><w:t>Note one</w:t></w:r></w:p></w:footnote>' }
			)
		)
		const editor = opened.editor
		const doc = editor.session.state.doc
		let ref = -1

		doc.descendants((node, pos) => {
			if (node.type === docxSchema.nodes.note_ref) ref = pos

			return true
		})

		const rect = editor.rangeRects(ref, ref + 1, 0)[0]

		expect(rect).toBeDefined()
		expect(hitObject(editor, 0, (rect?.x ?? 0) + (rect?.width ?? 0) / 2, (rect?.y ?? 0) + (rect?.height ?? 0) / 2)).toMatchObject({
			kind: "noteRef",
			pos: ref,
			story: editor.activeStory
		})

		const body = find(doc, "Body")
		const word = editor.rangeRects(body, body + 4, 0)[0]

		expect(hitObject(editor, 0, (word?.x ?? 0) + 2, (word?.y ?? 0) + (word?.height ?? 0) / 2)).toBeUndefined()

		// The caret before it is at the reference: Go to note can run.
		caret(editor, ref)
		expect(editor.snapshot().enabled["note.goto"]).toBe(true)
	})
})

describe("note mark hits", () => {
	it("finds a note's own mark at its start while the note is edited, and Go to note leaves for its reference", async () => {
		const opened = await open(
			spec(
				'<w:p><w:r><w:t xml:space="preserve">Body text </w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>',
				{
					footnotes:
						'<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note one</w:t></w:r></w:p></w:footnote>'
				}
			)
		)
		const editor = opened.editor
		const story = editor.edits.model.stories.find(entry => entry?.kind === "footnote" && entry.id === "1")

		if (story === undefined) throw new Error("no note story")
		expect(editor.enterStory(story.index, 0)).toBe(true)
		await settleAll(editor)

		let mark = -1

		editor.session.state.doc.descendants((node, pos) => {
			if (node.type === docxSchema.nodes.run_object && node.attrs["kind"] === "FootnoteMark") mark = pos

			return mark < 0
		})

		const rect = editor.rangeRects(mark, mark + 1, 0)[0]

		expect(rect).toBeDefined()
		expect(hitObject(editor, 0, (rect?.x ?? 0) + (rect?.width ?? 0) / 2, (rect?.y ?? 0) + (rect?.height ?? 0) / 2)).toMatchObject({
			kind: "noteRef",
			pos: mark,
			story: story.index
		})

		caret(editor, mark)
		expect(editor.snapshot().enabled["note.goto"]).toBe(true)
		expect(await editor.execute({ type: "note.goto" })).toBe(true)
		expect(editor.activeStory).toBe(editor.body.session.story.story.index)
	})
})

describe("insert.date", () => {
	function texts(doc: ProseMirrorNode): { text: string; fields: number } {
		let text = ""
		let fields = 0

		doc.descendants(node => {
			if (node.isText) text += node.text ?? ""
			if (node.type === docxSchema.nodes.field) fields++

			return true
		})

		return { text, fields }
	}

	for (const tracked of [false, true]) {
		it(`inserts the date as text by default in the dialog's sense, or as a DATE field (${tracked ? "tracked" : "untracked"})`, async () => {
			const opened = await open(spec(p("Today: ")))
			const editor = opened.editor
			const session = editor.session

			opened.context.identity.now = () => Date.UTC(2026, 9, 12, 14, 5, 9)

			if (tracked) {
				opened.context.identity.author = "Ann Lee"
				opened.context.review.tracking = true
			}

			caret(editor, find(session.state.doc, "Today: ") + 7)
			expect(parseDocxEditorCommand({ type: "insert.date", format: "d MMMM yyyy", field: false })).toEqual({
				type: "insert.date",
				format: "d MMMM yyyy",
				field: false
			})
			expect(parseDocxEditorCommand({ type: "insert.date", format: "d MMMM yyyy", field: "no" })).toBeUndefined()
			expect(editor.execute({ type: "insert.date", format: "d MMMM yyyy", field: false })).toBe(true)
			expect(texts(session.state.doc)).toEqual({ text: "Today: 12 October 2026", fields: 0 })
			expect(editor.execute({ type: "insert.date", format: "M/d/yyyy", field: true })).toBe(true)
			expect(texts(session.state.doc).fields).toBe(1)

			const xml = await documentXml(await save(opened))

			expect(xml).toContain("12 October 2026")
			expect(xml).toContain('<w:instrText xml:space="preserve"> DATE \\@ "M/d/yyyy" </w:instrText>')
			expect(/<w:ins w:id="\d+" w:author="Ann Lee"[^>]*><w:r><w:t>12 October 2026<\/w:t><\/w:r>/.test(xml)).toBe(tracked)

			const reopened = await open(await save(opened))

			expect(texts(reopened.editor.session.state.doc)).toMatchObject({
				text: expect.stringContaining("Today: 12 October 2026") as string,
				fields: 1
			})
		})
	}
})
