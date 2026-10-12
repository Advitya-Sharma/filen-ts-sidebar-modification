import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { OpcPackage, bytesSource } from "@filen/office/core/opc"
import type { DocxPackage } from "@filen/office/core/docx"
import { DocxEditorSession, type DocxEditContext, type DocxEditorCommand } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { DocxSaveError, saveDocx } from "@filen/office/core/docx/write"
import { markInternal } from "../core/docx/edit/invariants/index"
import { docSectionChanges } from "../core/docx/stories/sections"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"

const decoder = new TextDecoder()
const MARGINS = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const SECTION = (references: string): string => `<w:sectPr>${references}<w:pgSz w:w="12240" w:h="15840"/>${MARGINS}</w:sectPr>`
const HEADER = (type: string, id: string): string => `<w:headerReference w:type="${type}" r:id="${id}"/>`
// Two sections, each with a page of its own; the first refers to header rIdH, the second to none (it shows the
// first's).
const TWO_SECTIONS = buildDocx({
	body: p("First section") + `<w:p><w:pPr>${SECTION(HEADER("default", "rIdH"))}</w:pPr></w:p>` + p("Second section") + SECTION(""),
	headers: { rIdH: p("Shared header") }
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

async function saved({ editor, pkg, context }: Opened): Promise<Uint8Array> {
	await settleAll(editor)

	const result = await saveDocx({ pkg, context, sessions: editor.sessions })

	if (result.bytes === undefined) throw new Error("No bytes")

	expect(pageDigest(editor.layout)).toEqual(await savedDigest(result.bytes))

	return result.bytes
}

// The text of each page's header (region 1) as laid out.
async function headerTexts(editor: PagedEditor): Promise<string[]> {
	await settleAll(editor)

	return pageDigest(editor.layout).map(page =>
		page
			.split("\n")
			.filter(line => / 257 1 /.test(line))
			.map(line => /"([^"]*)"/.exec(line)?.[1] ?? "")
			.join("|")
	)
}

describe("docx headers and footers", () => {
	it("makes a header where the document has none, in Word's Header style, and enters it", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Body text") }))
		const { editor } = opened

		caret(editor, 3)
		expect(editor.snapshot().enabled["header.edit"]).toBe(true)
		expect(await run(editor, { type: "header.edit" })).toBe(true)
		expect(editor.session.story.story.kind).toBe("header")
		expect(editor.snapshot().context.story).toMatchObject({
			kind: "header",
			header: {
				kind: "header",
				section: 0,
				sections: 1,
				type: "default",
				differentFirst: false,
				differentOddEven: false,
				linkedToPrevious: false,
				canLink: false
			}
		})
		type(editor, "New header")

		const bytes = await saved(opened)

		expect(await read(bytes, "word/document.xml")).toMatch(
			/<w:sectPr><w:headerReference w:type="default" r:id="(rId\d+)"\/><\/w:sectPr>/
		)
		expect(await read(bytes, "word/header1.xml")).toContain(
			'<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:t>New header</w:t></w:r></w:p></w:hdr>'
		)
		expect(await read(bytes, "word/styles.xml")).toContain('w:styleId="Header"')
		expect(await headerTexts(editor)).toEqual(["New header"])

		// Undone in the body's history: the header part is not saved.
		editor.exitStory()
		expect(await run(editor, { type: "history.undo" })).toBe(true)
		expect(await run(editor, { type: "history.undo" })).toBe(true)

		const undone = await saved(opened)

		expect((await openBytes(undone)).pkg.has("word/header1.xml")).toBe(false)
		editor.dispose()
	})

	// A header is counted from its start on every page: an edit of a numbered paragraph in it numbers it again
	// from there, never on from the count of its layout before the edit.
	it("numbers a header's list the same on every page after its paragraphs are edited", async () => {
		const opened = await pagedEditor(
			buildDocx({
				body:
					p("First section") +
					`<w:p><w:pPr>${SECTION(HEADER("default", "rIdH"))}</w:pPr></w:p>` +
					p("Second section") +
					SECTION(""),
				headers: { rIdH: p("One") + p("Two") }
			})
		)
		const { editor } = opened

		caret(editor, 3)
		expect(await run(editor, { type: "header.edit" })).toBe(true)

		const session = editor.session

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, 1, session.state.doc.content.size - 1)))
		expect(await run(editor, { type: "list.numbered" })).toBe(true)
		await settleAll(editor)
		// The second paragraph alone written again: the first keeps its number, the second still counts it.
		caret(editor, session.state.doc.content.size - 1)
		type(editor, " more")
		await saved(opened)
		expect(await headerTexts(editor)).toEqual(["1.\\tOne|2.\\tTwo more", "1.\\tOne|2.\\tTwo more"])
		editor.dispose()
	})

	// An empty w:hdr (no paragraph, as some writers leave it) gets an empty paragraph to edit in: untouched, the
	// header is saved and laid out as it was.
	it("keeps an empty header part as it is until its paragraph is edited", async () => {
		const bytes = buildDocx({ body: p("Body text") + SECTION(HEADER("default", "rIdH")), headers: { rIdH: "" } })
		const document = await openBytes(bytes)
		const sessions = document.model.stories.map(story => new DocxEditorSession(document.context, { story }))
		const unedited = await saveDocx({ pkg: document.pkg, context: document.context, sessions })

		expect(unedited.changedParts).toEqual([])

		const opened = await pagedEditor(bytes)
		const { editor } = opened

		caret(editor, 3)
		expect(await run(editor, { type: "header.edit" })).toBe(true)
		expect(editor.session.state.doc.childCount).toBe(1)
		editor.exitStory()
		type(editor, "More ")
		expect(await read(await saved(opened), "word/header1.xml")).toMatch(/<w:hdr [^>]*><\/w:hdr>$/)

		caret(editor, 3)
		expect(await run(editor, { type: "header.edit" })).toBe(true)
		type(editor, "Typed")
		expect(await read(await saved(opened), "word/header1.xml")).toContain("<w:t>Typed</w:t>")
		editor.dispose()
	})

	it("links a section's header to the previous section's and unlinks it with a copy", async () => {
		const opened = await pagedEditor(TWO_SECTIONS)
		const { editor } = opened

		// The second section shows the first's header: linked.
		caret(editor, 20)
		expect(await run(editor, { type: "header.edit" })).toBe(true)
		expect(editor.snapshot().context.story.header).toMatchObject({ section: 1, linkedToPrevious: true, canLink: true })

		// Unlinked: its own header, a copy, edited apart.
		expect(await run(editor, { type: "header.linkToPrevious" })).toBe(true)
		expect(editor.snapshot().context.story.header).toMatchObject({ section: 1, linkedToPrevious: false })
		type(editor, "Own ")
		expect(await headerTexts(editor)).toEqual(["Shared header", "Own Shared header"])

		const bytes = await saved(opened)
		const document = await read(bytes, "word/document.xml")

		expect(document).toMatch(/<w:sectPr><w:headerReference w:type="default" r:id="rId\d+"\/><w:pgSz/)
		expect(await read(bytes, "word/header2.xml")).toContain("<w:t>Own Shared header</w:t>")

		// Linked again: the copy goes with the reference.
		expect(await run(editor, { type: "header.linkToPrevious" })).toBe(true)
		expect(editor.snapshot().context.story.header).toMatchObject({ linkedToPrevious: true })
		expect(await headerTexts(editor)).toEqual(["Shared header", "Shared header"])

		const linked = await saved(opened)

		expect((await openBytes(linked)).pkg.has("word/header2.xml")).toBe(false)
		editor.dispose()
	})

	it("shows a different first page and different even pages", async () => {
		const opened = await pagedEditor(TWO_SECTIONS)
		const { editor } = opened

		caret(editor, 3)
		expect(await run(editor, { type: "header.edit" })).toBe(true)
		expect(await run(editor, { type: "header.differentFirst" })).toBe(true)
		expect(editor.snapshot().context.story.header?.differentFirst).toBe(true)
		// The first page now shows the first-page header, which the section does not have: none.
		expect(await headerTexts(editor)).toEqual(["", "Shared header"])

		editor.exitStory()
		caret(editor, 3)
		expect(await run(editor, { type: "header.edit" })).toBe(true)
		expect(editor.snapshot().context.story.header).toMatchObject({ type: "first" })
		type(editor, "Title page")
		expect(await headerTexts(editor)).toEqual(["Title page", "Shared header"])

		expect(await run(editor, { type: "header.differentOddEven" })).toBe(true)
		expect(editor.snapshot().context.story.header?.differentOddEven).toBe(true)

		const bytes = await saved(opened)

		expect(await read(bytes, "word/settings.xml")).toContain("<w:evenAndOddHeaders/>")
		expect(await read(bytes, "word/document.xml")).toMatch(/<w:headerReference w:type="first" r:id="rId\d+"\/>.*<w:titlePg\/>/)
		// The second page is even: it shows the even header, which nothing has.
		expect(await headerTexts(editor)).toEqual(["Title page", ""])
		expect(docSectionChanges(editor.body.session.state.doc).evenAndOddHeaders).toBe(true)
		editor.dispose()
	})

	it("removes a header from every section showing it", async () => {
		const opened = await pagedEditor(TWO_SECTIONS)
		const { editor } = opened

		caret(editor, 20)
		expect(editor.snapshot().enabled["header.remove"]).toBe(true)
		expect(editor.snapshot().enabled["footer.remove"]).toBe(false)
		expect(await run(editor, { type: "header.remove" })).toBe(true)
		expect(await headerTexts(editor)).toEqual(["", ""])

		const bytes = await saved(opened)

		expect(await read(bytes, "word/document.xml")).not.toContain("headerReference")
		editor.dispose()
	})

	it("tracks what is typed in a header while changes are tracked", async () => {
		const opened = await pagedEditor(TWO_SECTIONS)
		const { editor, context } = opened

		context.identity.author = "Tester"
		context.review.tracking = true
		caret(editor, 3)
		expect(await run(editor, { type: "header.edit" })).toBe(true)
		type(editor, "Tracked ")
		expect(editor.snapshot().review.revisionCount).toBe(1)

		const bytes = await saved(opened)

		expect(await read(bytes, "word/header1.xml")).toMatch(
			/<w:ins w:id="\d+" w:author="Tester" w:date="[^"]+"><w:r><w:t xml:space="preserve">Tracked <\/w:t><\/w:r><\/w:ins>/
		)
		expect(await run(editor, { type: "review.rejectAll" })).toBe(true)
		expect(await read(await saved(opened), "word/header1.xml")).not.toContain("Tracked")
		editor.dispose()
	})

	it("refuses a section reference to a part that is no header", async () => {
		const opened = await pagedEditor(TWO_SECTIONS)
		const { editor, pkg, context } = opened
		const body = editor.body.session
		const key = body.state.doc.attrs["sectPr"] as number

		body.dispatch(
			markInternal(
				body.state.tr.setDocAttribute("sectionChanges", {
					final: { section: key, base: key, footers: { default: body.story.story.index } }
				})
			)
		)

		const error = await saveDocx({ pkg, context, sessions: editor.sessions }).catch((caught: unknown) => caught)

		expect(error).toBeInstanceOf(DocxSaveError)
		editor.dispose()
	})
})
