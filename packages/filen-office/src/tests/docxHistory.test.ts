import { closeHistory } from "prosemirror-history"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { saveDocx } from "@filen/office/core/docx/write"
import { buildDocx } from "./docxSynthetic"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"

// One undo order across the stories the editor edits: undo takes back the most recent edit wherever it was,
// entering that story first.

const para = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const SECTION =
	'<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

async function headerDocument(): Promise<Awaited<ReturnType<typeof pagedEditor>> & { header: number }> {
	const opened = await pagedEditor(
		buildDocx({
			body: Array.from({ length: 150 }, (_, k) => para(`Paragraph ${String(k)} with some words to fill the page`)).join("") + SECTION,
			headers: { rIdH1: para("Header") }
		})
	)
	const header = opened.editor.edits.model.stories.find(story => story.kind === "header")?.index

	if (header === undefined) throw new Error("no header")

	return { ...opened, header }
}

// Types `text` at the start of the story being edited, as one history event of its own.
function type(editor: PagedEditor, text: string): void {
	const state = editor.state
	const start = TextSelection.atStart(state.doc).from

	expect(editor.dispatch(closeHistory(state.tr.insertText(text, start)))).toBe(true)
}

function text(editor: PagedEditor, story: number): string {
	return editor.storySession(story)?.state.doc.textContent ?? ""
}

describe("docx undo order across stories", () => {
	it("undoes and redoes in the order of edits, entering each story", async () => {
		const { editor, header, pkg } = await headerDocument()
		const body = editor.body.session.story.story.index

		type(editor, "A")
		expect(editor.enterStory(header, 0)).toBe(true)
		type(editor, "H")
		editor.exitStory()
		type(editor, "B")

		expect(text(editor, body)).toMatch(/^BAParagraph 0/)
		expect(text(editor, header)).toBe("HHeader")
		expect(editor.snapshot()).toMatchObject({ canUndo: true, canRedo: false })

		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(editor.activeStory).toBe(body)
		expect(text(editor, body)).toMatch(/^AParagraph 0/)

		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(editor.activeStory).toBe(header)
		expect(text(editor, header)).toBe("Header")

		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(editor.activeStory).toBe(body)
		expect(text(editor, body)).toMatch(/^Paragraph 0/)
		expect(editor.snapshot()).toMatchObject({ canUndo: false, canRedo: true })
		expect(editor.execute({ type: "history.undo" })).toBe(false)

		expect(editor.execute({ type: "history.redo" })).toBe(true)
		expect(editor.execute({ type: "history.redo" })).toBe(true)
		expect(editor.activeStory).toBe(header)
		expect(text(editor, header)).toBe("HHeader")
		expect(editor.execute({ type: "history.redo" })).toBe(true)
		expect(editor.activeStory).toBe(body)
		expect(text(editor, body)).toMatch(/^BAParagraph 0/)
		expect(editor.execute({ type: "history.redo" })).toBe(false)

		// The layout follows every step, the header's on every page.
		await settleAll(editor)

		const saved = await saveDocx({ pkg, context: editor.body.session.context, sessions: editor.sessions })

		if (saved.bytes === undefined) throw new Error("no bytes")

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved.bytes))
		editor.dispose()
	})

	it("drops the redo order with a new edit and follows a story undoing itself", async () => {
		const { editor, header } = await headerDocument()
		const body = editor.body.session.story.story.index

		type(editor, "A")
		editor.enterStory(header, 0)
		type(editor, "H")
		editor.exitStory()

		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(editor.snapshot().canRedo).toBe(true)

		// A new edit in another story: nothing to redo any more.
		editor.exitStory()
		type(editor, "C")
		expect(editor.snapshot().canRedo).toBe(false)
		expect(editor.execute({ type: "history.redo" })).toBe(false)

		// The header's session undoes its own edit (its keys): the order follows.
		editor.enterStory(header, 0)
		type(editor, "K")
		editor.exitStory()
		expect(editor.storySession(header)?.execute({ type: "history.undo" })).toBe(true)
		expect(text(editor, header)).toBe("Header")
		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(text(editor, body)).toMatch(/^AParagraph 0/)
		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(text(editor, body)).toMatch(/^Paragraph 0/)
		expect(editor.execute({ type: "history.redo" })).toBe(true)
		expect(editor.execute({ type: "history.redo" })).toBe(true)
		expect(text(editor, body)).toMatch(/^CAParagraph 0/)
		expect(editor.execute({ type: "history.redo" })).toBe(true)
		expect(text(editor, header)).toBe("KHeader")
		editor.dispose()
	})

	it("undoes in a story no page shows without entering it, and enters others on the page last edited", async () => {
		const { editor, header } = await headerDocument()
		const body = editor.body.session.story.story.index
		const note = editor.addStory("footnote")

		if (note === undefined) throw new Error("not added")

		type(editor, "A")
		expect(note.dispatch(closeHistory(note.state.tr.insertText("note", 1)))).toBe(true)
		expect(editor.execute({ type: "history.undo" })).toBe(true)
		expect(editor.activeStory).toBe(body)
		expect(note.state.doc.textContent).toBe("")

		// A header edited on page 2: entered again without a page, its caret is on page 2.
		expect(editor.enterStory(header, 2)).toBe(true)
		type(editor, "H")
		await settleAll(editor)
		expect(editor.caret()?.page).toBe(2)
		editor.exitStory()
		expect(editor.enterStory(header, -1)).toBe(true)
		expect(editor.caret()?.page).toBe(2)
		editor.dispose()
	})

	it("refuses undo in read-only mode", async () => {
		const bytes = buildDocx({ body: para("One"), settings: '<w:documentProtection w:edit="readOnly" w:enforcement="1"/>' })
		const { editor } = await pagedEditor(bytes)

		expect(editor.session.readOnly).toBe(true)
		expect(editor.snapshot().enabled["history.undo"]).toBe(false)
		expect(editor.execute({ type: "history.undo" })).toBe(false)
		editor.dispose()
	})
})
