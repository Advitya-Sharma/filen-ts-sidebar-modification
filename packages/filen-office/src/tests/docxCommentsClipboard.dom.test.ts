// @vitest-environment happy-dom
import { TextSelection } from "prosemirror-state"
import { afterEach, beforeEach, describe, expect, it } from "vitest"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { EditSink } from "@filen/office/react/docx/editor"
import { saveDocx } from "@filen/office/core/docx/write"
import { docxComments } from "../core/docx/review/comments"
import { buildDocx, p } from "./docxSynthetic"
import { pagedEditor } from "./pagedDomFixtures"

// A comment and its reply on "beta".
const DOC = buildDocx({
	body:
		'<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>' +
		'<w:r><w:t>beta</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>' +
		'<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>' +
		p("Second paragraph"),
	comments:
		'<w:comment w:id="0" w:author="Bo"><w:p w14:paraId="11111111"><w:r><w:t>First</w:t></w:r></w:p></w:comment>' +
		'<w:comment w:id="1" w:author="Cy"><w:p w14:paraId="22222222"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment>',
	commentsExtended:
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
		'<w15:commentEx w15:paraId="11111111" w15:done="0"/><w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/></w15:commentsEx>'
})

// The span from the first comment anchor to after the last.
function anchorSpan(editor: PagedEditor): { from: number; to: number } {
	let from = -1
	let to = -1

	editor.state.doc.descendants((node, pos) => {
		if (node.type.name === "comment_ref" || (node.type.name === "marker" && String(node.attrs["kind"]).startsWith("comment"))) {
			if (from < 0) from = pos
			to = pos + 1
		}

		return true
	})

	return { from, to }
}

function references(editor: PagedEditor, id: string): number {
	let count = 0

	editor.state.doc.descendants(node => {
		if (node.type.name === "comment_ref" && node.attrs["id"] === id) count++

		return true
	})

	return count
}

function select(sink: EditSink, from: number, to: number): void {
	sink.view.dispatch(sink.view.state.tr.setSelection(TextSelection.create(sink.view.state.doc, from, to)))
}

describe("EditSink: comments on the clipboard", () => {
	let host: HTMLDivElement

	beforeEach(() => {
		host = document.createElement("div")
		document.body.appendChild(host)
	})

	afterEach(() => host.remove())

	it("moves the comments a cut takes when it is pasted back, once", async () => {
		const { editor } = await pagedEditor(DOC)
		const sink = new EditSink(editor, host, { label: "Document", caretClientRect: () => undefined })
		const store = docxComments(editor)
		const { from, to } = anchorSpan(editor)

		select(sink, from, to)

		const content = sink.clipboardContent()

		expect(sink.deleteSelection()).toBe(true)
		expect(store.threads()).toEqual([])

		const end = editor.state.doc.content.size - 1

		select(sink, end, end)
		expect(sink.paste({ text: content?.text ?? "" })).toBe(true)
		expect(store.threads().map(thread => [thread.id, thread.replies.map(reply => reply.id)])).toEqual([["0", ["1"]]])
		expect(store.anchor("0")?.from).toBeGreaterThan(editor.state.doc.child(0).nodeSize)

		// Pasted again it is a copy: no comment twice.
		expect(sink.paste({ text: content?.text ?? "" })).toBe(true)
		expect([references(editor, "0"), references(editor, "1")]).toEqual([1, 1])
		sink.destroy()
	})

	it("moves a comment a cut takes out of a footnote into the body it is pasted in", async () => {
		const { editor, pkg, context } = await pagedEditor(
			buildDocx({
				body:
					'<w:p><w:r><w:t xml:space="preserve">Body </w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
					p("Second paragraph"),
				footnotes:
					'<w:footnote w:id="1"><w:p><w:r><w:t xml:space="preserve">Note one </w:t></w:r><w:commentRangeStart w:id="5"/><w:r><w:t>marked</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r></w:p></w:footnote>',
				comments:
					'<w:comment w:id="5" w:author="Bo" w:date="2020-01-01T00:00:00Z"><w:p><w:r><w:t>On a note</w:t></w:r></w:p></w:comment>'
			})
		)
		const sink = new EditSink(editor, host, { label: "Document", caretClientRect: () => undefined })
		const store = docxComments(editor)
		const note = context.model.stories.findIndex(story => story.kind === "footnote" && story.id === "1")

		expect(editor.enterStory(note, 0)).toBe(true)

		const { from, to } = anchorSpan(editor)

		select(sink, from, to)

		const content = sink.clipboardContent()

		expect(sink.deleteSelection()).toBe(true)
		expect(store.threads()).toEqual([])
		editor.exitStory()

		const end = editor.state.doc.content.size - 1

		select(sink, end, end)
		expect(sink.paste({ text: content?.text ?? "" })).toBe(true)
		expect(references(editor, "5")).toBe(1)
		expect(store.threads().map(thread => thread.id)).toEqual(["5"])
		expect((await saveDocx({ pkg, context, sessions: editor.sessions })).bytes).toBeDefined()
		sink.destroy()
	})

	it("pastes a copy without the comments it holds", async () => {
		const { editor } = await pagedEditor(DOC)
		const sink = new EditSink(editor, host, { label: "Document", caretClientRect: () => undefined })
		const { from, to } = anchorSpan(editor)

		select(sink, from, to)

		const content = sink.clipboardContent()
		const end = editor.state.doc.content.size - 1

		select(sink, end, end)
		expect(sink.paste({ text: content?.text ?? "" })).toBe(true)
		expect(editor.state.doc.child(1).textContent).toBe("Second paragraphbeta")
		expect([references(editor, "0"), references(editor, "1")]).toEqual([1, 1])
		expect(docxComments(editor).anchor("0")?.from).toBeLessThan(editor.state.doc.child(0).nodeSize)
		sink.destroy()
	})
})
