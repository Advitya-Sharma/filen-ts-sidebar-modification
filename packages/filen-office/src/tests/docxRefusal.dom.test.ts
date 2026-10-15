// @vitest-environment happy-dom
import { Fragment, Slice } from "prosemirror-model"
import { afterEach, beforeEach, describe, expect, it } from "vitest"
import { docxSchema, type DocxRefusalReason } from "@filen/office/core/docx/edit"
import { EditSink } from "@filen/office/react/docx/editor"
import { buildDocx, p } from "./docxSynthetic"
import { pagedEditor } from "./pagedDomFixtures"

// What the view sends (typing, a paste, a drop, IME composition), refused while tracking waits for the author's
// name: one event per edit, the view put back as it was.

describe("EditSink refused edits", () => {
	let host: HTMLDivElement

	beforeEach(() => {
		host = document.createElement("div")
		document.body.appendChild(host)
	})

	afterEach(() => host.remove())

	it("hears refused typing, paste, drop and composition from the view once each", async () => {
		const { editor, context } = await pagedEditor(buildDocx({ body: p("Alpha beta") }))

		context.identity.author = ""
		context.review.tracking = true

		const reasons: DocxRefusalReason[] = []

		editor.onRefused(reason => reasons.push(reason))

		const sink = new EditSink(editor, host, { label: "Document", caretClientRect: () => undefined })
		const before = editor.state.doc
		const view = sink.view

		view.dispatch(view.state.tr.insertText("x", 2))
		expect(reasons).toHaveLength(1)
		expect(sink.paste({ text: "pasted" })).toBe(true)
		expect(reasons).toHaveLength(2)
		view.dispatch(
			view.state.tr.replaceRange(2, 2, new Slice(Fragment.from(docxSchema.text("dropped")), 0, 0)).setMeta("uiEvent", "drop")
		)
		expect(reasons).toHaveLength(3)
		view.dispatch(view.state.tr.insertText("か", 2).setMeta("composition", 1))
		expect(reasons).toEqual(["authorNeeded", "authorNeeded", "authorNeeded", "authorNeeded"])

		// Nothing was applied, and the view shows the document as it is.
		expect(editor.state.doc).toBe(before)
		expect(view.state.doc.eq(before)).toBe(true)
		await Promise.resolve()
		expect(reasons).toHaveLength(4)
		sink.destroy()
		editor.dispose()
	})
})
