import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { saveDocx } from "@filen/office/core/docx/write"
import { buildDocx, p } from "./docxSynthetic"
import { pagedEditor, settleAll } from "./pagedFixtures"

// A table of contents built as paragraphs stays over paragraphs: a field read back from one paragraph is one
// field atom, so an edit that would put its begin and end in one paragraph is refused, not left to fail the save.

describe("a field spanning paragraphs", () => {
	it("refuses an edit that would put an empty table of contents' begin and end in one paragraph, and saves", async () => {
		const { editor, pkg, context } = await pagedEditor(buildDocx({ body: p("one two") + p("three four") + p("five six") }))
		const session = editor.body.session

		await settleAll(editor)
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, 10)))
		expect(await editor.execute({ type: "toc.insert", levels: 3, hyperlinks: true })).toBe(true)

		const doc = session.state.doc
		const index = [...Array(doc.childCount).keys()].find(k => doc.child(k).textContent.includes("No table")) ?? -1
		let pos = 0

		for (let k = 0; k < index; k++) pos += doc.child(k).nodeSize

		const end = pos + doc.child(index).nodeSize

		// Its paragraph joined with the one holding its end, or a deletion from inside it into that paragraph.
		expect(session.dispatch(session.state.tr.join(end))).toBe(false)
		expect(session.dispatch(session.state.tr.delete(pos + 5, end + 1))).toBe(false)
		expect(session.state.doc).toBe(doc)

		// The paragraph after the end joined to it: the field stays over its paragraphs.
		const after = end + doc.child(index + 1).nodeSize

		expect(session.dispatch(session.state.tr.join(after))).toBe(true)
		await settleAll(editor)

		const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

		expect(saved.bytes).toBeDefined()
		editor.dispose()
	})
})
