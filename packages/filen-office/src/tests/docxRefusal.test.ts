import { Fragment, Slice } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema, type DocxEditContext, type DocxRefusalReason } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { markInternal } from "../core/docx/edit/invariants/core"
import { openSpec } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { pagedEditor } from "./pagedFixtures"

// A user edit refused because tracking needs the author's name is heard once, with why; undo, redo and the
// editor's own transactions never are.

const NOW = Date.UTC(2026, 9, 13, 9, 30)

function author(context: DocxEditContext, name: string | undefined): void {
	context.identity.author = name
	context.identity.now = () => NOW
}

function heard(target: { onRefused(listener: (reason: DocxRefusalReason) => void): () => void }): {
	readonly reasons: DocxRefusalReason[]
	readonly stop: () => void
} {
	const reasons: DocxRefusalReason[] = []

	return { reasons, stop: target.onRefused(reason => reasons.push(reason)) }
}

async function tracking(body: string, name: string | undefined): Promise<{ editor: PagedEditor; context: DocxEditContext }> {
	const opened = await pagedEditor(buildDocx({ body }))

	author(opened.context, name)
	opened.context.review.tracking = true

	return opened
}

function paste(editor: PagedEditor, text: string, event: "paste" | "drop"): boolean {
	const state = editor.state
	const slice = new Slice(Fragment.from(docxSchema.text(text)), 0, 0)

	return editor.dispatch(state.tr.replaceRange(2, 2, slice).setMeta("uiEvent", event))
}

describe("docx edit refusals", () => {
	it("hears a keystroke refused for the author's name once, and nothing once there is a name", async () => {
		const opened = await openSpec({ body: p("one") })

		author(opened.context, "")
		opened.context.review.tracking = true

		const session = new DocxEditorSession(opened.context)
		const { reasons, stop } = heard(session)

		expect(session.dispatch(session.state.tr.insertText("x", 2))).toBe(false)
		expect(reasons).toEqual(["authorNeeded"])

		// A selection change is no edit.
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, 3)))
		expect(reasons).toHaveLength(1)

		author(opened.context, "Ann")
		expect(session.dispatch(session.state.tr.insertText("x", 2))).toBe(true)
		expect(reasons).toHaveLength(1)

		// Not tracking: nothing to record, nothing refused.
		author(opened.context, undefined)
		opened.context.review.tracking = false
		expect(session.dispatch(session.state.tr.insertText("y", 2))).toBe(true)
		expect(reasons).toHaveLength(1)

		// Unsubscribed: not heard.
		opened.context.review.tracking = true
		stop()
		expect(session.dispatch(session.state.tr.insertText("z", 2))).toBe(false)
		expect(reasons).toHaveLength(1)
		session.dispose()
	})

	it("never hears undo, redo or the session's own transactions", async () => {
		const opened = await openSpec({ body: p("one two") })

		author(opened.context, "Ann")
		opened.context.review.tracking = true

		const session = new DocxEditorSession(opened.context)
		const { reasons } = heard(session)

		expect(session.dispatch(session.state.tr.insertText("x", 2))).toBe(true)
		author(opened.context, undefined)
		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(session.execute({ type: "history.redo" })).toBe(true)
		session.dispatch(markInternal(session.state.tr.insertText("y", 2)))
		expect(reasons).toEqual([])
		session.dispose()
	})

	it("relays each refused typing, paste and drop of any story once through the paged editor", async () => {
		const { editor } = await tracking(p("one") + p("two"), "")
		const { reasons, stop } = heard(editor)
		const before = editor.state.doc

		expect(editor.dispatch(editor.state.tr.insertText("x", 2))).toBe(false)
		expect(reasons).toEqual(["authorNeeded"])
		expect(paste(editor, "pasted", "paste")).toBe(false)
		expect(paste(editor, "dropped", "drop")).toBe(false)
		expect(reasons).toEqual(["authorNeeded", "authorNeeded", "authorNeeded"])
		expect(editor.state.doc).toBe(before)

		// A story the editor makes a session for later (a new footnote's) is heard too.
		const note = editor.addStory("footnote")

		if (note === undefined) throw new Error("no note story")

		expect(note.dispatch(note.state.tr.insertText("n", 1))).toBe(false)
		expect(reasons).toHaveLength(4)

		stop()
		editor.dispatch(editor.state.tr.insertText("x", 2))
		expect(reasons).toHaveLength(4)
		editor.dispose()
	})

	it("hears each refused composition step once, and never the marking that ends a composition", async () => {
		const { editor, context } = await tracking(p("one"), "")
		const { reasons } = heard(editor)

		editor.setComposition({ from: 2, to: 2 })
		expect(editor.dispatch(editor.state.tr.insertText("か", 2).setMeta("composition", 1))).toBe(false)
		expect(editor.dispatch(editor.state.tr.insertText("な", 2).setMeta("composition", 1))).toBe(false)
		editor.setComposition(undefined)
		await Promise.resolve()
		expect(reasons).toEqual(["authorNeeded", "authorNeeded"])

		// Composed with a name, which goes before the composition ends: the marking it ends with is the editor's own.
		author(context, "Ann")
		editor.setComposition({ from: 2, to: 2 })
		expect(editor.dispatch(editor.state.tr.insertText("か", 2).setMeta("composition", 1))).toBe(true)
		author(context, undefined)
		editor.setComposition(undefined)
		await Promise.resolve()
		expect(reasons).toHaveLength(2)
		editor.dispose()
	})

	it("hears every refused transaction, one event each", async () => {
		const { editor, context } = await tracking(p("one"), "Ann")
		const { reasons } = heard(editor)

		for (let k = 0; k < 50; k++) expect(editor.dispatch(editor.state.tr.insertText("x", 2))).toBe(true)

		expect(reasons).toEqual([])
		author(context, "")

		for (let k = 0; k < 50; k++) expect(editor.dispatch(editor.state.tr.insertText("x", 2))).toBe(false)

		// One event per refused transaction, never batched or repeated.
		expect(reasons).toHaveLength(50)
		editor.dispose()
	})
})
