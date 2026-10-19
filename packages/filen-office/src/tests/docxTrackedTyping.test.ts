import { undoDepth } from "prosemirror-history"
import { TextSelection, type EditorState } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { openSpec } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

// Typing while tracking undoes in the runs untracked typing does (Word undoes typing in runs, tracked or not):
// each key's insertion is marked in its own step, so the history joins the keys of a run.

async function session(tracking: boolean): Promise<DocxEditorSession> {
	const opened = await openSpec({ body: p("one two") })

	opened.context.identity.author = "Ann"
	opened.context.identity.now = () => Date.parse("2026-01-01T10:00:00Z")
	opened.context.review.tracking = tracking

	const out = new DocxEditorSession(opened.context)

	out.dispatch(out.state.tr.setSelection(TextSelection.create(out.state.doc, 4)))

	return out
}

// The history's items (steps it keeps to undo), read from its plugin state.
function historyItems(state: EditorState): number {
	const plugin = state.plugins.find(entry => (Reflect.get(entry, "key") as string).startsWith("history$"))
	const history: unknown = plugin?.getState(state)
	const done: unknown = typeof history === "object" && history !== null ? Reflect.get(history, "done") : undefined
	const items: unknown = typeof done === "object" && done !== null ? Reflect.get(done, "items") : undefined
	const length: unknown = typeof items === "object" && items !== null ? Reflect.get(items, "length") : undefined

	return typeof length === "number" ? length : -1
}

describe("tracked typing in the undo history", () => {
	it("is one history step for a run of keys, as untracked typing is, and one undo takes it all", async () => {
		const depths: number[] = []
		const items: number[] = []

		for (const tracking of [false, true]) {
			const typing = await session(tracking)

			for (let k = 0; k < 50; k++) typing.dispatch(typing.state.tr.insertText("x").setTime(1000))

			depths.push(undoDepth(typing.state))
			items.push(historyItems(typing.state))

			const paragraph = typing.state.doc.child(0)

			expect(paragraph.textContent).toBe(`one${"x".repeat(50)} two`)

			// Tracked, every typed character is one insertion of the author (one change id).
			if (tracking) {
				const ids = new Set<unknown>()

				paragraph.forEach(node => {
					const mark = docxSchema.marks.insertion.isInSet(node.marks)

					if (mark !== undefined) ids.add(mark.attrs["id"])
				})
				expect(ids.size).toBe(1)
			}

			expect(typing.execute({ type: "history.undo" })).toBe(true)
			expect(typing.state.doc.child(0).textContent).toBe("one two")
		}

		expect(depths).toEqual([1, 1])
		// The keys' steps joined into one, tracked as untracked (each key two unjoinable items before).
		expect(items[0]).toBeGreaterThan(0)
		expect(items[1]).toBe(items[0])
	})
})
