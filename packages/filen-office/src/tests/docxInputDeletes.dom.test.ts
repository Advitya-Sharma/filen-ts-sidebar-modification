import { TextSelection, type EditorState } from "prosemirror-state"
import type { EditorView } from "prosemirror-view"
import { describe, expect, it } from "vitest"
import { inputTransaction } from "../react/docx/edit/input"
import { sessionOf } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

// Word and line deletions a beforeinput names a range for stay in the caret's paragraph: a browser's range past
// its end (Chromium's Ctrl+K on macOS) deletes only to the paragraph's edge; at the edge, the paragraphs join.

// A view whose DOM positions are document positions, and a beforeinput naming [from, to].
function input(state: EditorState, type: string, from: number, to: number) {
	const view = { composing: false, posAtDOM: (_node: unknown, offset: number) => offset } as unknown as EditorView
	const range = { collapsed: from === to, startContainer: {}, startOffset: from, endContainer: {}, endOffset: to }
	const event = { inputType: type, isComposing: false, data: null, getTargetRanges: () => [range] } as unknown as InputEvent

	return inputTransaction(view, state, event)
}

async function at(offset: number): Promise<EditorState> {
	const session = await sessionOf({ body: p("Form of the account") + p("Next paragraph here") })
	const state = session.state

	return state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1 + offset)))
}

function texts(state: EditorState): string[] {
	const out: string[] = []

	state.doc.forEach(node => out.push(node.textContent))

	return out
}

describe("docx input: word and line deletions", () => {
	it("deletes a line forward only to the end of the caret's paragraph", async () => {
		const state = await at(3)
		// "m of the account", the paragraph's end, then "Nex": three characters past it.
		const tr = input(state, "deleteHardLineForward", 4, 4 + 16 + 2 + 3)

		expect(tr).toBeDefined()
		expect(texts(state.apply(tr as never))).toEqual(["For", "Next paragraph here"])
	})

	it("deletes a word or a soft line backward only to the start of the caret's paragraph", async () => {
		const session = await sessionOf({ body: p("Form of the account") + p("Next paragraph here") })
		const second = session.state.doc.child(0).nodeSize + 1
		const state = session.state.apply(session.state.tr.setSelection(TextSelection.create(session.state.doc, second + 4)))

		for (const type of ["deleteWordBackward", "deleteSoftLineBackward", "deleteHardLineBackward"]) {
			const tr = input(state, type, second - 5, second + 4)

			expect(texts(state.apply(tr as never)), type).toEqual(["Form of the account", " paragraph here"])
		}
	})

	it("clamps an entire soft line to the caret's paragraph", async () => {
		const state = await at(5)
		const tr = input(state, "deleteEntireSoftLine", 1, 1 + 16 + 2 + 4 + 3)

		expect(texts(state.apply(tr as never))).toEqual(["", "Next paragraph here"])
	})

	it("joins the next paragraph at the paragraph's end, deleting none of its text", async () => {
		const state = await at(19)
		const tr = input(state, "deleteHardLineForward", 20, 20 + 2 + 19)

		expect(texts(state.apply(tr as never))).toEqual(["Form of the accountNext paragraph here"])
	})

	it("joins the paragraph before at the paragraph's start", async () => {
		const session = await sessionOf({ body: p("Form of the account") + p("Next paragraph here") })
		const second = session.state.doc.child(0).nodeSize + 1
		const state = session.state.apply(session.state.tr.setSelection(TextSelection.create(session.state.doc, second)))
		const tr = input(state, "deleteWordBackward", second - 9, second)

		expect(texts(state.apply(tr as never))).toEqual(["Form of the accountNext paragraph here"])
	})

	it("leaves a word deletion inside the paragraph as the browser names it", async () => {
		const state = await at(8)
		const tr = input(state, "deleteWordBackward", 1 + 5, 1 + 8)

		expect(texts(state.apply(tr as never))).toEqual(["Form the account", "Next paragraph here"])
	})

	it("widens a range with an end inside a surrogate pair to the whole character", async () => {
		const session = await sessionOf({ body: p("a\u{1D400}\u{1D401}b") })
		const state = session.state.apply(session.state.tr.setSelection(TextSelection.create(session.state.doc, 6)))
		// From inside the first pair (3) to inside the second (5).
		const tr = input(state, "deleteWordBackward", 3, 5)

		expect(texts(state.apply(tr as never))).toEqual(["ab"])
	})
})
