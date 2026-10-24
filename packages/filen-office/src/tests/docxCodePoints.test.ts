import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { saveDocx } from "@filen/office/core/docx/write"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { openSpec } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

// Positions inside a UTF-16 surrogate pair are no place for a selection or an edit: a selection snaps to the
// code point's edges, and an edit splitting one is refused, so half a character never reaches a save.

// "a" + U+1D400 (two code units) + "b": the pair is at 2..4, its middle at 3.
const TEXT = "a\u{1D400}b"

async function opened() {
	const opened = await openSpec({ body: p(TEXT) })

	return { opened, session: new DocxEditorSession(opened.context) }
}

function select(session: DocxEditorSession, anchor: number, head = anchor): void {
	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, anchor, head)))
}

describe("docx editor: code points", () => {
	it("moves a caret put inside a pair to the edge it was moving towards", async () => {
		const { session } = await opened()

		select(session, 2)
		select(session, 3)
		expect(session.state.selection.head).toBe(4)

		select(session, 5)
		select(session, 3)
		expect(session.state.selection.head).toBe(2)
	})

	it("widens a selection with an end inside a pair to the whole character", async () => {
		const { session } = await opened()

		select(session, 1, 3)
		expect([session.state.selection.anchor, session.state.selection.head]).toEqual([1, 4])

		select(session, 3, 5)
		expect([session.state.selection.anchor, session.state.selection.head]).toEqual([2, 5])
	})

	it("refuses an edit splitting a pair, and saves the character whole", async () => {
		const { opened: o, session } = await opened()

		expect(session.dispatch(session.state.tr.delete(3, 4))).toBe(false)
		expect(session.dispatch(session.state.tr.insertText("x", 3))).toBe(false)
		expect(session.state.doc.textContent).toBe(TEXT)

		select(session, 1, 3)
		expect(session.dispatch(session.state.tr.deleteSelection())).toBe(true)
		expect(session.state.doc.textContent).toBe("b")

		const result = await saveDocx({ pkg: o.pkg, context: o.context, sessions: [session] })
		const xml = new TextDecoder().decode(
			await (await OpcPackage.open(bytesSource(result.bytes as Uint8Array))).read("word/document.xml")
		)

		expect(xml).not.toContain("�")
		expect(xml).toContain(">b<")
	})
})
