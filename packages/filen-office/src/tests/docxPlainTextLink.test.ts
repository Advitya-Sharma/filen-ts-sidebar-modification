import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { writerProof } from "@filen/office/core/docx/write"
import { openBytes } from "./docxEditFixtures"
import { buildDocx } from "./docxSynthetic"

// A plain text control holds one run with one set of formatting properties (ISO/IEC 29500-1 §17.5.2.44): a link
// there, by the command or by any transaction, is refused; a rich text control and plain text take one.

function control(kind: string, text: string): string {
	return `<w:sdt><w:sdtPr><w:id w:val="1"/>${kind}</w:sdtPr><w:sdtContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:sdtContent></w:sdt>`
}

async function linked(kind: string, how: "command" | "raw"): Promise<{ applied: boolean; links: number }> {
	const opened = await openBytes(buildDocx({ body: control(kind, "Sender name here") }))
	const session = new DocxEditorSession(opened.context, { proof: writerProof(opened.context) })
	let text = -1

	session.state.doc.descendants((node, pos) => {
		if (text < 0 && node.isText) text = pos
	})

	const [from, to] = [text + 7, text + 11]

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, from, to)))

	const before = session.state.doc
	const applied =
		how === "command"
			? session.execute({ type: "link.set", target: "https://filen.io/" })
			: (session.dispatch(session.state.tr.addMark(from, to, docxSchema.marks.link.create({ target: "https://filen.io/" }))),
				session.state.doc !== before)
	let links = 0

	session.state.doc.descendants(node => {
		if (node.isText && docxSchema.marks.link.isInSet(node.marks) !== undefined) links++
	})

	return { applied, links }
}

describe("links in plain text controls", () => {
	it("are refused, by the command and by a transaction", async () => {
		expect(await linked("<w:text/>", "command")).toEqual({ applied: false, links: 0 })
		expect(await linked("<w:text/>", "raw")).toEqual({ applied: false, links: 0 })
	})

	it("go on text in a rich text control", async () => {
		expect(await linked("", "command")).toEqual({ applied: true, links: 1 })
		expect(await linked("", "raw")).toEqual({ applied: true, links: 1 })
	})
})
