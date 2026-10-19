import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { openBytes, openSpec } from "./docxEditFixtures"

// A proofing mark left alone under a link is written without it (a hyperlink holding only w:proofErr reads back
// as an empty hyperlink), also when another link to the same target is elsewhere in the paragraph.
describe("docx writer: a proofing mark under a link", () => {
	it("writes no hyperlink around a lone proofing mark next to an equal link", async () => {
		const opened = await openSpec({
			body: '<w:p><w:r><w:t>first</w:t></w:r><w:r><w:t xml:space="preserve"> mid </w:t></w:r><w:proofErr w:type="gramEnd"/><w:r><w:t>end</w:t></w:r></w:p>'
		})
		const session = new DocxEditorSession(opened.context)
		const paragraph = session.state.doc.child(0)

		expect(paragraph.child(2).type).toBe(docxSchema.nodes.proof_err)

		const link = docxSchema.marks.link.create({ target: "https://filen.io/" })

		expect(session.dispatch(session.state.tr.addMark(1, 6, link).addMark(11, 12, link))).toBe(true)

		const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })

		if (result.bytes === undefined) throw new Error("no bytes")

		const reopened = new DocxEditorSession((await openBytes(result.bytes)).context)
		const kinds: string[] = []

		reopened.state.doc.child(0).forEach(child => kinds.push(`${child.type.name}:${child.marks.map(mark => mark.type.name).join("+")}`))

		expect(reopened.state.doc.child(0).textContent).toBe("first mid end")
		expect(kinds.filter(kind => kind.includes("link"))).toHaveLength(1)
		expect(kinds.some(kind => kind.startsWith("opaque_inline"))).toBe(false)
	})
})
