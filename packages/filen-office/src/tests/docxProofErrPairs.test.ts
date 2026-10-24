import { describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { openSpec } from "./docxEditFixtures"

// A proofing region (w:proofErr start and end, ECMA-376 §17.13.8) is written whole or not at all: an edited
// paragraph never keeps one half of a pair. A region whose own text is written as it was stays, so the saved
// file reads back as the editor holds it.

const decoder = new TextDecoder()

async function saved(body: string, edit: (session: DocxEditorSession) => void): Promise<string> {
	const opened = await openSpec({ body })
	const session = new DocxEditorSession(opened.context)

	edit(session)

	const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })

	if (result.bytes === undefined) throw new Error("no bytes")

	return decoder.decode(await (await OpcPackage.open(bytesSource(result.bytes))).read("word/document.xml"))
}

function count(xml: string, type: string): number {
	return xml.split(`w:type="${type}"`).length - 1
}

function at(session: DocxEditorSession, text: string): number {
	let found = -1

	session.state.doc.descendants((node, pos) => {
		if (found < 0 && node.isText && (node.text ?? "").includes(text)) found = pos + (node.text ?? "").indexOf(text)

		return found < 0
	})

	return found
}

const insert = (text: string, after: string) => (session: DocxEditorSession) => {
	expect(session.dispatch(session.state.tr.insertText(text, at(session, after) + after.length))).toBe(true)
}

describe("docx writer: proofing marks in an edited paragraph", () => {
	it("keeps both halves when only the start stands next to rebuilt text", async () => {
		const xml = await saved(
			'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve">plain </w:t></w:r><w:proofErr w:type="gramStart"/><w:r><w:t>word</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>',
			insert("X", "pla")
		)

		expect(xml).toContain("plaXin")
		expect(count(xml, "gramStart")).toBe(1)
		expect(count(xml, "gramEnd")).toBe(1)
	})

	it("keeps both halves when only the end stands next to rebuilt text", async () => {
		const xml = await saved(
			'<w:p><w:r><w:t xml:space="preserve">head </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>wrod</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> tail</w:t></w:r></w:p>',
			insert("Y", " ta")
		)

		expect(xml).toContain("taYil")
		expect(count(xml, "spellStart")).toBe(1)
		expect(count(xml, "spellEnd")).toBe(1)
	})

	it("drops a region whose inner text was rebuilt", async () => {
		const xml = await saved(
			'<w:p><w:proofErr w:type="gramStart"/><w:r><w:t xml:space="preserve">one </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t></w:r><w:r><w:t xml:space="preserve"> three</w:t></w:r><w:proofErr w:type="gramEnd"/></w:p>',
			insert("Z", "tw")
		)

		expect(xml).toContain("twZo")
		expect(count(xml, "gramStart")).toBe(0)
		expect(count(xml, "gramEnd")).toBe(0)
	})

	it("keeps a region whose text was written as it was", async () => {
		const xml = await saved(
			'<w:p><w:r><w:t xml:space="preserve">left </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">x </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>wrod</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>',
			insert("Q", "lef")
		)

		expect(xml).toContain("lefQt")
		expect(count(xml, "spellStart")).toBe(1)
		expect(count(xml, "spellEnd")).toBe(1)
	})

	it("keeps a half whose partner is in another paragraph", async () => {
		const xml = await saved(
			'<w:p><w:r><w:t xml:space="preserve">first </w:t></w:r><w:proofErr w:type="gramStart"/><w:r><w:t>end</w:t></w:r></w:p><w:p><w:r><w:t>next</w:t></w:r><w:proofErr w:type="gramEnd"/></w:p>',
			insert("W", "fir")
		)

		expect(xml).toContain("firWst")
		expect(count(xml, "gramStart")).toBe(1)
		expect(count(xml, "gramEnd")).toBe(1)
	})
})
