import { describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { saveDocx, writerProof } from "@filen/office/core/docx/write"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"

// A paragraph alignment of "start" or "end" (a raw attribute, or one carried from a file that wrote it) is
// written as "left" or "right": the same edges in Word for either direction, and what Word 2007 reads.

async function written(body: string, align: string): Promise<string> {
	const bytes = buildDocx({ body })
	const opened = await openBytes(bytes)
	const session = new DocxEditorSession(opened.context, { proof: writerProof(opened.context) })
	const tr = session.state.tr

	session.state.doc.forEach((node, offset) => {
		if (node.type.name === "paragraph") tr.setNodeMarkup(offset, undefined, { ...node.attrs, align })
	})
	expect(session.dispatch(tr)).toBe(true)

	const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })
	const pkg = await OpcPackage.open(bytesSource(result.bytes as Uint8Array))

	return new TextDecoder().decode(await pkg.read("word/document.xml"))
}

describe("paragraph alignment as written", () => {
	it("writes start and end as left and right, and the save passes Gate 2", async () => {
		const end = await written(p("Alpha") + p("Beta"), "end")

		expect(end.match(/<w:jc w:val="right"\/>/g)).toHaveLength(2)
		expect(end).not.toMatch(/w:val="(start|end)"/)
		expect(await written(p("Alpha"), "start")).toContain('<w:jc w:val="left"/>')
	})

	it("writes them so in a right-to-left paragraph too, whose left and right Word reads as its start and end", async () => {
		const xml = await written("<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>", "end")

		expect(xml).toMatch(/<w:bidi\/><w:jc w:val="right"\/>/)
	})

	it("keeps a source's own start and end where nothing changed them", async () => {
		const bytes = buildDocx({ body: '<w:p><w:pPr><w:jc w:val="end"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>' })
		const opened = await openBytes(bytes)
		const session = new DocxEditorSession(opened.context, { proof: writerProof(opened.context) })
		const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })

		expect(result.changedParts).toEqual([])
	})
})
