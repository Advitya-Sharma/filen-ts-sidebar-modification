import { Fragment, type Node } from "prosemirror-model"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { markCommand } from "../core/docx/edit/invariants/core"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"

// A host's edit that rebuilds a source block with copies of what is elsewhere inside it (a paste into a content
// control, the control written again): the copies become new content, as any other copy does, and the save
// writes what the editor shows.

const nodes = docxSchema.nodes
const SDT = (inner: string): string => `<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
const MARKED = '<w:p><w:bookmarkStart w:id="5" w:name="mark"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p>'
// A field with no end: its paragraph is locked.
const LOCKED = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r></w:p>'

function inventory(doc: Node): string[] {
	const out: string[] = []

	doc.descendants(node => {
		if (node.type === nodes.marker) out.push(`${String(node.attrs["kind"])}:${String(node.attrs["id"])}`)
		else if (node.type === nodes.locked_block || node.type === nodes.paragraph)
			out.push(`${node.type.name}#${String(node.attrs["src"])}`)

		return true
	})

	return out
}

async function rebuilt(
	body: string,
	prepare?: (session: DocxEditorSession) => void
): Promise<{ before: string[]; after: string[]; saved: boolean }> {
	const opened = await openBytes(buildDocx({ body }))
	const session = new DocxEditorSession(opened.context)

	prepare?.(session)

	const doc = session.state.doc
	const copied = doc.child(0)
	const sdt = doc.child(1)
	const before = inventory(doc)

	// The control written again with a copy of the first block before its own content.
	session.dispatch(
		session.state.tr.replaceWith(copied.nodeSize, copied.nodeSize + sdt.nodeSize, sdt.copy(Fragment.from(copied).append(sdt.content)))
	)

	const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })

	return { before, after: inventory(session.state.doc), saved: result.bytes !== undefined }
}

describe("copies inside a source block an edit rebuilt", () => {
	it("makes a copied paragraph new, its bookmark left with the paragraph it was copied from", async () => {
		const { before, after, saved } = await rebuilt(MARKED + SDT(p("Inside")) + p("After"))

		expect(after.filter(entry => entry.startsWith("bookmark"))).toEqual(before.filter(entry => entry.startsWith("bookmark")))
		expect(after.filter(entry => entry === "paragraph#-1")).toHaveLength(1)
		expect(saved).toBe(true)
	})

	it("leaves a bookmark the editor made (a table of contents' entry) with the paragraph it was copied from", async () => {
		const { before, after, saved } = await rebuilt(p("Heading") + SDT(p("Inside")) + p("After"), session => {
			const end = session.state.doc.child(0).nodeSize - 1

			session.dispatch(
				markCommand(
					session.state.tr
						.insert(end, nodes.marker.create({ src: -1, kind: "bookmarkEnd", id: "77", name: null }))
						.insert(1, nodes.marker.create({ src: -1, kind: "bookmarkStart", id: "77", name: "_Toc77" }))
				)
			)
		})

		expect(before).toContain("bookmarkStart:77")
		expect(after.filter(entry => entry === "bookmarkStart:77")).toHaveLength(1)
		expect(saved).toBe(true)
	})

	it("never puts a locked block in twice", async () => {
		const { before, after, saved } = await rebuilt(LOCKED + SDT(p("Inside")) + p("After"))

		expect(before.filter(entry => entry.startsWith("locked_block"))).toHaveLength(1)
		expect(after.filter(entry => entry.startsWith("locked_block"))).toHaveLength(1)
		expect(saved).toBe(true)
	})
})
