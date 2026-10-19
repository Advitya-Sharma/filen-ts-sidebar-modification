import { unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { openBytes } from "./docxEditFixtures"
import { buildDocx } from "./docxSynthetic"

// A deletion ending inside a row or cell it did not start in opens that row or cell again: ProseMirror's
// fitting rebuilds it from its type and attributes, without its node marks. What commands recorded on it (a
// tracked deletion's stamp, its properties) stays with it all the same, so the save writes it whoever edits
// after (the stamp's author, not the editor's, which may be none).

const cell = (text: string): string => `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`
const rows = Array.from({ length: 4 }, (_, r) => `<w:tr>${cell(`a${r} text`)}${cell(`b${r} text`)}${cell(`c${r} text`)}</w:tr>`).join("")
const BYTES = buildDocx({
	body: `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl><w:p/>`
})

async function editing(): Promise<{
	session: DocxEditorSession
	opened: Awaited<ReturnType<typeof openBytes>>
	at: (text: string) => number
	select: (from: number, to?: number) => void
	// Inside the empty cell a column inserted right of `a<row>`.
	inserted: (row: number) => number
}> {
	const opened = await openBytes(BYTES)
	const session = new DocxEditorSession(opened.context)
	const at = (text: string): number => {
		let found = -1

		session.state.doc.descendants((node, pos) => {
			if (found < 0 && node.type === docxSchema.nodes.paragraph && node.textContent.startsWith(text)) found = pos + 1

			return found < 0
		})

		if (found < 0) throw new Error(`no ${text}`)

		return found
	}
	const select = (from: number, to = from): void => {
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, from, to)))
	}

	select(at("a1") + 1)
	expect(session.execute({ type: "table.column.insertRight" })).toBe(true)

	return {
		session,
		opened,
		at,
		select,
		inserted: row => {
			const $a = session.state.doc.resolve(at(`a${String(row)}`))

			return $a.after($a.depth - 1) + 2
		}
	}
}

async function saved(session: DocxEditorSession, opened: Awaited<ReturnType<typeof openBytes>>): Promise<string> {
	const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })

	if (result.bytes === undefined) throw new Error("no bytes")

	return new TextDecoder().decode(unzipSync(result.bytes)["word/document.xml"])
}

describe("rows and cells a deletion opens again", () => {
	it("keep a tracked deletion's stamp: the save writes it after edits by no one", async () => {
		const { session, opened, at, select, inserted } = await editing()

		// A tracked column deletion by a reviewer, then no author and tracking off.
		opened.context.review.tracking = true
		opened.context.identity.author = "Reviewer"
		select(at("b1") + 1)
		expect(session.execute({ type: "table.column.delete" })).toBe(true)
		opened.context.review.tracking = false
		opened.context.identity.author = undefined

		// From the empty new cell into the tracked-deleted one after it.
		select(inserted(2), at("b2") + 3)
		session.dispatch(session.state.tr.deleteSelection())

		const xml = await saved(session, opened)

		expect(xml.match(/<w:cellDel [^>]*w:author="Reviewer"/g)?.length).toBe(4)
		expect(xml).not.toMatch(/<w:cellDel(?![^>]*w:author=)/)
	})

	it("keep the properties a command gave them", async () => {
		const { session, opened, at, select, inserted } = await editing()

		select(at("b1") + 1)
		expect(session.execute({ type: "table.cell.shading", fill: "FFCC00" })).toBe(true)
		select(inserted(1), at("b1") + 3)
		session.dispatch(session.state.tr.deleteSelection())

		expect(await saved(session, opened)).toMatch(/w:fill="FFCC00"/)
	})
})
