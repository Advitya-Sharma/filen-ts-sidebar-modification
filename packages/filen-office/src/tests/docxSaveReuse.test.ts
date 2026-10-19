import { unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { StoryWriter, type WriteEnvironment } from "../core/docx/write/story"
import { readBacksKept } from "../core/docx/write/verify"
import { openSpec } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

// A save after earlier saves (blocks written then are written again as they were, read-backs found by their
// markup) writes what one save of the same edits writes, and gate 2 still checks every rewritten paragraph.

const BODY = Array.from({ length: 40 }, (_, k) =>
	k % 7 === 3
		? `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold ${String(k)}</w:t></w:r><w:r><w:t xml:space="preserve"> and plain</w:t></w:r></w:p>`
		: p(`Paragraph ${String(k)}`)
).join("")

function typeAt(session: DocxEditorSession, paragraph: number, text: string): void {
	let pos = 0

	for (let k = 0; k < paragraph; k++) pos += session.state.doc.child(k).nodeSize

	const state = session.state

	expect(session.dispatch(state.tr.setSelection(TextSelection.create(state.doc, pos + 2)).insertText(text))).toBe(true)
}

describe("docx saves after saves", () => {
	it("writes the same bytes as a single save of the same edits", async () => {
		const edits: [number, string][] = [
			[1, "a"],
			[3, "b"],
			[10, "c"],
			[1, "d"],
			[24, "e"],
			[17, "f"]
		]
		const many = await openSpec({ body: BODY })
		const session = new DocxEditorSession(many.context)
		let last: Uint8Array | undefined

		for (const [paragraph, text] of edits) {
			typeAt(session, paragraph, text)

			const saved = await saveDocx({ pkg: many.pkg, context: many.context, sessions: [session] })

			session.markSaved()
			last = saved.bytes
		}

		const once = await openSpec({ body: BODY })
		const fresh = new DocxEditorSession(once.context)

		for (const [paragraph, text] of edits) typeAt(fresh, paragraph, text)

		const single = await saveDocx({ pkg: once.pkg, context: once.context, sessions: [fresh] })

		// The parts (the archive's times are the clock's).
		const parts = (bytes: Uint8Array | undefined): Record<string, Uint8Array> => (bytes === undefined ? {} : unzipSync(bytes))

		expect(last).toBeDefined()
		expect(parts(single.bytes)).toEqual(parts(last))
	})

	it("writes a block again only as the same kind of writer wrote it", async () => {
		const opened = await openSpec({
			body:
				'<w:p><w:r><w:t xml:space="preserve">Typed in </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>copied</w:t></w:r></w:p>' +
				p("Two")
		})
		const session = new DocxEditorSession(opened.context)

		typeAt(session, 0, "x")

		const env: WriteEnvironment = { context: opened.context, nextRevisionId: () => 1, numId: id => id, hyperlink: () => "rId9" }
		// A save rewriting stories inside the source bytes it copies (text boxes) splices them in.
		const splicing: WriteEnvironment = {
			...env,
			nested: (_part, from) => [{ from, to: from, bytes: new TextEncoder().encode("<!--spliced-->") }]
		}
		const write = (with_: WriteEnvironment): string =>
			new TextDecoder().decode(new StoryWriter(with_, session.story).draftStorySync(session.state.doc).out.finish())

		// The edited paragraph's copied run, with what is spliced into it.
		const copied = /<!--spliced--><w:r><w:rPr><w:b\/>/

		expect(write(env)).not.toMatch(copied)
		expect(write(splicing)).toMatch(copied)
		expect(write(env)).not.toMatch(copied)
	})
	it("keeps the read-backs of every story of a part it rewrote, not the last one's only", async () => {
		const opened = await openSpec({
			body: '<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
			footnotes:
				'<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
				'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
				'<w:footnote w:id="1"><w:p><w:r><w:t>First note</w:t></w:r></w:p></w:footnote>' +
				'<w:footnote w:id="2"><w:p><w:r><w:t>Second note</w:t></w:r></w:p></w:footnote>'
		})
		const notes = opened.context.model.stories.filter(story => story.kind === "footnote" && Number(story.id) > 0)
		const sessions = notes.map(story => new DocxEditorSession(opened.context, { story }))

		for (const session of sessions) expect(session.dispatch(session.state.tr.insertText("x", 2))).toBe(true)

		await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [new DocxEditorSession(opened.context), ...sessions] })

		const part = opened.context.model.parts[notes[0]?.part ?? -1]

		expect(part).toBeDefined()
		expect(readBacksKept(part as NonNullable<typeof part>)).toBe(2)
	})
})
