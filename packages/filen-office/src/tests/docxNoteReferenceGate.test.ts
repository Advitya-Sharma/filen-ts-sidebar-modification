import type { Node } from "prosemirror-model"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { markInternal } from "../core/docx/edit/invariants/index"
import { openSpec } from "./docxEditFixtures"
import { p } from "./docxSynthetic"
import { pngHeader } from "./fuzz/ops/drawings"

// Gate 2 refuses a new note reference that does not resolve to a normal note (none with its id, or only a separator
// or continuation notice with it), any new note reference outside the body, and a new floating drawing in a comment.

const notes = (kind: "footnote" | "endnote"): string =>
	`<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
	`<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
	`<w:${kind} w:type="continuationNotice" w:id="1"><w:p/></w:${kind}>` +
	`<w:${kind} w:id="2"><w:p><w:r><w:t>${kind} two</w:t></w:r></w:p></w:${kind}>`

const SPEC = {
	body:
		p("Body") +
		'<w:p><w:r><w:t>Notes</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>' +
		'<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr>',
	headers: { rIdH: p("Header") },
	footers: { rIdF: p("Footer") },
	footnotes: notes("footnote"),
	endnotes: notes("endnote"),
	comments: '<w:comment w:id="0" w:author="Bo"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment>'
}

async function saveWithReference(storyKind: string, kind: "footnote" | "endnote", id: string): Promise<unknown> {
	const opened = await openSpec(SPEC)
	const story = opened.model.stories.find(
		candidate => candidate.kind === storyKind && (candidate.id === undefined || candidate.id === "2" || storyKind === "comment")
	)

	if (story === undefined) throw new Error(`no ${storyKind} story`)

	const body = new DocxEditorSession(opened.context)
	const session = storyKind === "body" ? body : new DocxEditorSession(opened.context, { story })

	expect(session.dispatch(session.state.tr.insert(1, docxSchema.nodes.note_ref.create({ kind, id })))).toBe(true)

	try {
		await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: storyKind === "body" ? [body] : [body, session] })

		return "saved"
	} catch (error) {
		return error
	}
}

describe("docx gate 2: new note references", () => {
	for (const kind of ["footnote", "endnote"] as const) {
		it(`refuses a ${kind} reference in the body to a continuation notice, a separator or a missing note`, async () => {
			for (const id of ["1", "9", "0"])
				expect(await saveWithReference("body", kind, id)).toMatchObject({ reason: "danglingReference" })
		})
	}

	// Word keeps note references to the body (it inserts none elsewhere) and LibreOffice opens no file with one.
	for (const storyKind of ["header", "footer", "footnote", "endnote", "comment"]) {
		for (const kind of ["footnote", "endnote"] as const) {
			it(`refuses a ${kind} reference in a ${storyKind} story, even to a normal note`, async () => {
				for (const id of ["2", "1", "9"]) expect(await saveWithReference(storyKind, kind, id)).toMatchObject({ reason: "content" })
			})
		}
	}

	it("still saves the document with its normal notes' references", async () => {
		const opened = await openSpec(SPEC)
		const body = new DocxEditorSession(opened.context)

		expect(body.dispatch(body.state.tr.insertText("Edited ", 1))).toBe(true)

		const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [body] })

		expect(result.changedParts).toEqual(["word/document.xml"])
	})

	it("refuses a floating drawing in a comment, not an inline one", async () => {
		const save = async (floating: boolean): Promise<unknown> => {
			const opened = await openSpec(SPEC)
			const story = opened.model.stories.find(candidate => candidate.kind === "comment")

			if (story === undefined) throw new Error("no comment story")

			const comment = new DocxEditorSession(opened.context, { story })
			const media = opened.context.media.add(pngHeader(96, 48))

			if (media === undefined) throw new Error("no media")

			const picture = docxSchema.nodes.drawing.create({
				src: -1,
				kind: "drawing",
				index: -1,
				overrides: {
					media,
					cx: 914400,
					cy: 457200,
					docPr: opened.context.ids.next("docPr"),
					descr: "picture",
					...(floating ? { wrap: "square", posX: 0, posY: 0, relH: "column", relV: "paragraph" } : {})
				}
			})

			expect(comment.dispatch(comment.state.tr.insert(1, picture))).toBe(true)

			return saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [comment] }).then(
				() => "saved",
				(error: unknown) => error
			)
		}

		expect(await save(true)).toMatchObject({ reason: "content", message: expect.stringContaining("floating drawing") as unknown })
		expect(await save(false)).toBe("saved")
	})

	// ECMA-376 Part 1 §17.11.14 has no rule against two references to one note, but the editor keeps one per
	// note, as Word makes them (a second mention is a NOTEREF field): a typed copy is left out, a raw one refused.
	for (const kind of ["footnote", "endnote"] as const) {
		it(`keeps one reference per ${kind}`, async () => {
			const opened = await openSpec(SPEC)
			const body = new DocxEditorSession(opened.context)
			const copy = (): Node => docxSchema.nodes.note_ref.create({ kind, id: "2" })
			const count = (): number => {
				let found = 0

				body.state.doc.descendants(node => {
					if (node.type === docxSchema.nodes.note_ref && node.attrs["kind"] === kind) found++
				})

				return found
			}

			body.dispatch(body.state.tr.insert(1, copy()))
			expect(count()).toBe(1)

			body.dispatch(markInternal(body.state.tr.insert(1, copy())))
			expect(count()).toBe(2)
			expect(
				await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [body] }).then(
					() => "saved",
					(error: unknown) => error
				)
			).toMatchObject({ reason: "duplicateId" })
		})
	}
})
