import { Fragment, Slice, type Node } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import { ReplaceAroundStep } from "prosemirror-transform"
import { describe, expect, it } from "vitest"
import * as edit from "@filen/office/core/docx/edit"
import { DocxEditorSession, NO_SOURCE, docxInvariants, docxSchema } from "@filen/office/core/docx/edit"
import { RICH, bodyOf, markerKeys, openSpec, sessionOf } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

function types(node: Node): string[] {
	const out: string[] = []

	node.forEach(child => out.push(child.type.name))

	return out
}

function run(session: DocxEditorSession, build: (state: EditorState) => Transaction): boolean {
	return session.dispatch(build(session.state))
}

function sources(doc: Node): number[] {
	const out: number[] = []

	doc.descendants(node => {
		if (node.isBlock && typeof node.attrs["src"] === "number" && node.attrs["src"] !== NO_SOURCE) out.push(node.attrs["src"] as number)

		return true
	})

	return out
}

describe("docx editing: source identity", () => {
	it("keeps every source run, proofErr and rendered page break as its own node", async () => {
		const story = await bodyOf({
			body:
				'<w:p><w:r w:rsidR="1"><w:t>ab</w:t><w:t>cd</w:t></w:r><w:proofErr w:type="spellStart"/><w:r w:rsidR="2"><w:t>ef</w:t></w:r>' +
				'<w:hyperlink w:anchor="x"><w:proofErr w:type="spellEnd"/><w:r><w:lastRenderedPageBreak/><w:t>g</w:t></w:r></w:hyperlink><w:proofErr w:type="gramStart"/></w:p>'
		})
		const paragraph = story.doc.child(0)

		expect(types(paragraph)).toEqual(["text", "proof_err", "text", "proof_err", "rendered_break", "text", "proof_err"])
		// One run with two w:t is one text node with that run's mark.
		expect(paragraph.child(0).text).toBe("abcd")
		expect(paragraph.child(0).marks[0]?.attrs["src"]).not.toBe(paragraph.child(2).marks[0]?.attrs["src"])
		expect(paragraph.child(3).marks.map(mark => mark.type.name)).toEqual(["link"])
		expect(new Set([paragraph.child(1), paragraph.child(3), paragraph.child(6)].map(node => node.attrs["node"])).size).toBe(3)
	})

	it("keeps range markers inside a run as markers of that run", async () => {
		const session = await sessionOf({
			body: '<w:p><w:r><w:bookmarkStart w:id="9" w:name="inRun"/><w:t>x</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>' + p("y")
		})
		const paragraph = session.state.doc.child(0)

		expect(types(paragraph)).toEqual(["marker", "text", "marker"])
		expect(paragraph.child(0).marks.map(mark => mark.type.name)).toEqual(["run"])
		expect(paragraph.child(0).marks[0]?.attrs["src"]).toBe(paragraph.child(1).marks[0]?.attrs["src"])

		// Deleting the paragraph keeps both ends.
		const before = markerKeys(session.state.doc)

		expect(run(session, state => state.tr.delete(0, paragraph.nodeSize))).toBe(true)
		expect(markerKeys(session.state.doc)).toEqual(before)
	})

	it("locks every paragraph whose run properties are malformed, however often they repeat", async () => {
		const malformed = "<w:r><w:rPr><w:i/><w:b w:val/></w:rPr><w:t>x</w:t></w:r>"
		const { context, model } = await openSpec({ body: `<w:p>${malformed}</w:p><w:p>${malformed}</w:p>${p("ok")}` })
		const body = model.body

		if (body === undefined) throw new Error("no body")

		for (let build = 0; build < 2; build++) {
			const story = new edit.DocxDocBuilder(context).build(body)

			expect(types(story.doc)).toEqual(["locked_block", "locked_block", "paragraph"])
			expect(story.stats.lockReasons).toEqual({ malformed: 2 })
		}

		expect(context.gate).toContain("malformed")
	})

	it("turns inserted copies of source content into new content", async () => {
		const session = await sessionOf(RICH)
		const doc = session.state.doc
		const first = doc.child(0)
		const keys = markerKeys(doc)
		const end = doc.content.size

		// A host's transaction (no paste sanitiser) inserting a copy of the first paragraph and of a table.
		expect(run(session, state => state.tr.insert(end, [first.copy(first.content), doc.child(1)]))).toBe(true)

		const after = session.state.doc

		expect(markerKeys(after)).toEqual(keys)
		expect(new Set(sources(after)).size).toBe(sources(after).length)
		expect(after.child(after.childCount - 2).attrs["src"]).toBe(NO_SOURCE)
		expect(after.lastChild?.type.name).toBe("table")
		expect(after.lastChild?.attrs["src"]).toBe(NO_SOURCE)
		after.check()

		// A move keeps identity: the original is gone.
		const original = after.child(1)
		const at = after.child(0).nodeSize

		expect(original.attrs["src"]).toBe(doc.child(1).attrs["src"])
		expect(run(session, state => state.tr.delete(at, at + original.nodeSize).insert(0, original))).toBe(true)
		expect(session.state.doc.child(0).attrs["src"]).toBe(doc.child(1).attrs["src"])
	})

	it("drops formatting marks from paragraph-level atoms and refuses identity marks", async () => {
		const session = await sessionOf(RICH)
		const paragraph = session.state.doc.child(0)

		expect(
			run(session, state => state.tr.addMark(1, paragraph.nodeSize - 1, docxSchema.marks.bold.create({ value: true, cs: true })))
		).toBe(true)

		session.state.doc.child(0).forEach(child => {
			const formatted = child.marks.some(mark => mark.type.name === "bold")

			expect(formatted, child.type.name).toBe(edit.inlineLevel(child) === "run")
		})

		expect(run(session, state => state.tr.addMark(1, 4, docxSchema.marks.deletion.create({ src: 99, id: 1, author: "X" })))).toBe(false)
		expect(run(session, state => state.tr.addMark(1, 4, docxSchema.marks.run.create({ src: 3, base: -1 })))).toBe(false)
		expect(run(session, state => state.tr.addMark(1, 4, docxSchema.marks.link.create({ target: "https://example.org/" })))).toBe(true)
	})

	it("protects the attributes that tie nodes to their source", async () => {
		const session = await sessionOf(RICH)
		let locked = -1
		let section = -1

		session.state.doc.forEach((node, offset) => {
			if (locked < 0 && node.attrs["deletable"] === false) locked = offset
			if (section < 0 && node.type.name === "paragraph" && node.attrs["sectPr"] !== NO_SOURCE) section = offset
		})

		const doc = session.state.doc

		expect(run(session, state => state.tr.setNodeAttribute(locked, "deletable", true))).toBe(false)
		expect(run(session, state => state.tr.setNodeAttribute(section, "sectPr", NO_SOURCE))).toBe(false)
		expect(run(session, state => state.tr.setNodeMarkup(section, undefined, { ...state.doc.child(2).attrs, src: 1 }))).toBe(false)
		expect(run(session, state => state.tr.setDocAttribute("end", []))).toBe(false)
		expect(session.state.doc).toBe(doc)
		// Attributes commands edit are open.
		expect(run(session, state => state.tr.setNodeAttribute(section, "align", "center"))).toBe(true)
	})

	it("offers no way past the read-only gate", async () => {
		expect(Object.keys(edit)).not.toContain("INTERNAL_META")
		expect(Object.keys(edit)).not.toContain("markInternal")

		const session = await sessionOf(RICH, { readOnly: true })

		expect(run(session, state => state.tr.insertText("X", 3).setMeta("docx-internal", true))).toBe(false)
	})

	it("strips revision marks from content a ReplaceAroundStep inserts", async () => {
		for (const body of [p("first") + p("last"), p("first") + "<w:p/>"]) {
			const session = await sessionOf({ body })
			const doc = session.state.doc
			const last = doc.content.size - (doc.lastChild?.nodeSize ?? 0)
			const size = doc.lastChild?.nodeSize ?? 0
			const insertion = docxSchema.marks.insertion.create({ src: 5, id: 1, author: "Ann" })
			const slice = new Slice(Fragment.from(docxSchema.nodes.paragraph.create(null, docxSchema.text("Z", [insertion]))), 0, 0)

			expect(
				run(session, state => state.tr.step(new ReplaceAroundStep(last, last + size, last + 1, last + size - 1, slice, 2, true)))
			).toBe(true)

			session.state.doc.descendants(node => {
				expect(node.marks.some(mark => mark.type.name === "insertion")).toBe(false)

				return true
			})
		}
	})

	it("drops a drag-copy inside the document as new content and lets a move keep identity", async () => {
		const { context } = await openSpec({
			body: '<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>drag me</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' + p("target")
		})
		const session = new DocxEditorSession(context)
		const plugin = docxInvariants()
		type DropView = Parameters<NonNullable<typeof plugin.props.handleDrop>>[0]
		type DropEvent = Parameters<NonNullable<typeof plugin.props.handleDrop>>[1]
		const doc = session.state.doc
		const slice = doc.slice(0, doc.child(0).nodeSize)
		const target = doc.content.size - 1
		const view = {
			get state() {
				return session.state
			},
			dragging: { slice, move: true },
			posAtCoords: () => ({ pos: target, inside: -1 }),
			dispatch: (tr: Transaction) => session.dispatch(tr)
		} as unknown as DropView
		const drop = plugin.props.handleDrop

		// The slice is left alone while the drag is under way; a move is ProseMirror's to do.
		expect(plugin.props.transformPasted?.call(plugin, slice, view, false)).toBe(slice)
		expect(drop?.call(plugin, view, {} as DropEvent, slice, true)).toBe(false)
		// A copy goes in fresh.
		expect(drop?.call(plugin, view, { clientX: 1, clientY: 1 } as unknown as DropEvent, slice, false)).toBe(true)

		const after = session.state.doc

		expect(after.childCount).toBe(3)
		expect(markerKeys(after)).toEqual(markerKeys(doc))
		expect(new Set(sources(after)).size).toBe(sources(after).length)
	})
})
