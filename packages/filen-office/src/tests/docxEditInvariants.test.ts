import { Fragment, Slice, type Node } from "prosemirror-model"
import { NodeSelection, TextSelection, type EditorState, type Transaction } from "prosemirror-state"
import { canJoin, canSplit } from "prosemirror-transform"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, NO_SOURCE, docxSchema, freshSlice, normalizeListAttrs } from "@filen/office/core/docx/edit"
import { unzipSync } from "fflate"
import { saveDocx } from "@filen/office/core/docx/write"
import { survivorsOf, type MarkerList } from "../core/docx/edit/schema"
import { RICH, markerKeys, openSpec, sessionOf } from "./docxEditFixtures"
import { p } from "./docxSynthetic"

function sectionHolders(doc: Node): number[] {
	const out: number[] = []

	doc.descendants(node => {
		const sectPr = node.attrs["sectPr"]

		if (node.isBlock && typeof sectPr === "number" && sectPr !== NO_SOURCE) out.push(sectPr)

		return !node.isTextblock
	})

	return out.sort((a, b) => a - b)
}

// Position of the first text containing `text`, plus `offset`.
function find(doc: Node, text: string, offset: number = 0): number {
	let found = -1

	doc.descendants((node, pos) => {
		if (found >= 0) return false

		if (node.isText && node.text?.includes(text) === true) found = pos + node.text.indexOf(text) + offset

		return true
	})

	if (found < 0) throw new Error(`no "${text}"`)

	return found
}

function blockPos(doc: Node, index: number): number {
	let pos = 0

	for (let k = 0; k < index; k++) pos += doc.child(k).nodeSize

	return pos
}

function run(session: DocxEditorSession, build: (state: EditorState) => Transaction): boolean {
	return session.dispatch(build(session.state))
}

// A small deterministic PRNG (mulberry32).
function random(seed: number): () => number {
	let state = seed >>> 0

	return () => {
		state = (state + 0x6d2b79f5) >>> 0

		let t = state

		t = Math.imul(t ^ (t >>> 15), t | 1)
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61)

		return ((t ^ (t >>> 14)) >>> 0) / 4294967296
	}
}

describe("docx editing invariants", () => {
	it("puts back range markers a deletion took out, except a comment's taken whole with its reference", async () => {
		const session = await sessionOf(RICH)
		const before = markerKeys(session.state.doc)
		const doc = session.state.doc
		// From inside "Bold" to after the comment range: bookmark 0, comment range 3 and its reference go; the
		// bookmark's markers come back, the comment goes with them (as in Word).
		const from = find(doc, "Bold", 1)
		const to = find(doc, "c", 1) + 2

		expect(run(session, state => state.tr.delete(from, to))).toBe(true)
		expect(markerKeys(session.state.doc)).toEqual(before.filter(key => !key.startsWith("comment")))
		// They sit where the deletion was.
		const paragraph = session.state.doc.child(0)

		expect(paragraph.child(1).text).toBe("B")
		expect(paragraph.child(2).type.name).toBe("marker")
	})

	it("keeps the markers of a deleted table", async () => {
		const session = await sessionOf(RICH)
		const before = markerKeys(session.state.doc)
		const start = blockPos(session.state.doc, 1)
		const table = session.state.doc.child(1)

		expect(run(session, state => state.tr.delete(start, start + table.nodeSize))).toBe(true)
		expect(session.state.doc.child(1).type.name).toBe("paragraph")
		expect(markerKeys(session.state.doc)).toEqual(before)
	})

	it("keeps a block's leading markers when it is joined into the previous one", async () => {
		const session = await sessionOf({
			body: p("one") + '<w:bookmarkStart w:id="4" w:name="b"/>' + p("two") + '<w:bookmarkEnd w:id="4"/>' + p("three")
		})
		const before = markerKeys(session.state.doc)
		const join = session.state.doc.child(0).nodeSize

		expect(run(session, state => state.tr.join(join))).toBe(true)
		expect(session.state.doc.childCount).toBe(2)
		expect(markerKeys(session.state.doc)).toEqual(before)
		expect(session.state.doc.child(0).textContent).toBe("onetwo")
	})

	it("keeps a range over a paragraph when blocks after it move, and when it moves away and back", async () => {
		const body =
			'<w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>zero</w:t></w:r></w:p><w:bookmarkEnd w:id="0"/>' +
			p("one") +
			p("two") +
			p("three")
		// Moves top-level blocks first..last to before block `target`, as a drop does.
		const move = (session: DocxEditorSession, first: number, last: number, target: number): boolean =>
			run(session, state => {
				const doc = state.doc
				let from = 0
				let at = 0

				for (let k = 0; k < first; k++) from += doc.child(k).nodeSize
				for (let k = 0; k < target; k++) at += doc.child(k).nodeSize

				let to = from

				for (let k = first; k <= last; k++) to += doc.child(k).nodeSize

				const tr = state.tr.delete(from, to)
				const mapped = tr.mapping.map(at)

				return tr.replaceRange(mapped, mapped, doc.slice(from, to))
			})
		// Each top-level block's text with the range's markers around it.
		const shown = (session: DocxEditorSession): string[] => {
			const out: string[] = []

			session.state.doc.forEach(node => {
				const before = (node.attrs["before"] as MarkerList).map(marker => `<${marker.kind}>`).join("")
				let text = ""

				node.forEach(child => (text += child.isText ? (child.text ?? "") : `<${String(child.attrs["kind"])}>`))
				out.push(before + text)
			})

			return out
		}

		// What follows it moves before it: the range keeps to the paragraph.
		const after = await sessionOf({ body })

		expect(move(after, 2, 2, 1)).toBe(true)
		expect(move(after, 1, 2, 0)).toBe(true)
		expect(shown(after)).toEqual(["two", "one", "<bookmarkStart>zero", "<bookmarkEnd>three"])

		// The paragraph moves past the next one and back: the document is as it was.
		const back = await sessionOf({ body })
		const original = back.state.doc

		expect(move(back, 0, 0, 2)).toBe(true)
		expect(move(back, 1, 1, 0)).toBe(true)
		expect(back.state.doc.eq(original)).toBe(true)
	})

	it("refuses to join across a section break and moves the break with a split", async () => {
		const session = await sessionOf({
			body:
				p("before") +
				'<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>end of section</w:t></w:r></w:p>' +
				p("after")
		})
		const doc = session.state.doc
		const section = doc.child(0).nodeSize
		const sectPr = doc.child(1).attrs["sectPr"] as number

		expect(sectPr).toBe(doc.child(1).attrs["src"])
		// Backspace at the start of the section's last paragraph, Delete at its end, a range across it.
		expect(run(session, state => state.tr.join(section))).toBe(false)
		expect(run(session, state => state.tr.join(section + doc.child(1).nodeSize))).toBe(false)
		expect(run(session, state => state.tr.delete(section - 2, section + 3))).toBe(false)
		expect(run(session, state => state.tr.delete(section, section + doc.child(1).nodeSize))).toBe(false)
		expect(session.state.doc).toBe(doc)
		// Editing inside it is fine.
		expect(run(session, state => state.tr.insertText("!", section + 1))).toBe(true)

		// Enter in the middle: the second half ends the section.
		expect(run(session, state => state.tr.split(find(state.doc, "of section")))).toBe(true)

		const split = session.state.doc

		expect(split.child(1).attrs["sectPr"]).toBe(NO_SOURCE)
		expect(split.child(2).attrs["sectPr"]).toBe(sectPr)
		expect(split.child(2).textContent).toBe("of section")
		expect(sectionHolders(split)).toEqual(sectionHolders(doc))

		// Undo restores the paragraph, break included.
		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(session.execute({ type: "history.undo" })).toBe(true)
		expect(session.state.doc.eq(doc)).toBe(true)
	})

	it("deletes fields whole and refuses to delete locked blocks that hold structure", async () => {
		const session = await sessionOf(RICH)
		const doc = session.state.doc
		const fieldPos = find(doc, "gone", 4)

		expect(doc.nodeAt(fieldPos)?.type.name).toBe("field")
		expect(run(session, state => state.tr.delete(fieldPos, fieldPos + 1))).toBe(true)
		expect(session.state.doc.child(0).childCount).toBe(doc.child(0).childCount - 1)

		const toc = blockPos(session.state.doc, 3)

		expect(run(session, state => state.tr.delete(toc, toc + 1))).toBe(false)
		expect(run(session, state => state.tr.setSelection(NodeSelection.create(state.doc, toc)).deleteSelection())).toBe(false)

		// The content-locked control and the altChunk may go as a whole.
		const control = blockPos(session.state.doc, 6)

		expect(session.state.doc.child(6).attrs["reason"]).toBe("contentLocked")
		expect(run(session, state => state.tr.delete(control, control + 1))).toBe(true)
		expect(session.state.doc.child(6).attrs["reason"]).toBe("altChunk")
	})

	it("refuses to delete a content control locked against deletion", async () => {
		const session = await sessionOf(RICH)
		const doc = session.state.doc
		const control = blockPos(doc, 5)

		expect(run(session, state => state.tr.delete(control, control + doc.child(5).nodeSize))).toBe(false)
		// Its content stays editable.
		expect(run(session, state => state.tr.insertText("x", find(state.doc, "in control")))).toBe(true)
	})

	it("keeps deleted text read-only and new text out of others' revisions", async () => {
		const session = await sessionOf(RICH)

		expect(run(session, state => state.tr.insertText("x", find(state.doc, "gone", 2)))).toBe(false)

		const inside = find(session.state.doc, "inserted", 3)

		expect(run(session, state => state.tr.setSelection(TextSelection.create(state.doc, inside)).insertText("NEW"))).toBe(true)

		const paragraph = session.state.doc.child(0)
		const pieces: [string, string[]][] = []

		paragraph.forEach(child => {
			if (child.isText) pieces.push([child.text ?? "", child.marks.map(mark => mark.type.name)])
		})

		// The new text is in the run it was typed into, outside the revision.
		expect(pieces).toContainEqual(["NEW", ["run"]])
		expect(pieces).toContainEqual(["ins", ["insertion", "run"]])
		expect(pieces).toContainEqual(["erted", ["insertion", "run"]])
	})

	it("refuses every change in read-only mode", async () => {
		const session = await sessionOf(RICH, { readOnly: true })
		const doc = session.state.doc

		expect(run(session, state => state.tr.insertText("x", 2))).toBe(false)
		expect(run(session, state => state.tr.setSelection(TextSelection.create(state.doc, 3)))).toBe(true)
		expect(session.state.doc).toBe(doc)
		expect(session.state.selection.from).toBe(3)
	})

	it("strips source identity from pasted content", async () => {
		const session = await sessionOf(RICH)
		const doc = session.state.doc
		const slice = freshSlice(doc.slice(0, blockPos(doc, 4)))
		const seen: string[] = []

		slice.content.descendants(node => {
			seen.push(node.type.name)

			if ("src" in node.attrs) expect(node.attrs["src"], node.type.name).toBe(NO_SOURCE)
			if ("sectPr" in node.attrs) expect(node.attrs["sectPr"]).toBe(NO_SOURCE)
			if ("before" in node.attrs) expect(node.attrs["before"]).toEqual([])

			for (const mark of node.marks) {
				expect(["insertion", "deletion", "content_control", "wrapper", "format_change"]).not.toContain(mark.type.name)

				if (mark.type.name === "link") expect(mark.attrs["src"]).toBe(NO_SOURCE)
			}

			return true
		})

		expect(seen).not.toContain("marker")
		expect(seen).not.toContain("comment_ref")
		expect(seen).not.toContain("math")
		expect(seen).toContain("field")
		// The TOC paragraph comes back as its text.
		expect(slice.content.lastChild?.textContent).toBe("Entry 1")
		// Pasted content keeps formatting and goes in.
		const end = session.state.doc.content.size

		expect(run(session, state => state.tr.replace(end, end, new Slice(slice.content, 0, 0)))).toBe(true)
		expect(markerKeys(session.state.doc)).toEqual(markerKeys(doc))
	})

	it("normalises list attributes", () => {
		expect(normalizeListAttrs({ numId: null, ilvl: 3 })).toEqual({ numId: null, ilvl: null })
		expect(normalizeListAttrs({ numId: 4, ilvl: 12 })).toEqual({ numId: 4, ilvl: 8 })
		expect(normalizeListAttrs({ numId: 4, ilvl: -1 })).toEqual({ numId: 4, ilvl: 0 })
		expect(normalizeListAttrs({ numId: 4, ilvl: 2 })).toBeUndefined()
	})

	it("runs list normalisation hooks on changed paragraphs", async () => {
		const session = await sessionOf(
			{ body: p("a") + p("b") },
			{ normalizeList: attrs => (attrs.numId === 7 ? { numId: 8, ilvl: attrs.ilvl } : undefined) }
		)

		expect(run(session, state => state.tr.setNodeAttribute(0, "numId", 7).setNodeAttribute(0, "ilvl", 20))).toBe(true)
		expect(session.state.doc.child(0).attrs).toMatchObject({ numId: 8, ilvl: 8 })
		expect(session.state.doc.child(1).attrs).toMatchObject({ numId: null, ilvl: null })
	})

	it("holds every invariant under random deletions, joins, splits and inserts", { timeout: 120_000 }, async () => {
		for (let seed = 1; seed <= Number(process.env["FUZZ_SEEDS"] ?? 12); seed++) {
			const session = await sessionOf(RICH, { historyDepth: 10_000 })
			const next = random(seed)
			const opened = session.state.doc
			const markers = markerKeys(session.state.doc)
			const sections = sectionHolders(session.state.doc)
			let applied = 0

			for (let round = 0; round < Number(process.env["FUZZ_ROUNDS"] ?? 150); round++) {
				const state = session.state
				const size = state.doc.content.size
				const a = Math.floor(next() * (size + 1))
				const b = Math.min(size, a + Math.floor(next() * (next() < 0.8 ? 40 : 400)))
				const choice = next()
				let tr: Transaction | null = null

				try {
					if (choice < 0.45) {
						tr = state.tr.delete(a, b)
					} else if (choice < 0.6) {
						if (canJoin(state.doc, a)) tr = state.tr.join(a)
					} else if (choice < 0.75) {
						if (state.doc.resolve(a).parent.inlineContent && canSplit(state.doc, a)) tr = state.tr.split(a)
					} else if (choice < 0.9) {
						if (state.doc.resolve(a).parent.inlineContent) tr = state.tr.insertText("xy", a)
					} else if (choice < 0.95) {
						tr = state.tr.replace(a, b, new Slice(Fragment.from(docxSchema.text("pasted")), 0, 0))
					} else {
						// Copy and paste inside the document.
						const c = Math.floor(next() * (size + 1))

						tr = state.tr.replace(a, b, freshSlice(state.doc.slice(c, Math.min(size, c + Math.floor(next() * 300)))))
					}
				} catch {
					tr = null
				}

				if (tr === null || !tr.docChanged) continue

				const previous = session.state.doc

				if (session.dispatch(tr)) applied++

				if (process.env["FUZZ_DEBUG"] !== undefined && markerKeys(session.state.doc).length !== markers.length) {
					console.log(
						seed,
						round,
						JSON.stringify(tr.steps.map(step => step.toJSON())),
						"\nBEFORE",
						previous.toString().slice(0, 3000),
						"\nAFTER",
						session.state.doc.toString().slice(0, 3000)
					)
				}

				const doc = session.state.doc

				doc.check()
				// Every marker stays, except comments taken whole (start, end and reference together).
				const keys = markerKeys(doc)
				const comments = new Set(keys.filter(key => key.startsWith("comment")).map(key => key.slice(key.indexOf(":"))))

				expect(keys, `seed ${seed} round ${round}`).toEqual(
					markers.filter(key => !key.startsWith("comment") || comments.has(key.slice(key.indexOf(":"))))
				)
				expect(sectionHolders(doc), `seed ${seed} round ${round}`).toEqual(sections)

				let lockedDeletable = 0

				doc.descendants(node => {
					if (node.type.name === "locked_block" && node.attrs["deletable"] !== true) lockedDeletable++

					return !node.isTextblock
				})

				expect(lockedDeletable).toBe(2)
			}

			expect(applied).toBeGreaterThan(20)

			// Everything undoes back to the opened document.
			while (session.execute({ type: "history.undo" })) {
				// undo all
			}

			expect(session.state.doc.eq(opened)).toBe(true)
			expect(session.snapshot().dirty).toBe(false)
		}
	})

	it("puts a comment range a block move turned around back in order, collapsed where its end went", async () => {
		const opened = await openSpec({
			body:
				'<w:p><w:commentRangeStart w:id="9"/><w:r><w:t>Start</w:t></w:r></w:p>' +
				p("Middle") +
				'<w:commentRangeEnd w:id="9"/><w:p><w:r><w:commentReference w:id="9"/></w:r><w:r><w:t>End</w:t></w:r></w:p>' +
				p("After"),
			comments: '<w:comment w:id="9" w:author="Bo"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment>'
		})
		const session = new DocxEditorSession(opened.context)
		const doc = session.state.doc
		const from = doc.child(0).nodeSize + doc.child(1).nodeSize
		const to = from + doc.child(2).nodeSize
		const tr = session.state.tr.delete(from, to)

		// The paragraph holding the end (in its marker list) and the reference dragged to the top.
		tr.replaceRange(0, 0, doc.slice(from, to))
		expect(session.dispatch(tr)).toBe(true)

		const order: string[] = []

		session.state.doc.descendants(node => {
			for (const marker of survivorsOf(node)) if (marker.id === "9") order.push(marker.kind)
			if (node.isBlock)
				for (const marker of (node.attrs["before"] ?? []) as MarkerList) if (marker.id === "9") order.push(marker.kind)

			return true
		})
		expect(order).toEqual(["commentRangeStart", "commentRangeEnd", "commentReference"])
		await expect(saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })).resolves.toBeDefined()
	})

	it("keeps the row and cell properties of a pasted copy of a source row", async () => {
		const cell = (text: string): string => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`
		const opened = await openSpec({
			body:
				'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
				`<w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:w="1000" w:type="dxa"/></w:trPr>${cell("one")}</w:tr></w:tbl>` +
				p("after")
		})
		const session = new DocxEditorSession(opened.context)
		const table = session.state.doc.child(0)
		const row = table.child(0)

		// The row pasted again below itself.
		expect(session.dispatch(session.state.tr.insert(1 + row.nodeSize, row))).toBe(true)

		const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })
		const xml = new TextDecoder().decode(unzipSync(result.bytes as Uint8Array)["word/document.xml"])

		expect(xml.match(/<w:gridBefore w:val="1"\/>/g)).toHaveLength(2)
		expect(xml.match(/<w:tcW w:w="3000" w:type="dxa"\/>/g)).toHaveLength(2)
	})
})
