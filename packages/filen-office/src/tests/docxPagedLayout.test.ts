import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { refPart, type DocxPackage } from "@filen/office/core/docx"
import { docxSchema, type DocxEditContext, type DocxEditorSession } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { saveDocx } from "@filen/office/core/docx/write"
import { MeasureCache, StyleTable } from "@filen/office/core/text"
import { DecorationTable, RunStyles } from "../core/docx/layout/styles"
import { buildDocx } from "./docxSynthetic"
import { readBytes } from "./layoutFixtures"
import { editedLayout, finishLayout, pageDigest, pagedEditor, savedDigest, settleAll, type EditedLayout } from "./pagedFixtures"
import { harness } from "./textFixtures"

// Edits laid out incrementally show exactly what a fresh layout of the saved file shows.

const FIXTURES = new URL("./fixtures/", import.meta.url).pathname

function paragraphs(doc: Node, inTables: boolean, minimum = 8): { node: Node; pos: number }[] {
	const out: { node: Node; pos: number }[] = []

	doc.descendants((node, pos, parent) => {
		if (node.type === docxSchema.nodes.paragraph) {
			const inCell = parent?.type === docxSchema.nodes.table_cell

			if (inCell === inTables && node.textContent.length >= minimum && node.firstChild?.isText === true) out.push({ node, pos })

			return false
		}

		return !node.isInline
	})

	return out
}

function select(session: DocxEditorSession, from: number, to: number = from): void {
	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, from, to)))
}

async function expectSaved(edited: EditedLayout): Promise<void> {
	const saved = await savedDigest(await edited.save())

	expect(pageDigest(edited.layout)).toEqual(saved)
}

const words = (count: number, seed = 1): string =>
	Array.from({ length: count }, (_, k) => ["alpha", "beta", "gamma", "delta", "epsilon", "zeta"][(k * 7 + seed) % 6]).join(" ")

function longDocument(paragraphCount: number): Uint8Array {
	const p = (text: string, pPr = ""): string =>
		`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
	const body =
		Array.from({ length: paragraphCount }, (_, k) =>
			k % 17 === 9
				? `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${Array.from(
						{ length: 3 },
						(_, r) =>
							`<w:tr><w:tc><w:p><w:r><w:t>cell ${r} ${words(6, k + r)}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>${words(4, r)}</w:t></w:r></w:p></w:tc></w:tr>`
					).join("")}</w:tbl>`
				: k % 23 === 11
					? p(words(20, k), "<w:pageBreakBefore/>")
					: p(words(30 + (k % 7) * 9, k), k % 13 === 0 ? "<w:keepNext/>" : "")
		).join("") +
		'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

	return buildDocx({ body })
}

describe("paged layout of edits", { timeout: 60_000 }, () => {
	it("lays out typing, deletion and formatting like the saved file, stopping at the next page", async () => {
		const edited = await editedLayout(longDocument(160))
		const pages = edited.layout.pages.length

		expect(pages).toBeGreaterThan(12)

		const target = paragraphs(edited.session.state.doc, false)[40]

		if (target === undefined) throw new Error("no paragraph")

		await edited.edit(edited.session.state.tr.insertText("x", target.pos + 3))

		// One character: the page is laid out again, the rest is kept.
		expect(edited.last.laidOut).toBeLessThanOrEqual(2)
		expect(edited.layout.pages.length).toBe(pages)
		await expectSaved(edited)

		const other = paragraphs(edited.session.state.doc, false)[60]

		if (other === undefined) throw new Error("no paragraph")

		await edited.edit(edited.session.state.tr.delete(other.pos + 2, other.pos + 9))
		select(edited.session, other.pos + 2, other.pos + 8)
		await edited.edit(session => session.execute({ type: "format.bold" }))
		await expectSaved(edited)
	})

	it("follows paragraphs added, split, joined and removed", async () => {
		const edited = await editedLayout(longDocument(120))
		const doc = (): Node => edited.session.state.doc
		const list = paragraphs(doc(), false)
		const split = list[10]

		if (split === undefined) throw new Error("no paragraph")

		await edited.edit(edited.session.state.tr.split(split.pos + 6))
		await expectSaved(edited)

		const again = paragraphs(doc(), false)
		const first = again[30]
		const second = again[31]

		if (first !== undefined && second !== undefined && first.pos + first.node.nodeSize === second.pos) {
			await edited.edit(edited.session.state.tr.join(second.pos))
			await expectSaved(edited)
		}

		const removed = paragraphs(doc(), false)[50]

		if (removed === undefined) throw new Error("no paragraph")

		await edited.edit(edited.session.state.tr.delete(removed.pos, removed.pos + removed.node.nodeSize))
		await expectSaved(edited)

		// Several paragraphs typed in a row (Enter after Enter).
		for (let k = 0; k < 6; k++) {
			const at = paragraphs(doc(), false)[70]

			if (at === undefined) break

			await edited.edit(edited.session.state.tr.split(at.pos + 4).insertText("new line ", at.pos + 6))
		}

		await expectSaved(edited)
	})

	it("edits table cells without laying the table out from an overlay", async () => {
		const edited = await editedLayout(longDocument(80))
		const cell = paragraphs(edited.session.state.doc, true, 4)[2]

		if (cell === undefined) throw new Error("no cell")

		await edited.edit(edited.session.state.tr.insertText("typed ", cell.pos + 1))

		// The table keeps its source; the paragraph is substituted.
		expect([...edited.edits.overlayRecords()].every(record => record.node.type === docxSchema.nodes.paragraph)).toBe(true)
		await expectSaved(edited)

		// A new paragraph in a cell changes the table's structure: the table is laid out from an overlay.
		const next = paragraphs(edited.session.state.doc, true, 4)[2]

		if (next === undefined) throw new Error("no cell")

		await edited.edit(edited.session.state.tr.split(next.pos + 3))

		const table = [...edited.edits.overlayRecords()].find(record => record.node.type === docxSchema.nodes.table)

		expect(table).toBeDefined()
		await expectSaved(edited)

		// Typing in its cells afterwards, the new paragraph included, lays out just those paragraphs from overlays
		// over the table's own.
		for (const index of [2, 3, 5]) {
			const at = paragraphs(edited.session.state.doc, true, 1)[index]

			if (at === undefined) throw new Error("no cell")

			await edited.edit(edited.session.state.tr.insertText("more ", at.pos + 1))
		}

		const records = [...edited.edits.overlayRecords()]

		expect(records.filter(record => record.node.type === docxSchema.nodes.table)).toEqual([table])
		expect(records.filter(record => record.node.type === docxSchema.nodes.paragraph).length).toBeGreaterThanOrEqual(3)
		expect(records.some(record => record.node.type === docxSchema.nodes.paragraph && refPart(record.base) === table?.part.index)).toBe(
			true
		)
		await expectSaved(edited)
	})

	it("writes the blocks one edit changes into one overlay part", async () => {
		const edited = await editedLayout(longDocument(60))
		const session = edited.session

		select(session, 1, session.state.doc.content.size - 1)
		await edited.edit(s => s.execute({ type: "format.bold" }))

		const records = [...edited.edits.overlayRecords()]

		expect(records.length).toBeGreaterThan(40)
		expect(new Set(records.map(record => record.part)).size).toBe(1)
		await expectSaved(edited)

		// Typing in one of them gives it a part of its own; the others stay.
		const target = paragraphs(session.state.doc, false)[10]

		if (target === undefined) throw new Error("no paragraph")

		await edited.edit(session.state.tr.insertText("x", target.pos + 2))
		expect(new Set([...edited.edits.overlayRecords()].map(record => record.part)).size).toBe(2)
		await expectSaved(edited)
	})

	it("undoes back to the source and keeps every page as laid out before", async () => {
		const edited = await editedLayout(longDocument(60))
		const before = pageDigest(edited.layout)
		const target = paragraphs(edited.session.state.doc, false)[12]

		if (target === undefined) throw new Error("no paragraph")

		await edited.edit(edited.session.state.tr.insertText("inserted words that wrap ", target.pos + 3))
		await edited.edit(edited.session.state.tr.split(target.pos + 10))
		await edited.edit(session => session.execute({ type: "history.undo" }))
		await edited.edit(session => session.execute({ type: "history.undo" }))

		expect(pageDigest(edited.layout)).toEqual(before)
	})

	it("keeps later pages when an edit adds a page before a page break, renumbered", async () => {
		const field = (instruction: string): string =>
			`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
		const p = (text: string, pPr = ""): string =>
			`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		// Chapters of three short paragraphs, each starting on a new page.
		const body =
			Array.from({ length: 24 }, (_, chapter) =>
				[
					p(`Chapter ${chapter} ${words(10, chapter)}`, chapter === 0 ? "" : "<w:pageBreakBefore/>"),
					p(words(40, chapter)),
					p(words(40, chapter + 1))
				].join("")
			).join("") +
			'<w:sectPr><w:footerReference w:type="default" r:id="rIdF1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(
			buildDocx({
				body,
				footers: {
					rIdF1: `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field("PAGE")}<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field("NUMPAGES")}</w:p>`
				}
			})
		)
		const pages = edited.layout.pages.length

		expect(pages).toBe(24)

		const target = paragraphs(edited.session.state.doc, false)[3 * 5 + 1]

		if (target === undefined) throw new Error("no paragraph")

		// Enough text to push the chapter onto a second page.
		await edited.edit(edited.session.state.tr.insertText(`${words(900, 3)} `, target.pos + 2))

		expect(edited.layout.pages.length).toBe(pages + 1)
		// Laid out up to the next chapter only; the rest was moved and renumbered.
		expect(edited.last.resyncedAt).toBeGreaterThan(edited.last.restart)
		expect(edited.last.resyncedAt).toBeLessThan(edited.last.restart + 4)
		await expectSaved(edited)
	})

	for (const name of ["01-plain.docx", "03-header-footer-footnotes.docx", "05-fields-toc.docx"]) {
		it(`lays out a scripted edit suite of ${name} like the saved file`, async () => {
			const edited = await editedLayout(readBytes(FIXTURES + name))
			const session = edited.session
			const pick = (index: number, inTables = false): { node: Node; pos: number } | undefined => {
				const list = paragraphs(session.state.doc, inTables)

				return list[Math.min(list.length - 1, index)]
			}
			const type = pick(0)

			if (type !== undefined) await edited.edit(session.state.tr.insertText("Filen ", type.pos + 1))

			const remove = pick(1)

			if (remove !== undefined) await edited.edit(session.state.tr.delete(remove.pos + 2, remove.pos + 5))

			const bold = pick(2)

			if (bold !== undefined) {
				select(session, bold.pos + 1, bold.pos + 5)
				await edited.edit(s => s.execute({ type: "format.bold" }))
			}

			const heading = pick(3)

			if (heading !== undefined) {
				select(session, heading.pos + 1)
				await edited.edit(s => s.execute({ type: "style.set", style: "heading2" }))
			}

			const list = pick(4)

			if (list !== undefined) {
				select(session, list.pos + 1)
				await edited.edit(s => s.execute({ type: "list.bullet" }))
			}

			const numbered = pick(6)

			if (numbered !== undefined) {
				select(session, numbered.pos + 1)
				await edited.edit(s => s.execute({ type: "list.numbered" }))
			}

			const quote = pick(7)

			if (quote !== undefined) {
				select(session, quote.pos + 1)
				await edited.edit(s => s.execute({ type: "style.set", style: "quote" }))
			}

			const linked = pick(8)

			if (linked !== undefined) {
				await edited.edit(
					session.state.tr.addMark(linked.pos + 1, linked.pos + 4, docxSchema.marks.link.create({ target: "https://filen.io/" }))
				)

				// The link is laid out as one (its target resolves) before it is saved.
				expect(
					edited.layout.pages.some(page =>
						page.sourceParts.some(part =>
							edited.edits.part(part)?.relationships.all.some(r => r.target === "https://filen.io/")
						)
					)
				).toBe(true)
			}

			const split = pick(5)

			if (split !== undefined) await edited.edit(session.state.tr.split(split.pos + 4))

			const cell = pick(0, true)

			if (cell !== undefined) await edited.edit(session.state.tr.insertText("cell ", cell.pos + 1))

			await expectSaved(edited)
		})
	}

	it("edits a header on every page that shows it and goes back to the body", async () => {
		const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const body =
			Array.from({ length: 140 }, (_, k) => p(`Paragraph ${k} ${words(40, k)}`)).join("") +
			'<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const { editor, pkg, context } = await pagedEditor(buildDocx({ body, headers: { rIdH1: p("Report header") } }))
		const pages = editor.pageCount

		expect(pages).toBeGreaterThan(5)

		// The header shows above the body on page 3.
		const band = editor.bodyBand(3)
		const story = editor.storyAt(3, 300, (band?.top ?? 100) - 30)

		expect(story).toBeGreaterThanOrEqual(0)
		expect(editor.enterStory(story, 3)).toBe(true)
		expect(editor.activeStory).toBe(story)

		const hit = editor.hitTest(3, 400, (band?.top ?? 100) - 30)

		if (hit === undefined) throw new Error("no hit in the header")

		editor.select(hit)
		expect(editor.caret()?.page).toBe(3)
		editor.dispatch(editor.session.state.tr.insertText(" (draft)"))
		await settleAll(editor)

		// Every page shows the new header; the body is as it was.
		for (let index = 0; index < editor.pageCount; index++) {
			const page = editor.page(index)

			expect(page?.sourceTexts.some(text => text.includes("Report header (draft)"))).toBe(true)
		}

		expect(editor.pageCount).toBe(pages)

		const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

		if (saved.bytes === undefined) throw new Error("no bytes")

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved.bytes))

		editor.exitStory()
		expect(editor.activeStory).toBe(editor.body.session.story.story.index)
	})

	it("counts fonts still loading as missing, so layout waits for them", async () => {
		let release = (): void => undefined
		const prepared: string[] = []
		const fonts = {
			prepare: (requests: Iterable<{ family: string }>) =>
				new Promise<void>(resolve => {
					for (const request of requests) prepared.push(request.family)

					release = resolve
				}),
			resolve: () => {
				throw new Error("not resolved here")
			}
		}
		const styles = new RunStyles(new StyleTable(new MeasureCache(harness().measurer)), fonts, new DecorationTable())
		const request = { family: "Lato", bold: false, italic: false, slot: "ascii" as const }
		const loading = styles.prepare([request])

		expect(styles.hasMissing).toBe(true)

		// Another wait for the same fonts waits for the load under way instead of passing at once.
		let done = false
		const waiting = styles.prepareMissing().then(() => {
			done = true
		})

		await Promise.resolve()
		expect(done).toBe(false)

		release()
		await loading
		await waiting
		expect(done).toBe(true)
		expect(styles.hasMissing).toBe(false)
		expect(prepared).toEqual(["Lato"])
	})

	it("meets the earlier layout again when an edit comes while another is still being laid out", async () => {
		const bytes = readBytes(`${FIXTURES}08-large-300p.docx`)
		const saveOf = async (editor: PagedEditor, pkg: DocxPackage, context: DocxEditContext): Promise<string[]> => {
			const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

			if (saved.bytes === undefined) throw new Error("no bytes")

			return savedDigest(saved.bytes)
		}

		for (const variant of ["delete", "paste", "paste, laid out"] as const) {
			// Edits lay out the pages in view in their own task, the rest later: nothing runs in between here.
			const { editor, pkg, context } = await pagedEditor(bytes, 0)
			const session = editor.session
			const onPage = (page: number): number => {
				const found = paragraphs(session.state.doc, false, 20).find(at => editor.geometry.locate(at.pos + 2, 1)?.page === page)

				if (found === undefined) throw new Error(`no paragraph on page ${String(page)}`)

				return found.pos + 2
			}

			editor.setVisible(20, 25)

			const from = onPage(20)

			if (variant === "delete") {
				editor.dispatch(session.state.tr.delete(from, onPage(23)))
			} else {
				editor.dispatch(session.state.tr.replace(from, from, session.state.doc.slice(onPage(40), onPage(43))))
			}

			if (variant === "paste, laid out") {
				// Laid out to the end, with later pages' page numbers not brought up to date yet (out of view).
				await finishLayout(editor.layout)

				expect(editor.layout.stale.size).toBeGreaterThan(0)
			} else {
				expect(editor.layout.final).toBe(false)
			}

			// Before that layout gets further (or its pages are brought up to date): an edit on an earlier page.
			editor.dispatch(session.state.tr.split(onPage(10)))
			await settleAll(editor)

			expect(pageDigest(editor.layout)).toEqual(await saveOf(editor, pkg, context))
			editor.dispose()
		}
	}, 60_000)

	it("renumbers pages of later sections up to one numbering its pages anew", async () => {
		const field =
			'<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
		const p = (text: string, pPr = ""): string =>
			`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const sectPr = (extra = ""): string =>
			`<w:sectPr><w:footerReference w:type="default" r:id="rIdF1"/>${extra}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
		const section = (seed: number, extra?: string): string =>
			Array.from({ length: 12 }, (_, k) => p(words(60, seed + k), k % 4 === 0 && k > 0 ? "<w:pageBreakBefore/>" : "")).join("") +
			(extra === undefined ? "" : `<w:p><w:pPr>${sectPr(extra)}</w:pPr></w:p>`)
		const body = section(0, "") + section(20, "") + section(40, '<w:pgNumType w:start="1"/>') + section(60) + sectPr()
		const edited = await editedLayout(buildDocx({ body, footers: { rIdF1: field } }))
		const target = paragraphs(edited.session.state.doc, false)[2]

		if (target === undefined) throw new Error("no paragraph")

		// A page more in the first section: the second section's pages move on; the third numbers its pages anew,
		// so neither it nor the fourth moves.
		await edited.edit(edited.session.state.tr.insertText(`${words(900, 5)} `, target.pos + 2))
		expect(edited.last.resyncedAt).toBeGreaterThan(0)
		expect(edited.layout.pages.map(page => page.number)).toEqual([1, 2, 3, 4, 5, 6, 7, 1, 2, 3, 4, 5, 6])
		await expectSaved(edited)
	})

	it("lets a shortened paragraph opening a page move back to the page before", async () => {
		const p = (text: string): string => `<w:p><w:pPr><w:widowControl/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const body =
			Array.from({ length: 60 }, (_, k) => p(words(20 + ((k * 13) % 7) * 6, k))).join("") +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const bytes = buildDocx({ body })
		// Paragraphs opening their page.
		const probe = await editedLayout(bytes)
		const pages = (index: number): readonly number[] => {
			const entry = probe.story.entries[index]

			return entry === undefined ? [] : probe.index.pagesOf(entry)
		}
		const opening = Array.from({ length: 59 }, (_, k) => k + 1).filter(
			block => pages(block)[0] !== undefined && pages(block)[0] !== pages(block - 1).at(-1)
		)

		expect(opening.length).toBeGreaterThanOrEqual(2)

		for (const block of opening.slice(0, 2)) {
			const edited = await editedLayout(bytes)
			const node = edited.session.state.doc.child(block)
			const start = edited.story.starts[block] ?? 0

			await edited.edit(edited.session.state.tr.delete(start + 1 + 10, start + node.nodeSize - 1))
			await expectSaved(edited)
		}
	}, 30_000)

	it("enters only the document's own stories, not those of overlays", async () => {
		const { editor } = await pagedEditor(longDocument(30))
		const own = editor.edits.storyCount
		const target = paragraphs(editor.session.state.doc, false)[3]

		if (target === undefined) throw new Error("no paragraph")

		editor.dispatch(editor.session.state.tr.insertText("typed ", target.pos + 2))
		await settleAll(editor)

		const stories = editor.edits.model.stories

		expect(stories.length).toBeGreaterThan(own)
		expect(stories[own]?.kind).toBe("body")
		expect(editor.enterStory(own, 0)).toBe(false)
		expect(editor.activeStory).toBe(editor.body.session.story.story.index)
		editor.dispose()
	})

	it("moves section boundaries with the section properties a split hands to the second half", async () => {
		const p = (text: string, pPr = ""): string =>
			`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const geometry =
			'<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'
		const body =
			p(words(12, 1), `<w:sectPr><w:type w:val="continuous"/>${geometry}</w:sectPr>`) +
			p(words(8, 2)) +
			`<w:sectPr>${geometry}<w:pgNumType w:start="1"/></w:sectPr>`
		const edited = await editedLayout(buildDocx({ body }))
		const first = paragraphs(edited.session.state.doc, false)[0]

		if (first === undefined) throw new Error("no paragraph")

		// The split keeps the section properties with the second half.
		await edited.edit(edited.session.state.tr.split(first.pos + 20))
		await expectSaved(edited)

		const second = paragraphs(edited.session.state.doc, false)[2]

		if (second === undefined) throw new Error("no paragraph")

		select(edited.session, second.pos + 2)
		await edited.edit(session => session.execute({ type: "align.center" }))
		await expectSaved(edited)
	})

	it("keeps the numbers of a section numbering its pages anew where layout meets the earlier one", async () => {
		const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const field =
			'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
		const sectPr = (extra = ""): string =>
			`<w:sectPr><w:footerReference w:type="default" r:id="rIdF1"/>${extra}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
		const section = (seed: number): string => Array.from({ length: 12 }, (_, k) => p(words(60, seed + k))).join("")
		const body = section(0) + `<w:p><w:pPr>${sectPr()}</w:pPr></w:p>` + section(20) + sectPr('<w:pgNumType w:start="0"/>')
		const edited = await editedLayout(buildDocx({ body, footers: { rIdF1: field } }))
		const target = paragraphs(edited.session.state.doc, false)[2]

		if (target === undefined) throw new Error("no paragraph")

		// A page more in the first section; layout meets the earlier layout where the second section starts.
		await edited.edit(edited.session.state.tr.insertText(`${words(400, 5)} `, target.pos + 2))
		expect(edited.last.resyncedAt).toBeGreaterThan(0)
		expect(edited.layout.pages.at(-1)?.number).toBe(edited.layout.pages.filter(page => page.section === 1).length - 1)
		await expectSaved(edited)
	})

	it("counts lists a document gains anew whenever layout restarts before them", async () => {
		const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const body =
			Array.from({ length: 60 }, (_, k) => p(words(40, k))).join("") +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(buildDocx({ body }))

		// Layout restarts from snapshots taken before the document had lists.
		for (const index of [40, 41, 25]) {
			const at = paragraphs(edited.session.state.doc, false)[index]

			if (at === undefined) throw new Error("no paragraph")

			select(edited.session, at.pos + 2)
			await edited.edit(session => session.execute({ type: "list.numbered" }))
			await expectSaved(edited)
		}
	})

	it("lays out a header edited again while layout is still under way", async () => {
		const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const body =
			Array.from({ length: 140 }, (_, k) => p(`Paragraph ${String(k)} ${words(40, k)}`)).join("") +
			'<w:sectPr><w:headerReference w:type="first" r:id="rIdH0"/><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:titlePg/></w:sectPr>'
		const { editor, pkg, context } = await pagedEditor(
			buildDocx({ body, headers: { rIdH0: p("Title page header"), rIdH1: p("Report header") } }),
			0
		)
		const band = editor.bodyBand(1)
		const story = editor.storyAt(1, 300, (band?.top ?? 100) - 30)

		editor.setVisible(1, 1)
		expect(editor.enterStory(story, 1)).toBe(true)

		// A longer header (from page 1 on) moves every body down: layout runs from page 1 on, in slices.
		const end = editor.session.state.doc.content.size - 1

		editor.dispatch(editor.session.state.tr.insertText(` ${words(60, 3)}`, end))
		expect(editor.layout.final).toBe(false)

		// Before it gets further, the header changes again.
		editor.dispatch(editor.session.state.tr.split(editor.session.state.doc.content.size - 1))
		editor.exitStory()
		await settleAll(editor)

		const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

		if (saved.bytes === undefined) throw new Error("no bytes")

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved.bytes))
		editor.dispose()
	}, 30_000)

	it("draws line numbers in their own font even when nothing else uses it", async () => {
		const cambria = '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/></w:rPr>'
		const body =
			Array.from({ length: 3 }, (_, k) => `<w:p><w:pPr>${cambria}</w:pPr><w:r>${cambria}<w:t>${words(10, k)}</w:t></w:r></w:p>`).join(
				""
			) +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:lnNumType w:countBy="1"/></w:sectPr>'
		const digest = await savedDigest(buildDocx({ body }))
		const numbers = (digest[0] ?? "").split("\n").filter(line => / 6 "\d+" /.test(line))

		expect(numbers.length).toBe(3)
		expect(numbers.every(line => !line.includes(" null:"))).toBe(true)
	})
})
