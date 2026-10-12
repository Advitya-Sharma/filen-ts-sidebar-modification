import { unzipSync } from "fflate"
import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { Transform } from "prosemirror-transform"
import { describe, expect, it } from "vitest"
import { refPart, type DocxPackage } from "@filen/office/core/docx"
import { docxSchema, type DocxEditContext, type DocxEditorSession } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { saveDocx, StoryWriter } from "@filen/office/core/docx/write"
import { MeasureCache, StyleTable } from "@filen/office/core/text"
import { DecorationTable, RunStyles } from "../core/docx/layout/styles"
import { Item } from "../core/docx/model/index"
import { namePrefixes } from "../core/docx/paged/edits"
import { buildDocx } from "./docxSynthetic"
import { layOut, readBytes } from "./layoutFixtures"
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

const para = (text: string, pPr = ""): string =>
	`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

const PAGE =
	'<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const LANDSCAPE =
	'<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const NARROW =
	'<w:pgSz w:w="8000" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'

// A content control holding `inner`.
const control = (inner: string): string => `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`

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

// Paragraphs with a footnote each, notes of a few lengths.
function notedDocument(): Uint8Array {
	const noted = (k: number): string =>
		`<w:p><w:r><w:t xml:space="preserve">${words(25 + (k % 4) * 11, k)} </w:t></w:r><w:r><w:footnoteReference w:id="${String(k + 1)}"/></w:r><w:r><w:t xml:space="preserve"> ${words(12, k + 2)}</w:t></w:r></w:p>`
	const note = (k: number): string =>
		`<w:footnote w:id="${String(k + 1)}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${words(8 + (k % 3) * 20, k)}</w:t></w:r></w:p></w:footnote>`

	return buildDocx({
		body: Array.from({ length: 40 }, (_, k) => noted(k)).join("") + `<w:sectPr>${PAGE}</w:sectPr>`,
		footnotes:
			'<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
			Array.from({ length: 40 }, (_, k) => note(k)).join(""),
		settings: '<w:footnotePr><w:footnote w:id="-1"/></w:footnotePr>'
	})
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

	it("lays out a paragraph command from the paragraph's page, not the whole document", async () => {
		const edited = await editedLayout(longDocument(160))
		const target = paragraphs(edited.session.state.doc, false)[100]

		if (target === undefined) throw new Error("no paragraph")

		select(edited.session, target.pos + 2)
		await edited.edit(session => session.execute({ type: "align.center" }))

		expect(edited.last.restart).toBeGreaterThan(5)
		expect(edited.last.laidOut).toBeLessThanOrEqual(3)
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

	it("lays table paragraphs out from their source again when a paragraph command over them is undone", async () => {
		const edited = await editedLayout(longDocument(60))
		const session = edited.session
		const before = pageDigest(edited.layout)

		select(session, 1, session.state.doc.content.size - 1)
		await edited.edit(s => s.execute({ type: "align.center" }))
		expect([...edited.edits.overlayRecords()].length).toBeGreaterThan(40)
		await edited.edit(s => s.execute({ type: "history.undo" }))

		// No overlay left: every paragraph, those in table cells too, is its source again.
		expect([...edited.edits.overlayRecords()]).toEqual([])
		expect(pageDigest(edited.layout)).toEqual(before)
		await expectSaved(edited)
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

	it("lays a split row out the same whether fonts its nested table needs arrive during it or not", async () => {
		const run = (text: string, font?: string): string =>
			`<w:r><w:rPr>${font === undefined ? "" : `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`}<w:sz w:val="72"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`
		const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
		const row = (...cells: string[]): string => `<w:tr>${cells.join("")}</w:tr>`
		const grid = '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
		// A nested table whose last row needs a font nothing before it uses, and runs over a page.
		const nested =
			`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${grid}` +
			row(cell(`<w:p>${run(words(12, 1))}</w:p>`), cell("<w:p/>")) +
			row(cell(`<w:p>${run(words(30, 2), "Courier New")}</w:p>`), cell("<w:p/>")) +
			"</w:tbl>"
		const body =
			`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${grid}` +
			row(cell(`<w:p>${run(words(30, 3))}</w:p>${nested}<w:p>${run(words(8, 4))}</w:p>`), cell(`<w:p>${run(words(4, 5))}</w:p>`)) +
			"</w:tbl><w:p/>" +
			'<w:sectPr><w:pgSz w:w="11906" w:h="8391"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const layout = await layOut(buildDocx({ body }))
		const first = pageDigest(layout)

		expect(first.some(page => page.includes("fotest-cousine"))).toBe(true)

		// Laid out again with every font loaded: nothing is waited for this time.
		layout.beginEdit(0, undefined)
		await finishLayout(layout)
		expect(pageDigest(layout)).toEqual(first)
	})

	it("adds a section when a paragraph ending one is pasted", async () => {
		const p = (text: string, pPr = ""): string =>
			`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const geometry =
			'<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
		const body =
			p(words(20, 1)) +
			p(words(20, 2), `<w:sectPr>${geometry}<w:titlePg/></w:sectPr>`) +
			Array.from({ length: 8 }, (_, k) => p(words(30, k + 3))).join("") +
			`<w:sectPr>${geometry}</w:sectPr>`
		const edited = await editedLayout(buildDocx({ body }))
		const doc = edited.session.state.doc
		const breakAt = paragraphs(doc, false)[1]
		const target = paragraphs(doc, false)[6]

		if (breakAt === undefined || target === undefined) throw new Error("no paragraph")

		// Text up to the start of the paragraph ending the first section, pasted inside a later paragraph: its end
		// takes that paragraph's properties, section break included (as pasting in Word does).
		const slice = doc.slice(breakAt.pos - 3, breakAt.pos + 4)

		await edited.edit(edited.session.state.tr.replace(target.pos + 3, target.pos + 3, slice))
		expect(edited.edits.model.sections.length).toBe(3)
		await expectSaved(edited)

		await edited.edit(session => session.execute({ type: "history.undo" }))
		expect(edited.edits.model.sections.length).toBe(2)
		await expectSaved(edited)
	})

	it("lays a pasted copy out as the save writes it: the first in the document has the source", async () => {
		// A paragraph whose mark is a tracked deletion: it joins the next one where the mark is written.
		const body =
			`<w:p><w:r><w:t xml:space="preserve">${words(12, 1)}</w:t></w:r></w:p>` +
			'<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Joined with the next </w:t></w:r></w:p>' +
			Array.from({ length: 6 }, (_, k) => `<w:p><w:r><w:t xml:space="preserve">${words(30, k)}</w:t></w:r></w:p>`).join("") +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(buildDocx({ body }))
		const doc = edited.session.state.doc
		const [first, marked] = paragraphs(doc, false)

		if (first === undefined || marked === undefined) throw new Error("no paragraph")

		// Text up to the start of the marked paragraph, pasted into the paragraph before it: the paste ends with
		// a copy of the marked paragraph, now the first holding its source.
		const slice = doc.slice(marked.pos - 3, marked.pos + 4)

		await edited.edit(edited.session.state.tr.replace(first.pos + 5, first.pos + 5, slice))
		await expectSaved(edited)

		// The copy removed again: the original has its source back.
		await edited.edit(session => session.execute({ type: "history.undo" }))
		await expectSaved(edited)
	})

	it("gives a source to one half of a paragraph split inside a content control", async () => {
		const body =
			'<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>' +
			`<w:p><w:r><w:t xml:space="preserve">${words(40, 1)}</w:t></w:r></w:p>` +
			"</w:sdtContent></w:sdt>" +
			Array.from({ length: 6 }, (_, k) => `<w:p><w:r><w:t xml:space="preserve">${words(30, k)}</w:t></w:r></w:p>`).join("") +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(buildDocx({ body }))
		const inside = paragraphs(edited.session.state.doc, false)[0]

		if (inside === undefined) throw new Error("no paragraph")

		await edited.edit(edited.session.state.tr.split(inside.pos + 30))
		await expectSaved(edited)
	})

	it("lays out a table an edit moves and changes as it now reads", async () => {
		const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const table =
			'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
			`<w:tr><w:tc>${p(words(6, 1))}</w:tc><w:tc>${p(words(6, 2))}</w:tc></w:tr></w:tbl>`
		const body =
			p(words(20, 3)) +
			p(words(20, 4)) +
			table +
			p(words(20, 5)) +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(buildDocx({ body }))
		const doc = edited.session.state.doc
		const second = paragraphs(doc, false)[1]
		const cell = paragraphs(doc, true, 4)[0]

		if (second === undefined || cell === undefined) throw new Error("no paragraph")

		// One edit removes the paragraph before the table and types in the table: the table moves up a block.
		const tr = edited.session.state.tr.insertText("typed ", cell.pos + 1)

		tr.delete(second.pos, second.pos + second.node.nodeSize)
		await edited.edit(tr)
		await expectSaved(edited)
	})

	it("lays pages out again rather than keep them on the other side when mirrored margins shift by a page", async () => {
		const p = (text: string, pPr = ""): string =>
			`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const body =
			Array.from({ length: 6 }, (_, chapter) =>
				[p(`Chapter ${String(chapter)}`, chapter === 0 ? "" : "<w:pageBreakBefore/>"), p(words(40, chapter))].join("")
			).join("") +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="4200" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(buildDocx({ body, settings: "<w:mirrorMargins/>" }))
		const target = paragraphs(edited.session.state.doc, false)[1]

		if (target === undefined) throw new Error("no paragraph")

		// One page more before the later chapters: they now fall on pages of the other parity.
		await edited.edit(edited.session.state.tr.insertText(`${words(900, 2)} `, target.pos + 2))
		await expectSaved(edited)
	})

	it("lays out again the page a keep-with-next chain moved off when the chain changes", async () => {
		const p = (text: string, pPr = ""): string =>
			`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const body =
			Array.from({ length: 53 }, (_, k) => p(words(14, k))).join("") +
			p("Kept with the next one", "<w:keepNext/>") +
			p(words(80, 3), "<w:keepLines/>") +
			p(words(30, 4)) +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(buildDocx({ body }))
		const kept = edited.story.entries[53]
		const next = paragraphs(edited.session.state.doc, false)[54]

		if (kept === undefined || next === undefined) throw new Error("no paragraph")

		// The chain did not fit the first page and starts the second; shortened, it fits the first.
		expect(edited.index.pagesOf(kept)).toEqual([1])
		expect(edited.layout.keepFrom(1)).toBe(0)
		await edited.edit(edited.session.state.tr.delete(next.pos + 10, next.pos + next.node.nodeSize - 1))
		await expectSaved(edited)
	})

	it("keeps blocks a paste adds in the section they are in when it also moves a source", async () => {
		const p = (text: string, pPr = ""): string =>
			`<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		const geometry =
			'<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
		const body =
			p(words(20, 1), `<w:sectPr>${geometry}</w:sectPr>`) +
			Array.from({ length: 12 }, (_, k) => p(words(90, k + 2))).join("") +
			`<w:sectPr>${geometry}</w:sectPr>`
		const edited = await editedLayout(buildDocx({ body }))
		const doc = edited.session.state.doc
		const [, into, from] = paragraphs(doc, false)
		const copied = paragraphs(doc, false)[6]

		if (into === undefined || from === undefined || copied === undefined) throw new Error("no paragraph")

		// A paste of several paragraphs ending with a copy of a later one: the copy, now first, has its source.
		await edited.edit(edited.session.state.tr.replace(into.pos + 3, into.pos + 3, doc.slice(from.pos - 3, copied.pos + 4)))
		// Every block after the first section's break is in the second section (dirty blocks carry it).
		expect(edited.story.entries.slice(1).map(entry => entry.section)).toEqual(edited.story.entries.slice(1).map(() => 1))
		await expectSaved(edited)

		const later = paragraphs(edited.session.state.doc, false)[8]

		if (later === undefined) throw new Error("no paragraph")

		await edited.edit(edited.session.state.tr.insertText(`${words(60, 9)} `, later.pos + 2))
		await expectSaved(edited)
	})
	it("lays out again from the page a keep-with-next run moved off when a later page of the run changes", async () => {
		// The first page has no header; later ones a tall one, so less room than the first page has left.
		const header = Array.from({ length: 40 }, (_, k) => para(`Header line ${String(k)}`)).join("")
		const body =
			para(words(12, 1)) +
			para(words(150, 2), "<w:keepNext/>") +
			para(words(250, 3), "<w:keepNext/>") +
			para(words(600, 4), "<w:keepNext/>") +
			para(words(40, 5)) +
			Array.from({ length: 6 }, (_, k) => para(words(60, k + 6))).join("") +
			`<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:headerReference w:type="first" r:id="rIdH2"/>${PAGE}<w:titlePg/></w:sectPr>`
		const edited = await editedLayout(buildDocx({ body, headers: { rIdH1: header, rIdH2: "<w:p/>" } }))
		const last = edited.story.entries[3]
		const target = paragraphs(edited.session.state.doc, false)[3]

		if (last === undefined || target === undefined) throw new Error("no paragraph")

		// The run moved off the first page and the last of it begins pages on, after a chain of its own moved too.
		const page = edited.index.pagesOf(last)[0] ?? -1

		expect(page).toBeGreaterThanOrEqual(3)
		expect(edited.layout.keepFrom(page)).toBe(0)

		// Shortened, the whole run fits the first page.
		await edited.edit(edited.session.state.tr.delete(target.pos + 10, target.pos + target.node.nodeSize - 1))

		const shortened = edited.story.entries[3]

		expect(shortened === undefined ? [] : edited.index.pagesOf(shortened)).toEqual([0])
		await expectSaved(edited)
	})

	it("restarts a keep-with-next document one page back, not from its first page", async () => {
		const body =
			Array.from(
				{ length: 300 },
				(_, k) => para(`Question ${String(k)}`, "<w:keepNext/>") + para(words(40 + ((k * 37) % 90), k), "<w:keepLines/>")
			).join("") + `<w:sectPr>${PAGE}</w:sectPr>`
		const edited = await editedLayout(buildDocx({ body }))
		const answers = paragraphs(edited.session.state.doc, false, 20)
		const target = answers[answers.length - 5]

		if (target === undefined) throw new Error("no paragraph")

		const entry = edited.story.entries[edited.story.entryAt(target.pos + 3)]
		const page = entry === undefined ? -1 : (edited.index.pagesOf(entry)[0] ?? -1)

		expect(page).toBeGreaterThan(20)
		await edited.edit(edited.session.state.tr.insertText("x", target.pos + 3))
		expect(edited.last.restart).toBeGreaterThanOrEqual(page - 2)
		await expectSaved(edited)
	})

	it("lays out a few pages per edit in a document that is one long keep-with-next run", async () => {
		const body =
			para(words(200, 1)) +
			Array.from({ length: 1200 }, (_, k) => para(words(30 + (k % 20), k), "<w:keepNext/>")).join("") +
			para(words(10, 2)) +
			`<w:sectPr>${PAGE}</w:sectPr>`
		const edited = await editedLayout(buildDocx({ body }))
		const pages = edited.layout.pages.length

		expect(pages).toBeGreaterThan(40)

		// The run moved off the first page, where the paragraph it ran out of room at decides that page's end.
		expect(edited.layout.keepFrom(1)).toBe(0)

		for (const at of [0.3, 0.6, 0.9]) {
			const all = paragraphs(edited.session.state.doc, false, 20)
			const target = all[Math.floor(all.length * at)]

			if (target === undefined) throw new Error("no paragraph")

			const entry = edited.story.entries[edited.story.entryAt(target.pos + 2)]
			const page = entry === undefined ? -1 : (edited.index.pagesOf(entry)[0] ?? -1)

			await edited.edit(edited.session.state.tr.insertText("x", target.pos + 2))
			expect(edited.last.restart).toBeGreaterThanOrEqual(page - 2)
			expect(edited.last.laidOut).toBeLessThanOrEqual(4)
		}

		await expectSaved(edited)
	})

	it("numbers pages right when an edit comes while a run laying out pages of the other parity is under way", async () => {
		// Chapters starting pages of their own: pages after an edit start where pages before it did.
		const body =
			Array.from(
				{ length: 30 },
				(_, chapter) =>
					para(`Chapter ${String(chapter)}`, chapter === 0 ? "" : "<w:pageBreakBefore/>") +
					para(words(150, chapter)) +
					para(words(120, chapter + 1))
			).join("") +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="4200" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const { editor, pkg, context } = await pagedEditor(buildDocx({ body, settings: "<w:mirrorMargins/>" }), 0)
		const session = editor.session
		const onPage = (page: number): number => {
			const found = paragraphs(session.state.doc, false, 20).find(at => editor.geometry.locate(at.pos + 2, 1)?.page === page)

			if (found === undefined) throw new Error(`no paragraph on page ${String(page)}`)

			return found.pos + 2
		}
		const pages = editor.layout.pages.length

		editor.setVisible(0, 1)
		// A page more from page 1 on: the pages after it are on the other side, so the run meets none of them.
		editor.dispatch(session.state.tr.insertText(`${words(560, 3)} `, onPage(1)))
		editor.layout.advance(-Infinity, 14)
		expect(editor.layout.final).toBe(false)

		// An edit on a page that run laid out again, before it gets further.
		editor.dispatch(session.state.tr.insertText("x", onPage(5)))
		await settleAll(editor)

		const numbers = editor.layout.pages.map(page => page.number)

		expect(numbers.length % 2).not.toBe(pages % 2)
		expect(numbers).toEqual(numbers.map((_, k) => k + 1))

		const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

		if (saved.bytes === undefined) throw new Error("no bytes")

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved.bytes))
		editor.dispose()
	})

	it("shows the pages a run under way kept when an edit restarts it and meets them again", async () => {
		const body =
			Array.from(
				{ length: 20 },
				(_, chapter) => para(`Chapter ${String(chapter)}`, chapter === 0 ? "" : "<w:pageBreakBefore/>") + para(words(150, chapter))
			).join("") + `<w:sectPr>${PAGE}</w:sectPr>`
		const { editor } = await pagedEditor(buildDocx({ body }), 0)
		const session = editor.session
		const second = paragraphs(session.state.doc, false)[3]

		if (second === undefined) throw new Error("no paragraph")

		editor.setVisible(1, 1)
		// A page more in the second chapter: the run stops just before meeting the third chapter's page, one on.
		editor.dispatch(session.state.tr.insertText(`${words(900, 3)} `, second.pos + 2))
		editor.layout.advance(-Infinity, 2)
		expect(editor.layout.final).toBe(false)

		// Typed there again: that run meets the pages after first, keeping them a page on, and the new run meets
		// those again after the pages in view went out.
		editor.dispatch(session.state.tr.insertText("x", second.pos + 2))
		await settleAll(editor)

		expect(editor.layout.pages.length).toBe(21)
		expect(Array.from({ length: editor.pageCount }, (_, k) => editor.page(k))).toEqual(editor.layout.pages)
		editor.dispose()
	})

	it("numbers a note anew that the run meeting the earlier layout laid out last", async () => {
		const bytes = notedDocument()
		const { editor, pkg, context } = await pagedEditor(bytes)
		const session = editor.session
		const blocks = paragraphs(session.state.doc, false)
		const later = blocks[4]
		const first = blocks[1]

		if (later === undefined || first === undefined) throw new Error("no paragraph")

		// Laid out again to where it meets the earlier layout; then a note before goes: the notes after it are
		// numbered one less, those that run laid out last included.
		editor.dispatch(session.state.tr.insertText("x", later.pos + 2))
		await settleAll(editor)
		editor.dispatch(session.state.tr.delete(first.pos, first.pos + first.node.nodeSize))
		await settleAll(editor)

		const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

		if (saved.bytes === undefined) throw new Error("no bytes")

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved.bytes))
		editor.dispose()
	})

	it("lays a note edited out as it now reads, though the run before laid it out without placing it", async () => {
		const bytes = notedDocument()
		const { editor, pkg, context } = await pagedEditor(bytes)
		const session = editor.session
		const later = paragraphs(session.state.doc, false)[4]
		const note = context.model.footnotes.get("11")

		if (later === undefined || note === undefined) throw new Error("no paragraph")

		// The run meets the earlier layout again after laying out a note it then left for the next page; that
		// note is edited.
		editor.dispatch(session.state.tr.insertText("x", later.pos + 2))
		await settleAll(editor)
		editor.enterStory(note.index, 0)
		editor.session.dispatch(editor.session.state.tr.insertText(`${words(60)} `, 3))
		editor.exitStory()
		await settleAll(editor)

		const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

		if (saved.bytes === undefined) throw new Error("no bytes")

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved.bytes))
		editor.dispose()
	})

	it("lays pages drawn while a font loaded out again once, when the run under way meets the earlier layout first", async () => {
		const body =
			Array.from(
				{ length: 20 },
				(_, chapter) => para(`Chapter ${String(chapter)}`, chapter === 0 ? "" : "<w:pageBreakBefore/>") + para(words(150, chapter))
			).join("") + `<w:sectPr>${PAGE}</w:sectPr>`
		const { editor, pkg, context } = await pagedEditor(buildDocx({ body }), 0)
		const session = editor.session
		const second = paragraphs(session.state.doc, false)[3]

		if (second === undefined) throw new Error("no paragraph")

		editor.setVisible(1, 1)
		// A page more in the second chapter; the run stops between pages before meeting the third chapter's.
		editor.dispatch(session.state.tr.insertText(`${words(900, 3)} `, second.pos + 2))
		editor.layout.advance(-Infinity, 2)
		expect(editor.layout.final).toBe(false)

		// A page it laid out drew stand-ins for a font loading meanwhile (a header, note or split row): the run
		// lays out again from there, meeting the earlier layout where it stands first.
		const paginator = (editor.layout as unknown as { driver: { paginator: { missedFonts: number } } }).driver.paginator

		paginator.missedFonts = 1
		await settleAll(editor)

		const saved = await saveDocx({ pkg, context, sessions: editor.sessions })

		if (saved.bytes === undefined) throw new Error("no bytes")

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(saved.bytes))
		editor.dispose()
	})

	it("ends the sections a content control holds after the blocks that end them", async () => {
		for (const inner of [
			para(words(20, 1), `<w:sectPr>${LANDSCAPE}</w:sectPr>`) + para(words(40, 2)),
			para(words(20, 1), `<w:sectPr>${LANDSCAPE}</w:sectPr>`) + para(words(20, 2), `<w:sectPr>${NARROW}</w:sectPr>`)
		]) {
			const body =
				control(inner) + Array.from({ length: 4 }, (_, k) => para(words(30, k + 3))).join("") + `<w:sectPr>${PAGE}</w:sectPr>`
			const edited = await editedLayout(buildDocx({ body }))
			const target = paragraphs(edited.session.state.doc, false)[4]

			if (target === undefined) throw new Error("no paragraph")

			// A split elsewhere: the layout reads the body's block list from now on.
			await edited.edit(edited.session.state.tr.split(target.pos + 5))
			await expectSaved(edited)
		}
	})

	it("lays a section out with the properties a pasted section break gives it, from the section's start", async () => {
		for (const type of ["", '<w:type w:val="continuous"/>', '<w:type w:val="oddPage"/>']) {
			const body =
				para(words(30, 1)) +
				para(words(20, 2), `<w:sectPr>${type}${PAGE}</w:sectPr>`) +
				Array.from({ length: 30 }, (_, k) => para(words(120, k + 3))).join("") +
				`<w:sectPr>${type}${NARROW}</w:sectPr>`
			const edited = await editedLayout(buildDocx({ body }))
			const doc = edited.session.state.doc
			const breakAt = paragraphs(doc, false)[1]
			const target = paragraphs(doc, false)[25]

			if (breakAt === undefined || target === undefined) throw new Error("no paragraph")

			// The second section's pages before the paste point take the first section's properties.
			await edited.edit(edited.session.state.tr.replace(target.pos + 3, target.pos + 3, doc.slice(breakAt.pos - 3, breakAt.pos + 4)))
			expect(edited.edits.model.sections.length).toBe(3)
			await expectSaved(edited)
		}
	})

	it("knows which section each block of a content control holding a section break is in", async () => {
		const body =
			control(
				para(words(300, 1)) +
					para(words(20, 2), `<w:sectPr>${LANDSCAPE}</w:sectPr>`) +
					Array.from({ length: 6 }, (_, k) => para(words(250, k + 3))).join("")
			) +
			para(words(100, 11)) +
			`<w:sectPr>${PAGE}</w:sectPr>`
		const edited = await editedLayout(buildDocx({ body }))
		const target = paragraphs(edited.session.state.doc, false)[5]

		if (target === undefined) throw new Error("no paragraph")

		// Text added in the second section, inside the control: layout meets the earlier pages only after it.
		await edited.edit(edited.session.state.tr.insertText(`${words(400, 4)} `, target.pos + 3))
		await expectSaved(edited)
	})

	it("puts a block added after a content control in the section after every one the control ends", async () => {
		const body =
			control(
				para(words(20, 1), `<w:sectPr>${LANDSCAPE}</w:sectPr>`) +
					Array.from({ length: 4 }, (_, k) => para(words(200, k + 2))).join("") +
					para(words(20, 6), `<w:sectPr>${NARROW}</w:sectPr>`)
			) +
			Array.from({ length: 3 }, (_, k) => para(words(60, k + 7))).join("") +
			`<w:sectPr>${PAGE}</w:sectPr>`
		const edited = await editedLayout(buildDocx({ body }))
		const at = edited.session.state.doc.child(0).nodeSize
		const added = docxSchema.nodes.paragraph.create(null, docxSchema.text(words(80, 9)))

		await edited.edit(edited.session.state.tr.insert(at, added))
		expect(edited.story.entries[1]?.section).toBe(2)
		await expectSaved(edited)
	})

	it("ends a body without section properties of its own as a document opened so ends it", async () => {
		// Ending with a section break: no section after it, whatever is typed in it.
		const ending =
			para(words(30, 1)) +
			para(words(20, 2), `<w:sectPr>${LANDSCAPE}</w:sectPr>`) +
			Array.from({ length: 4 }, (_, k) => para(words(120, k + 3))).join("")
		const edited = await editedLayout(buildDocx({ body: ending + para(words(20, 9), `<w:sectPr>${NARROW}</w:sectPr>`) }))
		const last = (): { node: Node; pos: number } => {
			const all = paragraphs(edited.session.state.doc, false)
			const found = all[all.length - 1]

			if (found === undefined) throw new Error("no paragraph")

			return found
		}

		await edited.edit(edited.session.state.tr.insertText("Typed ", last().pos + 2))
		expect(edited.edits.model.sections.length).toBe(2)
		await expectSaved(edited)

		// A paragraph after that break: a last section of default properties holds it.
		const after = last().pos + last().node.nodeSize

		await edited.edit(edited.session.state.tr.insert(after, docxSchema.nodes.paragraph.create(null, docxSchema.text(words(40, 3)))))
		expect(edited.edits.model.sections.length).toBe(3)
		await expectSaved(edited)

		// Removed again, and in a document opened with blocks after its last break.
		await edited.edit(session => session.execute({ type: "history.undo" }))
		expect(edited.edits.model.sections.length).toBe(2)
		await expectSaved(edited)

		const trailing = await editedLayout(
			buildDocx({ body: ending + para(words(20, 9), `<w:sectPr>${NARROW}</w:sectPr>`) + para(words(40, 3)) })
		)
		const end = trailing.session.state.doc.content.size
		const lastChild = trailing.session.state.doc.lastChild

		if (lastChild === null) throw new Error("no paragraph")

		await trailing.edit(trailing.session.state.tr.delete(end - lastChild.nodeSize, end))
		expect(trailing.edits.model.sections.length).toBe(2)
		await expectSaved(trailing)
	})

	it("keeps pages of mirrored margins that do not differ when an edit moves them to the other side", async () => {
		const body =
			Array.from({ length: 8 }, (_, chapter) =>
				[para(`Chapter ${String(chapter)}`, chapter === 0 ? "" : "<w:pageBreakBefore/>"), para(words(40, chapter))].join("")
			).join("") + `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>${PAGE}</w:sectPr>`

		for (const even of [false, true]) {
			// The even pages' header: the same as the others', or one of their own.
			const evenHeader = `<w:headerReference w:type="even" r:id="${even ? "rIdH2" : "rIdH1"}"/>`
			const edited = await editedLayout(
				buildDocx({
					body: body.replace("<w:sectPr>", `<w:sectPr>${evenHeader}`),
					headers: { rIdH1: para("Header"), rIdH2: para("Even header") },
					settings: "<w:evenAndOddHeaders/><w:mirrorMargins/>"
				})
			)
			const target = paragraphs(edited.session.state.doc, false)[1]

			if (target === undefined) throw new Error("no paragraph")

			// One page more before the later chapters: equal margins and the same header look alike on either side.
			await edited.edit(edited.session.state.tr.insertText(`${words(1500, 2)} `, target.pos + 2))
			expect(edited.layout.pages.length).toBe(9)

			if (even) expect(edited.last.resyncedAt).toBe(-1)
			else expect(edited.last.resyncedAt).toBeGreaterThan(0)

			await expectSaved(edited)
		}
	})

	it("counts placements by the page's parity for the layout that made them only", async () => {
		// Chapters filling both columns of a page: a page stays open from one step to the next.
		const columns = `<w:sectPr>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`
		const body =
			Array.from({ length: 8 }, (_, chapter) =>
				[para(`Chapter ${String(chapter)}`, chapter === 0 ? "" : "<w:pageBreakBefore/>"), para(words(700, chapter))].join("")
			).join("") + columns
		const edited = await editedLayout(buildDocx({ body }))
		// Another document placing a frame by the page's parity on every page.
		const other = await layOut(
			buildDocx({
				body:
					Array.from(
						{ length: 6 },
						(_, k) =>
							para(words(20, k), '<w:framePr w:w="2000" w:hAnchor="margin" w:xAlign="inside" w:y="200"/>') +
							para(words(200, k))
					).join("") + `<w:sectPr>${PAGE}</w:sectPr>`
			})
		)

		// Laid out again in steps, the other document laid out between them (pages stay open across steps).
		edited.layout.beginEdit(0, undefined)

		for (let round = 0; round < 10_000; round++) {
			const result = edited.layout.advance(-Infinity, 0)

			if (result === "done") break
			if (result === "fonts") await edited.layout.prepareFonts()

			other.beginEdit(0, undefined)
			await finishLayout(other)
		}

		const target = paragraphs(edited.session.state.doc, false)[1]

		if (target === undefined) throw new Error("no paragraph")

		// One page more: the later pages, on the other side now, are kept.
		await edited.edit(edited.session.state.tr.insertText(`${words(900, 2)} `, target.pos + 2))
		expect(edited.layout.pages.length).toBe(9)
		expect(edited.last.resyncedAt).toBeGreaterThan(0)
		await expectSaved(edited)
	})

	it("keeps one holder of a source as itself and more as a list", async () => {
		const edited = await editedLayout(longDocument(40))
		const holders = (edited.story as unknown as { holders: Map<number, unknown> }).holders
		const doc = edited.session.state.doc
		const [into, copied] = paragraphs(doc, false)

		if (into === undefined || copied === undefined) throw new Error("no paragraph")

		expect(holders.size).toBeGreaterThan(30)
		expect([...holders.values()].every(held => !Array.isArray(held) && !(held instanceof Set))).toBe(true)

		// A copy pasted: its source has two holders.
		await edited.edit(edited.session.state.tr.replace(into.pos + 3, into.pos + 3, doc.slice(copied.pos - 3, copied.pos + 4)))
		expect([...holders.values()].filter(held => Array.isArray(held) && held.length === 2).length).toBe(1)

		await edited.edit(session => session.execute({ type: "history.undo" }))
		expect([...holders.values()].every(held => !Array.isArray(held))).toBe(true)
		await expectSaved(edited)
	})

	it("writes no section break into a table cell but where the cell's paragraph had one", async () => {
		const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
		// The second cell's paragraph carries a section break, which the model ignores there.
		const table =
			'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
			`<w:tr>${cell(para(words(12, 1)))}${cell(para(words(12, 2), `<w:sectPr>${NARROW}</w:sectPr>`))}</w:tr></w:tbl>`
		const body =
			para(words(30, 3)) +
			para(words(20, 4), `<w:sectPr>${LANDSCAPE}</w:sectPr>`) +
			table +
			Array.from({ length: 6 }, (_, k) => para(words(80, k + 5))).join("") +
			`<w:sectPr>${PAGE}</w:sectPr>`
		// Per cell, whether the saved file has a section break in it.
		const cellBreaks = (bytes: Uint8Array): boolean[] => {
			const xml = new TextDecoder().decode(unzipSync(bytes)["word/document.xml"])

			return [...xml.matchAll(/<w:tc>.*?<\/w:tc>/g)].map(match => match[0].includes("<w:sectPr"))
		}
		const edited = await editedLayout(buildDocx({ body }))
		const doc = edited.session.state.doc
		const breakAt = paragraphs(doc, false)[1]
		const first = paragraphs(doc, true)[0]

		if (breakAt === undefined || first === undefined) throw new Error("no paragraph")

		expect(cellBreaks(await edited.save())).toEqual([false, true])

		// Text up to the start of the paragraph ending the first section, pasted into the first cell: the cell's
		// paragraph takes that paragraph's mark, section break included, which is none there.
		await edited.edit(edited.session.state.tr.replace(first.pos + 3, first.pos + 3, doc.slice(breakAt.pos - 3, breakAt.pos + 4)))
		expect(paragraphs(edited.session.state.doc, true, 1).some(at => at.node.attrs["sectPr"] === breakAt.node.attrs["sectPr"])).toBe(
			true
		)

		const saved = await edited.save()

		expect(cellBreaks(saved)).toEqual([false, true])
		expect(edited.edits.model.sections.length).toBe(2)
		expect(pageDigest(edited.layout)).toEqual(await savedDigest(saved))
	})

	it("writes a section break copied unchanged before its source in or out of table cells as it was", async () => {
		const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
		const table = (first = ""): string =>
			'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
			`<w:tr>${cell(para("cell one text") + first)}${cell(para("cell two text"))}</w:tr></w:tbl>`
		// Per cell, whether the saved file has a section break in it; and every section break in it.
		const breaks = (bytes: Uint8Array): { cells: boolean[]; all: number } => {
			const xml = new TextDecoder().decode(unzipSync(bytes)["word/document.xml"])

			return {
				cells: [...xml.matchAll(/<w:tc>.*?<\/w:tc>/g)].map(match => match[0].includes("<w:sectPr")),
				all: (xml.match(/<w:sectPr/g) ?? []).length
			}
		}
		const check = async (edited: EditedLayout, cells: boolean[], all: number): Promise<void> => {
			const saved = await edited.save()

			expect(breaks(saved)).toEqual({ cells, all })
			expect(pageDigest(edited.layout)).toEqual(await savedDigest(saved))
		}
		const sectionBreak = para("Break paragraph", `<w:sectPr>${LANDSCAPE}</w:sectPr>`)
		const cellEnd = (doc: Node): number => {
			const first = paragraphs(doc, true, 1)[0]

			if (first === undefined) throw new Error("no paragraph")

			return first.pos + first.node.nodeSize - 1
		}
		const breakParagraph = (doc: Node): { node: Node; pos: number } => {
			const found = paragraphs(doc, false, 1).find(at => (at.node.attrs["sectPr"] as number) >= 0)

			if (found === undefined) throw new Error("no section break")

			return found
		}

		// A section break paragraph copied, unchanged, to the end of a cell of an earlier table: it has the source
		// now and is still no section break there.
		const into = await editedLayout(
			buildDocx({
				body: table() + para(words(30, 1)) + sectionBreak + para(words(40, 2)) + `<w:sectPr>${PAGE}</w:sectPr>`
			})
		)
		const doc = into.session.state.doc
		const source = breakParagraph(doc)

		await into.edit(
			into.session.state.tr.replace(cellEnd(doc), cellEnd(doc), doc.slice(source.pos - 2, source.pos + source.node.nodeSize - 1))
		)
		await check(into, [false, false], 2)

		// A content control holding one, moved whole into a cell (the editor refuses that move; the writer still
		// must not copy the break into the cell): written without it there.
		const controlled = await editedLayout(
			buildDocx({ body: table() + para(words(30, 1)) + control(sectionBreak + para(words(20, 3))) + `<w:sectPr>${PAGE}</w:sectPr>` })
		)
		const doc2 = controlled.session.state.doc
		let at = -1

		doc2.forEach((node, offset) => {
			if (node.type === docxSchema.nodes.sdt) at = offset
		})

		const holder = doc2.nodeAt(at)

		if (holder === null) throw new Error("no content control")

		const writer = new StoryWriter(
			{
				context: controlled.context,
				nextRevisionId: () => 1,
				numId: id => id,
				hyperlink: () => {
					throw new Error("no hyperlink")
				}
			},
			controlled.session.story
		)
		const moved = new Transform(doc2).delete(at, at + holder.nodeSize).insert(cellEnd(doc2) + 1, holder).doc
		const written = new TextDecoder().decode(writer.writeStory(moved).bytes)

		expect([...written.matchAll(/<w:tc>.*?<\/w:tc>/g)].map(match => match[0].includes("<w:sectPr"))).toEqual([false, false])
		expect(written).toContain("Break paragraph")

		// A cell's paragraph holding one (which the model ignores there) copied to the top level before the table:
		// no section there either, in the save or in what it laid out from.
		const out = await editedLayout(
			buildDocx({
				body:
					para("P zero text") +
					table(para("cell Y with a break", `<w:sectPr>${LANDSCAPE}</w:sectPr>`)) +
					para(words(30, 1)) +
					`<w:sectPr>${PAGE}</w:sectPr>`
			})
		)
		const doc3 = out.session.state.doc
		const [zero] = paragraphs(doc3, false, 1)
		const [one, y] = paragraphs(doc3, true, 1)

		if (zero === undefined || one === undefined || y === undefined) throw new Error("no paragraph")

		const end = zero.pos + zero.node.nodeSize - 1

		await out.edit(out.session.state.tr.replace(end, end, doc3.slice(one.pos + one.node.nodeSize - 3, y.pos + y.node.nodeSize - 1)))
		expect(out.session.state.doc.child(1).attrs["sectPr"]).toBe(y.node.attrs["sectPr"])
		await check(out, [true, false], 2)

		for (const record of out.edits.overlayRecords()) {
			const document = record.part.document
			const node = record.part.node(record.item)

			expect(new TextDecoder().decode(document.bytes.subarray(document.start(node), document.end(node)))).not.toContain("sectPr")
		}
	})

	it("keeps pages across an odd page shift only when nothing on them depends on the page's parity", async () => {
		const header = (text: string, lines: number): string =>
			Array.from({ length: lines }, (_, k) => para(`${text} ${String(k)}`)).join("")
		const references =
			'<w:headerReference w:type="default" r:id="rIdH1"/><w:headerReference w:type="even" r:id="rIdH2"/><w:headerReference w:type="first" r:id="rIdH3"/>'

		for (const variant of ["evenAndOdd", "titlePg", "oddPage", "evenPage"]) {
			const sections = variant === "oddPage" || variant === "evenPage"
			const type = sections ? `<w:type w:val="${variant}"/>` : ""
			const sectPr = `<w:sectPr>${references}${type}${PAGE}${variant === "titlePg" ? "<w:titlePg/>" : ""}</w:sectPr>`
			// Chapters starting pages of their own, pages after an edit starting where pages before it did; for odd
			// and even page sections, a section every three chapters.
			const body =
				Array.from(
					{ length: 8 },
					(_, chapter) =>
						para(`Chapter ${String(chapter)}`, chapter === 0 || (sections && chapter % 3 === 0) ? "" : "<w:pageBreakBefore/>") +
						para(words(120, chapter), sections && chapter % 3 === 2 ? sectPr : "")
				).join("") + sectPr
			const edited = await editedLayout(
				buildDocx({
					body,
					headers: { rIdH1: header("Odd header", 1), rIdH2: header("Even header", 6), rIdH3: header("First header", 3) },
					settings: sections ? "" : "<w:evenAndOddHeaders/>"
				})
			)
			const target = paragraphs(edited.session.state.doc, false)[1]

			if (target === undefined) throw new Error("no paragraph")

			// A page more in the first chapter: later pages are on the other side.
			await edited.edit(edited.session.state.tr.insertText(`${words(1500, 1)} `, target.pos + 2))
			await expectSaved(edited)
		}
	})

	it("lays a page out again when a row that could not wait for a font drew stand-ins", async () => {
		const run = (text: string, font: string): string =>
			`<w:p><w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`
		const table = (widths: readonly number[], rows: string): string =>
			`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>${rows}</w:tbl>`
		const inner = (count: number, font: string, name: string): string =>
			Array.from({ length: count }, (_, k) => `<w:tr><w:tc>${run(`${name} ${k}`, font)}</w:tc></w:tr>`).join("")
		// A merge running out of its page inside a nested table cannot be laid out again later: its row continues
		// on the next page with stand-ins for a font first needed there.
		const merged = `<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${table([3000], inner(40, "Arial", "arial") + inner(20, "Courier New", "mono"))}</w:tc>`
		const rows =
			`<w:tr>${merged}<w:tc>${run("b", "Arial")}</w:tc></w:tr>` +
			`<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc>${run("c", "Arial")}</w:tc></w:tr>`
		const body = Array.from({ length: 40 }, (_, k) => run(`line ${k}`, "Arial")).join("") + table([4000, 2000], rows)
		const pages = await savedDigest(buildDocx({ body }))

		expect(pages.join("\n")).toMatch(/"mono 0" \[[^\]]* fotest-cousine:/)
		expect(pages.join("\n")).not.toContain("null:")
	})

	it("lays out blocks put back after the body's last block, before its final section properties", async () => {
		const cell = (text: string): string => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${para(text)}</w:tc>`
		const rows = Array.from({ length: 60 }, (_, k) => `<w:tr>${cell(`a${k}`)}${cell(`b${k}`)}</w:tr>`).join("")
		const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`
		const sectPr =
			'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		const edited = await editedLayout(buildDocx({ body: para("first") + table + sectPr }))
		const session = edited.session
		const last = (): { pos: number; node: Node } => {
			const node = session.state.doc.lastChild as Node

			return { pos: session.state.doc.content.size - node.nodeSize, node }
		}
		const { pos, node } = last()

		await edited.edit(session.state.tr.delete(pos, pos + node.nodeSize))
		await expectSaved(edited)
		// The undo puts the table back after the last block: the last section grows with it.
		await edited.edit(each => each.execute({ type: "history.undo" }))
		await expectSaved(edited)
		await edited.edit(session.state.tr.insert(session.state.doc.content.size, docxSchema.nodes.paragraph.create()))
		await expectSaved(edited)
	})

	it("saves a locked table holding a section break in a cell, and refuses a locked section break moved into one", async () => {
		const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
		const row = `<w:tr>${cell(para("cell one text", `<w:sectPr>${LANDSCAPE}</w:sectPr>`))}${cell(para("cell two text"))}</w:tr>`
		// A row inside w:customXml locks the table.
		const table =
			'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
			`<w:customXml w:element="r">${row}</w:customXml>${row}</w:tbl>`
		const edited = await editedLayout(
			buildDocx({ body: para("P zero text") + table + para(words(30, 1)) + `<w:sectPr>${PAGE}</w:sectPr>` })
		)
		const locked = edited.session.state.doc.child(1)

		expect(locked.type).toBe(docxSchema.nodes.locked_block)
		expect(locked.attrs["sectPr"]).toBeGreaterThanOrEqual(0)

		// An edit elsewhere: the table is written as it was, its break where it was.
		await edited.edit(edited.session.state.tr.insertText("XX ", 3))
		await expectSaved(edited)

		// A locked paragraph ending a section (w:customXml around it), moved into a table cell: its bytes would put
		// the break in the cell, so it is not written (the editor refuses that move; the writer refuses too).
		const wrapped = await editedLayout(
			buildDocx({
				body:
					table.replace(
						`<w:customXml w:element="r">${row}</w:customXml>`,
						`<w:tr>${cell(para("cell one text"))}${cell(para("cell two text"))}</w:tr>`
					) +
					`<w:customXml w:element="x">${para("Wrapped break", `<w:sectPr>${LANDSCAPE}</w:sectPr>`)}</w:customXml>` +
					para(words(30, 2)) +
					`<w:sectPr>${PAGE}</w:sectPr>`
			})
		)
		const doc = wrapped.session.state.doc
		const block = doc.child(1)
		const into = doc.child(0)
		let cellAt = -1

		doc.descendants((node, pos) => {
			if (cellAt < 0 && node.type === docxSchema.nodes.paragraph) cellAt = pos + node.nodeSize

			return cellAt < 0
		})

		expect(block.type).toBe(docxSchema.nodes.locked_block)
		expect(into.type).not.toBe(docxSchema.nodes.locked_block)

		const writer = new StoryWriter(
			{
				context: wrapped.context,
				nextRevisionId: () => 1,
				numId: id => id,
				hyperlink: () => {
					throw new Error("no hyperlink")
				}
			},
			wrapped.session.story
		)
		const moved = new Transform(doc).delete(into.nodeSize, into.nodeSize + block.nodeSize).insert(cellAt, block).doc

		expect(() => writer.writeStory(moved)).toThrow("A locked section break moved into or out of a table cell")
		expect(() => writer.writeStory(doc)).not.toThrow()
	})
})

describe("overlay parts", () => {
	const prefixesOf = (xml: string): string[] | undefined => {
		const out = new Set<string>()

		return namePrefixes(new TextEncoder().encode(xml), out) ? [...out].sort() : undefined
	}

	it("declare only the prefixes their names use, and fall back where names alone do not tell", () => {
		// `xml` is bound without a declaration; attributes without a prefix are in no namespace.
		expect(prefixesOf('<w:p w14:paraId="1"><w:r><w:t xml:space="preserve">a: b &lt; c</w:t></w:r></w:p>')).toEqual(["w", "w14"])
		expect(prefixesOf('<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="2"/></wp:inline></w:drawing></w:r>')).toEqual(["w", "wp"])
		expect(prefixesOf('<w:p><mc:AlternateContent><mc:Choice Requires="wps"/></mc:AlternateContent></w:p>')).toBeUndefined()
		expect(prefixesOf('<w:p xmlns:x="urn:x"><x:y/></w:p>')).toBeUndefined()
		expect(prefixesOf("<w:p><!-- note --></w:p>")).toBeUndefined()
		expect(prefixesOf("<w:p a=\"1\" b='>'/>")).toEqual(["w"])
		expect(prefixesOf("<p/>")).toEqual([""])
	})

	it("parse an edited block in the minimal root as in the source's own", async () => {
		const edited = await editedLayout(buildDocx({ body: para("one") }))
		const source = edited.story.part
		const container = edited.story.story.story.item
		const block =
			'<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r>' +
			'<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="2"/></wp:inline></w:drawing></w:r></w:p>'
		const encode = (xml: string): Uint8Array => new TextEncoder().encode(xml)
		const minimal = edited.edits.addPart(source, container, encode(block), true)
		// A comment: the names alone do not tell, so the source's own root tag is used.
		const full = edited.edits.addPart(source, container, encode(`${block}<!-- full -->`), true)

		if (minimal === undefined || full === undefined) throw new Error("not parsed")

		expect(minimal.offset).toBeLessThan(full.offset)
		expect(minimal.items.length).toBe(1)

		const items = (part: typeof minimal.part, first: number): string[] => {
			const out: string[] = []

			for (let item = first; item < part.end(first); item++)
				out.push(`${String(part.kinds[item])}:${part.kinds[item] === Item.Text ? part.text(item) : ""}`)

			return out
		}

		expect(items(minimal.part, minimal.items[0] ?? -1)).toEqual(items(full.part, full.items[0] ?? -1))
	})
})
