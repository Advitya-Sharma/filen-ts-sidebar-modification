import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { OpcPackage, bytesSource } from "@filen/office/core/opc"
import type { DocxPackage } from "@filen/office/core/docx"
import { docxSchema, type DocxEditContext, type DocxEditorCommand } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { DocxSaveError, saveDocx } from "@filen/office/core/docx/write"
import { markInternal } from "../core/docx/edit/invariants/index"
import { bodySections, docSectionChanges } from "../core/docx/stories/sections"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"

const decoder = new TextDecoder()
const marks = docxSchema.marks

const MARGINS = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const SECTION = (extra = ""): string =>
	`<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}${extra}</w:sectPr>`
const LONG = Array.from({ length: 30 }, (_value, k) => p(`Paragraph ${String(k)} with some words to fill the line`)).join("")

async function read(bytes: Uint8Array, name: string): Promise<string> {
	return decoder.decode(await (await OpcPackage.open(bytesSource(bytes))).read(name))
}

interface Opened {
	readonly editor: PagedEditor
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
}

function caret(editor: PagedEditor, pos: number): void {
	const session = editor.session

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, pos)))
}

async function run(editor: PagedEditor, command: DocxEditorCommand): Promise<boolean> {
	return await editor.execute(command)
}

async function saved({ editor, pkg, context }: Opened): Promise<Uint8Array> {
	await settleAll(editor)

	const result = await saveDocx({ pkg, context, sessions: editor.sessions })

	if (result.bytes === undefined) throw new Error("No bytes")

	expect(pageDigest(editor.layout)).toEqual(await savedDigest(result.bytes))

	return result.bytes
}

// The w:sectPr elements of the saved main part (not those a change records inside one), in order.
async function sections(bytes: Uint8Array): Promise<string[]> {
	const xml = await read(bytes, "word/document.xml")
	const out: string[] = []
	let depth = 0
	let start = 0

	for (const match of xml.matchAll(/<w:sectPr>|<\/w:sectPr>/g)) {
		if (match[0] === "<w:sectPr>") {
			if (depth === 0) start = match.index

			depth++
		} else if (--depth === 0) {
			out.push(xml.slice(start, match.index + match[0].length))
		}
	}

	return out
}

describe("docx sections and page setup", () => {
	it.each(["nextPage", "continuous", "evenPage", "oddPage"] as const)(
		"inserts a %s break: the new section takes the properties and references, the one after links to it",
		async kind => {
			const opened = await pagedEditor(
				buildDocx({ body: p("Alpha beta") + p("Second") + SECTION('<w:pgNumType w:start="4"/>'), headers: { rIdH: p("Head") } })
			)
			const { editor } = opened

			caret(editor, 6)
			expect(editor.snapshot().enabled["insert.sectionBreak"]).toBe(true)
			expect(await run(editor, { type: "insert.sectionBreak", kind })).toBe(true)
			// The caret starts the new section.
			expect(editor.session.state.selection.from).toBe(8)

			const bytes = await saved(opened)
			const [first, last] = await sections(bytes)

			expect(first).toBe(
				`<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}<w:pgNumType w:start="4"/></w:sectPr>`
			)
			expect(last).toBe(
				`<w:sectPr>${kind === "nextPage" ? "" : `<w:type w:val="${kind}"/>`}<w:pgSz w:w="12240" w:h="15840"/>${MARGINS}</w:sectPr>`
			)
			expect((await openBytes(bytes)).model.sections).toHaveLength(2)

			// One undo takes the break back.
			expect(await run(editor, { type: "history.undo" })).toBe(true)
			expect(await sections(await saved(opened))).toHaveLength(1)
			editor.dispose()
		}
	)

	it("never breaks a section inside a table", async () => {
		const opened = await pagedEditor(
			buildDocx({
				body: `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${p("cell")}</w:tc></w:tr></w:tbl>${p("after")}`
			})
		)
		const { editor } = opened

		caret(editor, 4)
		expect(editor.snapshot().enabled["insert.sectionBreak"]).toBe(false)
		expect(await run(editor, { type: "insert.sectionBreak", kind: "nextPage" })).toBe(false)
		editor.dispose()
	})

	it("deletes a break: the text before joins the section after, which keeps its properties and what it showed", async () => {
		const opened = await pagedEditor(
			buildDocx({
				body:
					p("One") +
					`<w:p><w:pPr>${SECTION('<w:cols w:num="2" w:space="720"/>')}</w:pPr></w:p>` +
					p("Two") +
					`<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>${MARGINS}</w:sectPr>`,
				headers: { rIdH: p("Head") }
			})
		)
		const { editor } = opened

		caret(editor, 2)
		expect(await run(editor, { type: "section.deleteBreak" })).toBe(true)

		const bytes = await saved(opened)
		const list = await sections(bytes)

		expect(list).toEqual([
			`<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>${MARGINS}</w:sectPr>`
		])
		// The last section has no break to delete.
		expect(editor.snapshot().enabled["section.deleteBreak"]).toBe(false)
		editor.dispose()
	})

	it("sets up pages: size, orientation, margins, columns, page numbers and line numbers, for a section or all", async () => {
		const opened = await pagedEditor(
			buildDocx({ body: LONG + `<w:p><w:pPr>${SECTION()}</w:pPr></w:p>` + LONG + SECTION(), headers: { rIdH: p("Head") } })
		)
		const { editor } = opened

		caret(editor, 2)
		expect(
			await run(editor, {
				type: "page.setup",
				size: { width: 11906, height: 16838 },
				orientation: "landscape",
				margins: { left: 720, gutter: 360 }
			})
		).toBe(true)
		expect(
			await run(editor, {
				type: "page.setup",
				columns: { count: 2, space: 360, separator: true },
				pageNumbers: { start: 7, format: "upperRoman" }
			})
		).toBe(true)
		expect(await run(editor, { type: "page.setup", lineNumbers: { countBy: 5, start: 3, restart: "newSection" } })).toBe(true)

		let list = await sections(await saved(opened))

		expect(list[0]).toBe(
			'<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
				'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="720" w:header="720" w:footer="720" w:gutter="360"/>' +
				'<w:lnNumType w:countBy="5" w:start="2" w:restart="newSection"/><w:pgNumType w:fmt="upperRoman" w:start="7"/>' +
				'<w:cols w:num="2" w:space="360" w:sep="1"/></w:sectPr>'
		)
		expect(list[1]).toBe(SECTION())

		// The whole document: every section.
		expect(await run(editor, { type: "page.setup", margins: { top: 720 }, scope: "document" })).toBe(true)
		list = await sections(await saved(opened))
		expect(list.every(section => section.includes('w:top="720"'))).toBe(true)

		// Numbering continues again; no line numbers.
		expect(await run(editor, { type: "page.setup", pageNumbers: { start: null }, lineNumbers: null })).toBe(true)
		list = await sections(await saved(opened))
		expect(list[0]).toContain('<w:pgNumType w:fmt="upperRoman"/>')
		expect(list[0]).not.toContain("lnNumType")
		editor.dispose()
	})

	it("sets up the pages of a body without section properties of its own", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Plain") }))
		const { editor } = opened

		expect(await run(editor, { type: "page.setup", orientation: "landscape" })).toBe(true)
		expect(await sections(await saved(opened))).toEqual(['<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr>'])
		editor.dispose()
	})

	it("tracks page setup while changes are tracked, and accepts or rejects it", async () => {
		for (const accept of [true, false]) {
			const opened = await pagedEditor(buildDocx({ body: p("A1") + p("A2") + SECTION(), headers: { rIdH: p("Head") } }))
			const { editor, context } = opened

			context.identity.author = "Tester"
			context.review.tracking = true
			caret(editor, 2)
			expect(await run(editor, { type: "page.setup", orientation: "landscape" })).toBe(true)
			expect(editor.snapshot().review.revisionCount).toBe(1)

			const [recorded] = await sections(await saved(opened))

			expect(recorded).toMatch(
				/<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>.*<w:sectPrChange w:id="\d+" w:author="Tester" w:date="[^"]+"><w:sectPr><w:pgSz w:w="12240" w:h="15840"\/>.*<\/w:sectPr><\/w:sectPrChange>/
			)
			expect(recorded).not.toMatch(/<w:sectPrChange.*headerReference/)

			expect(await run(editor, { type: accept ? "review.acceptAll" : "review.rejectAll" })).toBe(true)
			expect(editor.snapshot().review.revisionCount).toBe(0)

			const [resolved] = await sections(await saved(opened))

			expect(resolved).not.toContain("sectPrChange")
			expect(resolved?.includes("landscape")).toBe(accept)
			editor.dispose()
		}
	})

	it("accepts or rejects a section change the file recorded", async () => {
		const change = `<w:sectPrChange w:id="9" w:author="Word" w:date="2026-01-01T00:00:00Z"><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}</w:sectPr></w:sectPrChange>`
		const bytes = buildDocx({
			body: p("Text") + `<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>${MARGINS}${change}</w:sectPr>`
		})

		for (const accept of [true, false]) {
			const opened = await pagedEditor(bytes)
			const { editor } = opened

			expect(editor.snapshot().review.revisionCount).toBe(1)
			expect(await run(editor, { type: accept ? "review.acceptAll" : "review.rejectAll" })).toBe(true)
			expect(editor.snapshot().review.revisionCount).toBe(0)
			expect(await sections(await saved(opened))).toEqual([
				accept
					? `<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>${MARGINS}</w:sectPr>`
					: `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}</w:sectPr>`
			])
			editor.dispose()
		}
	})

	it("refuses section properties that leave no room for text or name a part that is no header", async () => {
		const opened = await pagedEditor(buildDocx({ body: p("Text") + SECTION(), headers: { rIdH: p("Head") } }))
		const { editor, pkg, context } = opened
		const body = editor.body.session
		const edits = [
			{ pageMargins: { top: 1440, right: 6000, bottom: 1440, left: 6240, header: 720, footer: 720, gutter: 0 } },
			{ columns: { count: 40, space: 1440, separator: false } },
			{ headers: { default: body.story.story.index } }
		]

		for (const edit of edits) {
			const key = body.state.doc.attrs["sectPr"] as number

			body.dispatch(
				markInternal(
					body.state.tr.setDocAttribute("sectionChanges", {
						...docSectionChanges(body.state.doc),
						final: { section: key, base: key, ...edit }
					})
				)
			)

			const error = await saveDocx({ pkg, context, sessions: editor.sessions }).catch((caught: unknown) => caught)

			expect(error, JSON.stringify(edit)).toBeInstanceOf(DocxSaveError)
		}

		// A new break whose edit is lost cannot be written: refused, not dropped.
		const doc = body.state.doc

		body.dispatch(markInternal(body.state.tr.setNodeAttribute(0, "sectPr", -7).setDocAttribute("sectionChanges", null)))
		expect(bodySections(body.state.doc, body.story.part)[0]?.key).toBe(-7)
		expect(
			await saveDocx({ pkg, context, sessions: editor.sessions }).catch((caught: unknown) => (caught as DocxSaveError).reason)
		).toBe("write")
		expect(marks.section_change.isInSet(doc.child(0).marks)).toBeUndefined()
		editor.dispose()
	})
})
