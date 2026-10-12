import { readFileSync } from "node:fs"
import { NodeSelection, TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { DocxEditorSession, docxSchema } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { DocxVerifyError, saveDocx, type CheckInput } from "@filen/office/core/docx/write"
import { parseXml } from "@filen/office/core/xml"
import { textboxChecks } from "../core/docx/write/checks/textbox"
import { textboxCopies } from "../core/docx/write/nested"
import { openBytes } from "./docxEditFixtures"
import { buildDocx, p } from "./docxSynthetic"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"

const decoder = new TextDecoder()
const nodes = docxSchema.nodes
const FIXTURES = new URL("./fixtures/", import.meta.url).pathname

// ---- documents -------------------------------------------------------------------------------------------

function content(text: string): string {
	return `<w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent>`
}

// A text box as Word 16 writes one: a wps shape in mc:Choice, the same content in the VML of mc:Fallback.
function textbox(text: string, options: { fallbackText?: string; linkedId?: string; fallback?: boolean } = {}): string {
	const choice =
		'<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
		'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
		'<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
		'<wp:extent cx="2743200" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>' +
		'<wp:docPr id="3" name="Text Box 3"/><wp:cNvGraphicFramePr/>' +
		'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>' +
		'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
		`<wps:txbx${options.linkedId === undefined ? "" : ` id="${options.linkedId}"`}>${content(text)}</wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`

	if (options.fallback === false) return `<w:r>${choice}</w:r>`

	const fallback =
		'<w:pict><v:shape id="Text Box 3" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:0;margin-top:0;width:216pt;height:72pt;z-index:251659264">' +
		`<v:textbox>${content(options.fallbackText ?? text)}</v:textbox><w10:wrap type="square"/></v:shape></w:pict>`

	return `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></w:r>`
}

const NS =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
	'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
	'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"'

function docx(body: string, header?: string): Uint8Array {
	const sectPr = `<w:sectPr>${header === undefined ? "" : '<w:headerReference w:type="default" r:id="rIdH"/>'}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`

	return buildDocx({
		document: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${sectPr}</w:body></w:document>`,
		...(header === undefined
			? {}
			: {
					relationships: [{ id: "rIdH", type: "header", target: "header1.xml" }],
					parts: [
						{
							name: "word/header1.xml",
							contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
							data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}>${header}</w:hdr>`
						}
					]
				})
	})
}

const BOXED = docx(`<w:p>${textbox("Inside the box")}<w:r><w:t>Body text holding the box.</w:t></w:r></w:p>${p("After")}`)

// ---- helpers ---------------------------------------------------------------------------------------------

function textboxStory(editor: PagedEditor, index = 0): number {
	const story = editor.edits.model.stories.filter(entry => entry?.kind === "textbox")[index]

	if (story === undefined) throw new Error("no text box")

	return story.index
}

function enterBox(editor: PagedEditor, index = 0): void {
	expect(editor.enterStory(textboxStory(editor, index), 0)).toBe(true)
	expect(editor.session.readOnly).toBe(false)
}

async function save(editor: PagedEditor, pkg: Parameters<typeof saveDocx>[0]["pkg"]): Promise<Uint8Array> {
	await settleAll(editor)

	const result = await saveDocx({ pkg, context: editor.body.session.context, sessions: editor.sessions })

	return result.bytes as Uint8Array
}

async function documentXml(bytes: Uint8Array, name = "word/document.xml"): Promise<string> {
	const opened = await openBytes(bytes)

	return decoder.decode(await opened.pkg.opc.read(name))
}

function boxes(xml: string): string[] {
	return [...xml.matchAll(/<w:txbxContent>.*?<\/w:txbxContent>/g)].map(match => match[0])
}

async function sameLayout(editor: PagedEditor, bytes: Uint8Array): Promise<void> {
	await settleAll(editor)
	expect(pageDigest(editor.layout)).toEqual(await savedDigest(bytes))
}

// ---- tests -----------------------------------------------------------------------------------------------

describe("docx text box editing", () => {
	it("opens text boxes editable, except linked ones and those whose copies do not match up", async () => {
		const opened = await openBytes(
			docx(
				`<w:p>${textbox("one")}${textbox("two", { fallback: false })}${textbox("three", { linkedId: "1" })}</w:p>` +
					'<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="9" name="Next"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:linkedTxbx id="1" seq="1"/><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
			)
		)
		const reasons = opened.model.stories
			.filter(story => story.kind === "textbox")
			.map(story => new DocxEditorSession(opened.context, { story }).snapshot().readOnlyReasons)

		expect(reasons).toEqual([[], [], ["linkedTextbox"]])

		const part = opened.model.parts[0]

		if (part === undefined) throw new Error("no part")

		// A Fallback with more text boxes than its Choice cannot be matched up.
		const mismatched = await openBytes(
			docx(
				`<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="2" name="B"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx>${content("x")}</wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect><v:textbox>${content("x")}</v:textbox></v:rect><v:rect><v:textbox>${content("y")}</v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`
			)
		)
		const box = mismatched.model.stories.find(story => story.kind === "textbox")

		if (box === undefined) throw new Error("no text box")

		expect(new DocxEditorSession(mismatched.context, { story: box }).snapshot().readOnlyReasons).toEqual(["textboxFallback"])
		expect(textboxCopies(part, part.node(opened.model.stories.find(story => story.kind === "textbox")?.item ?? -1))).toHaveLength(1)
	})

	it("writes an edited text box into its mc:Choice and its VML fallback alike", async () => {
		const { editor, pkg } = await pagedEditor(BOXED)

		enterBox(editor)
		editor.dispatch(editor.session.state.tr.insertText("Edited: ", 1))

		const bytes = await save(editor, pkg)
		const xml = await documentXml(bytes)
		const written = boxes(xml)

		expect(written).toHaveLength(2)
		expect(written[0]).toContain("Edited: Inside the box")
		expect(written[1]).toBe(written[0])
		expect(xml).toContain("<mc:Fallback>")
		expect(xml).toContain("Body text holding the box.")

		const reopened = await openBytes(bytes)
		const story = reopened.model.stories.find(entry => entry.kind === "textbox")

		if (story === undefined) throw new Error("no text box")

		expect(new DocxEditorSession(reopened.context, { story }).state.doc.textContent).toBe("Edited: Inside the box")
		await sameLayout(editor, bytes)

		// Undone, the document is its source again.
		expect(await editor.execute({ type: "history.undo" })).toBe(true)
		expect(await documentXml(await save(editor, pkg))).toBe(await documentXml(BOXED))
	})

	// The fallback copy's revisions are renumbered from the save's own count, which the other revisions it writes
	// take theirs from too: never one id twice.
	it("numbers tracked changes in a text box's fallback apart from the save's others", async () => {
		const { editor, pkg } = await pagedEditor(BOXED)

		editor.body.session.context.identity.author = "Author"
		expect(await editor.execute({ type: "review.track", on: true })).toBe(true)
		// One deletion across two paragraphs: one change, two w:del.
		editor.dispatch(editor.state.tr.delete(4, editor.state.doc.content.size - 3))
		enterBox(editor)
		editor.dispatch(editor.session.state.tr.insertText("Boxed ", 1))
		editor.exitStory()

		const xml = await documentXml(await save(editor, pkg))
		const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map(match => match[1])

		expect(ids.length).toBeGreaterThanOrEqual(3)
		expect(new Set(ids).size).toBe(ids.length)
	})

	it("writes a text box edited together with the paragraph holding it, and with the shape resized", async () => {
		const { editor, pkg } = await pagedEditor(BOXED)

		enterBox(editor)
		editor.dispatch(editor.session.state.tr.insertText("Boxed ", 1))
		await settleAll(editor)
		editor.exitStory()

		const doc = editor.session.state.doc
		let drawing = -1

		doc.descendants((node, pos) => {
			if (node.type === nodes.drawing) drawing = pos

			return drawing < 0
		})

		editor.dispatch(editor.session.state.tr.insertText("More ", drawing + 1))
		editor.dispatch(editor.session.state.tr.setSelection(NodeSelection.create(editor.session.state.doc, drawing)))
		expect(await editor.execute({ type: "drawing.resize", cx: 3657600, cy: 1828800 })).toBe(true)
		await settleAll(editor)

		const bytes = await save(editor, pkg)
		const xml = await documentXml(bytes)
		const written = boxes(xml)

		expect(written).toHaveLength(2)
		expect(written[0]).toContain("Boxed Inside the box")
		expect(written[1]).toBe(written[0])
		expect(xml).toContain('<w:t xml:space="preserve">More </w:t></w:r><w:r><w:t>Body text')
		expect(xml).toContain("width:288pt;height:144pt")
		await sameLayout(editor, bytes)

		// Typing in the box again after the paragraph was rewritten.
		enterBox(editor)
		editor.dispatch(editor.session.state.tr.insertText("!", editor.session.state.doc.content.size - 1))

		const again = await save(editor, pkg)

		expect(boxes(await documentXml(again))[1]).toContain("Boxed Inside the box!")
		await sameLayout(editor, again)
	})

	it("edits a text box in a header written with its header", async () => {
		const bytes = docx(p("Body"), `<w:p>${textbox("Header box")}<w:r><w:t>Header</w:t></w:r></w:p>`)
		const { editor, pkg } = await pagedEditor(bytes)

		enterBox(editor)
		editor.dispatch(editor.session.state.tr.insertText("New ", 1))

		const saved = await save(editor, pkg)
		const xml = await documentXml(saved, "word/header1.xml")

		expect(boxes(xml)).toEqual([boxes(xml)[0], boxes(xml)[0]])
		expect(xml).toContain("New Header box")
		await sameLayout(editor, saved)
	})

	it("edits the text boxes of a Word 16 fixture with their fallbacks", async () => {
		const { editor, pkg } = await pagedEditor(new Uint8Array(readFileSync(`${FIXTURES}06-textbox-shape.docx`)))
		const story = textboxStory(editor)

		expect(editor.enterStory(story, -1)).toBe(true)

		expect(editor.session.readOnly).toBe(false)

		editor.dispatch(editor.session.state.tr.setSelection(TextSelection.atEnd(editor.session.state.doc)))
		editor.dispatch(editor.session.state.tr.insertText(" (edited)"))

		const saved = await save(editor, pkg)

		await sameLayout(editor, saved)
	})
})

describe("docx text box checks", () => {
	it("refuses a rewritten mc:AlternateContent that lost its fallback", async () => {
		const source = await openBytes(BOXED)
		const part = source.model.parts[0]

		if (part === undefined) throw new Error("no part")

		const without = new TextEncoder().encode(
			`<w:document ${NS}><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>x</w:t></mc:Choice></mc:AlternateContent></w:r></w:p></w:body></w:document>`
		)
		const check = textboxChecks[0]
		const input = { original: source.pkg, context: source.context } as unknown as CheckInput

		await expect(
			(async () => {
				await check?.part?.(input, { name: part.name, part, document: parseXml(without), skeleton: false })
			})()
		).rejects.toBeInstanceOf(DocxVerifyError)
	})
})
