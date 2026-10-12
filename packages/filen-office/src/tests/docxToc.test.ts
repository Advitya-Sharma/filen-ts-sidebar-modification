import type { Node as ProseMirrorNode } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { docxSchema, type DocxEditContext } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import type { DocxPackage } from "@filen/office/core/docx"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { sessionBookmarks } from "../core/docx/fields/bookmarks"
import { parseFieldCode } from "../core/docx/fields/evaluate"
import { tocOptions, tocRegions } from "../core/docx/fields/toc"
import { markable } from "../core/docx/review/track"
import { STYLES_PART } from "./docxEditFixtures"
import { buildDocx, p, type SyntheticDocx } from "./docxSynthetic"
import { pagedEditor, savedDigest, pageDigest, settleAll } from "./pagedFixtures"

const decoder = new TextDecoder()
const nodes = docxSchema.nodes

const STYLES = STYLES_PART(
	'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
		'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
		'<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
		'<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>' +
		'<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'
)

const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

function heading(level: number, text: string): string {
	return p(text, `<w:pStyle w:val="Heading${String(level)}"/>`)
}

function spec(body: string): SyntheticDocx {
	return { body, parts: [STYLES], relationships: [{ id: "rIdS", type: "styles", target: "styles.xml" }] }
}

const HEADINGS =
	p("Contents") +
	heading(1, "Introduction") +
	p("Text one.") +
	pageBreak +
	heading(2, "Background") +
	p("Text two.") +
	heading(4, "Too deep") +
	pageBreak +
	heading(1, "Results") +
	p("Text three.")

interface Opened {
	readonly editor: PagedEditor
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
}

async function open(source: SyntheticDocx | Uint8Array): Promise<Opened> {
	const opened = await pagedEditor(source instanceof Uint8Array ? source : buildDocx(source))

	await settleAll(opened.editor)

	return opened
}

async function save(opened: Opened): Promise<Uint8Array> {
	return (await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: opened.editor.sessions })).bytes as Uint8Array
}

async function read(bytes: Uint8Array, name: string): Promise<string> {
	const pkg = await OpcPackage.open(bytesSource(bytes))

	return decoder.decode(await pkg.read(name))
}

function caretAt(editor: PagedEditor, pos: number): void {
	const session = editor.session

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, pos)))
}

// Where the text `text` starts in the body.
function find(doc: ProseMirrorNode, text: string): number {
	let found = -1

	doc.descendants((node, pos) => {
		if (found >= 0) return false
		if (node.isText && (node.text ?? "").includes(text)) found = pos + (node.text ?? "").indexOf(text)

		return true
	})

	return found
}

// The entries of the table of contents: text and page number, by paragraph.
function entries(doc: ProseMirrorNode, from: number, to: number): string[] {
	const out: string[] = []

	doc.nodesBetween(from, to, node => {
		if (node.type !== nodes.paragraph) return true

		let line = ""

		node.forEach(child => {
			if (child.isText) line += child.text ?? ""
			else if (child.type === nodes.tab) line += "\t"
			else if (child.type === nodes.field) line += `[${child.attrs["result"] as string}]`
		})

		out.push(line)

		return false
	})

	return out
}

describe("tables of contents", () => {
	it("inserts Word's default table from the headings, with links, bookmarks and page numbers", async () => {
		const opened = await open(spec(HEADINGS))
		const editor = opened.editor

		caretAt(editor, find(editor.session.state.doc, "Introduction"))

		expect(editor.canExecute("toc.insert")).toBe(true)
		expect(await editor.execute({ type: "toc.insert", levels: 3, hyperlinks: true })).toBe(true)
		await settleAll(editor)

		const session = editor.session
		const [region] = tocRegions(session)

		expect(region).toBeDefined()
		expect(region?.source).toBe(false)
		// The table took a page of its own: every heading moved on by one.
		expect(entries(session.state.doc, region?.from ?? 0, region?.to ?? 0)).toEqual([
			"Introduction\t[1]",
			"Background\t[2]",
			"Results\t[3]"
		])

		const bookmarks = sessionBookmarks(session).list.filter(bookmark => bookmark.name.startsWith("_Toc"))

		expect(bookmarks).toHaveLength(3)

		const bytes = await save(opened)
		const xml = await read(bytes, "word/document.xml")

		expect(xml).toContain(
			'<w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'
		)
		expect(xml).toContain(`<w:hyperlink w:anchor="${bookmarks[0]?.name ?? ""}" w:history="1">`)
		expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d{9}"\/>/)

		// Reopened, the table is the file's: locked, as Word's tables are.
		const reopened = await open(bytes)
		const regions = tocRegions(reopened.editor.session)

		expect(regions).toHaveLength(1)
		expect(regions[0]?.source).toBe(true)
		expect(regions[0]?.updatable).toBe(true)
		expect(pageDigest(reopened.editor.layout)).toEqual(await savedDigest(bytes))
	})

	it("undoes an insertion with its page numbers in one step", async () => {
		const opened = await open(spec(HEADINGS))
		const editor = opened.editor
		const before = editor.session.state.doc

		caretAt(editor, 1)
		await editor.execute({ type: "toc.insert", levels: 3, hyperlinks: true })
		await settleAll(editor)

		expect(editor.session.state.doc.eq(before)).toBe(false)
		expect(await editor.execute({ type: "history.undo" })).toBe(true)
		expect(editor.session.state.doc.eq(before)).toBe(true)
	})

	it("rebuilds the file's table from the headings as they are now, keeping its bookmarks and entry properties", async () => {
		const entry = (name: string, text: string, page: string): string =>
			'<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>' +
			`<w:hyperlink w:anchor="${name}" w:history="1"><w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r>` +
			`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${name} \\h </w:instrText></w:r>` +
			`<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${page}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>`
		const body =
			'<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>' +
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
			entry("_Toc100000001", "Introduction", "7").slice(entry("", "", "").indexOf("<w:hyperlink")) +
			"</w:p>" +
			entry("_Toc100000002", "Results", "9") +
			'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
			pageBreak +
			'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc100000001"/><w:r><w:t>Introduction</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
			p("Text.") +
			'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="2" w:name="_Toc100000002"/><w:r><w:t>Results</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>'
		const opened = await open(spec(body))
		const editor = opened.editor
		const session = editor.session
		const [before] = tocRegions(session)

		expect(before?.source).toBe(true)
		expect(before?.updatable).toBe(true)

		// A heading added between the two.
		const at = find(session.state.doc, "Text.")

		session.dispatch(session.state.tr.insert(at + 5, nodes.paragraph.create({ styleId: "Heading2" }, docxSchema.text("Method"))))
		caretAt(editor, 1)
		expect(editor.canExecute("toc.update")).toBe(true)
		expect(await editor.execute({ type: "toc.update", pagesOnly: false })).toBe(true)
		await settleAll(editor)

		const [region] = tocRegions(session)

		expect(region?.source).toBe(false)
		expect(entries(session.state.doc, region?.from ?? 0, region?.to ?? 0)).toEqual(["Introduction\t[2]", "Method\t[2]", "Results\t[2]"])

		const names = sessionBookmarks(session)
			.list.filter(bookmark => bookmark.name.startsWith("_Toc"))
			.map(bookmark => bookmark.name)

		// The headings' bookmarks are reused; the new heading gets one.
		expect(names).toHaveLength(3)
		expect(names[0]).toBe("_Toc100000001")
		expect(names[2]).toBe("_Toc100000002")

		const bytes = await save(opened)
		const xml = await read(bytes, "word/document.xml")

		// The level 1 entries keep the table's own paragraph properties (its right tab with dots); the level 2 entry
		// (no entry of its level to copy, its style without a right tab) gets Word's direct right tab at the margin.
		expect(xml.match(/<w:tab w:val="right" w:leader="dot" w:pos="9350"\/>/g)).toHaveLength(3)
		expect(xml).toContain(`<w:hyperlink w:anchor="${names[1] ?? ""}" w:history="1">`)

		const reopened = await open(bytes)

		expect(tocRegions(reopened.editor.session)[0]?.source).toBe(true)
		expect(pageDigest(reopened.editor.layout)).toEqual(await savedDigest(bytes))
	})

	it("deletes a table whole when an edit takes its end", async () => {
		const opened = await open(spec(HEADINGS))
		const editor = opened.editor
		const session = editor.session

		caretAt(editor, 1)
		await editor.execute({ type: "toc.insert", levels: 2, hyperlinks: false })
		await settleAll(editor)

		const [region] = tocRegions(session)
		let end = -1

		session.state.doc.nodesBetween(region?.from ?? 0, region?.to ?? 0, (node, pos) => {
			if (node.type === nodes.field_char && node.attrs["kind"] === "end") end = pos

			return true
		})

		// Backspace over the end character.
		expect(session.dispatch(session.state.tr.delete(end, end + 1))).toBe(true)
		expect(tocRegions(session)).toHaveLength(0)

		let chars = 0

		session.state.doc.descendants(node => {
			if (node.type === nodes.field_char) chars++

			return true
		})

		expect(chars).toBe(0)
		expect(find(session.state.doc, "Introduction")).toBeGreaterThan(0)
		expect(await editor.execute({ type: "history.undo" })).toBe(true)
		expect(tocRegions(session)).toHaveLength(1)
	})

	it("inserts a table as a tracked insertion while tracking, or untracked while tracked changes hold no fields", async () => {
		const opened = await open(spec(HEADINGS))
		const editor = opened.editor

		opened.context.identity.author = "Ann Lee"
		opened.context.review.tracking = true
		caretAt(editor, find(editor.session.state.doc, "Introduction"))
		expect(await editor.execute({ type: "toc.insert", levels: 3, hyperlinks: true })).toBe(true)
		await settleAll(editor)

		const bytes = await save(opened)
		const xml = await read(bytes, "word/document.xml")

		// Recorded whole (field characters and fields in w:ins), or not at all: never half.
		if (markable(nodes.field_char.create()) && markable(nodes.field.create())) {
			expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Ann Lee"[^>]*><w:r><w:fldChar w:fldCharType="begin"\/>/)
		} else {
			expect(xml).not.toContain("<w:ins ")
			expect(xml).toContain(
				'<w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'
			)
		}

		expect(pageDigest((await open(bytes)).editor.layout)).toEqual(await savedDigest(bytes))
	})

	it("keeps a table of one entry over two paragraphs, its end in the second", async () => {
		const opened = await open(spec(p("Start") + heading(1, "Chapter") + heading(2, "Section")))
		const editor = opened.editor
		const session = editor.session

		caretAt(editor, 1)
		expect(await editor.execute({ type: "toc.insert", levels: 1, hyperlinks: false })).toBe(true)
		await settleAll(editor)

		const [region] = tocRegions(session)

		// One paragraph would read back as a field of that paragraph, not a table.
		expect(entries(session.state.doc, region?.from ?? 0, region?.to ?? 0)).toEqual(["Chapter\t[1]", ""])

		const bytes = await save(opened)

		expect(tocRegions((await open(bytes)).editor.session)[0]?.source).toBe(true)
	})

	it("reads Word's switches", () => {
		const options = tocOptions(parseFieldCode(' TOC \\o "2-4" \\t "Appendix,1;Annex;Quote,3" \\n "3-9" \\p "-" \\h \\b Part '))

		expect(options.headings).toEqual([2, 4])
		expect([...options.styles]).toEqual([
			["appendix", 1],
			["annex", 1],
			["quote", 3]
		])
		expect(options.omit).toEqual([3, 9])
		expect(options.separator).toBe("-")
		expect(options.hyperlinks).toBe(true)
		expect(options.bookmark).toBe("Part")
		expect(options.supported).toBe(true)
		expect(tocOptions(parseFieldCode(" TOC \\f \\l 1-2 ")).supported).toBe(false)
	})
})
