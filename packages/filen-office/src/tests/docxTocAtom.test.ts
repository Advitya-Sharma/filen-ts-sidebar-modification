import type { Node as ProseMirrorNode } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { docxSchema, type DocxEditContext } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import type { DocxPackage } from "@filen/office/core/docx"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { tocRegions } from "../core/docx/fields/toc"
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

function heading(level: number, text: string): string {
	return p(text, `<w:pStyle w:val="Heading${String(level)}"/>`)
}

function spec(body: string): SyntheticDocx {
	return { body, parts: [STYLES], relationships: [{ id: "rIdS", type: "styles", target: "styles.xml" }] }
}

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

// An empty table of contents, as Word leaves one (its field within one paragraph, read as one field atom): Update
// Table rebuilds it, in one undo step, recorded like any other rebuild while tracking.

const EMPTY_TOC =
	'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
	'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>No table of contents entries found.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
const BODY = p("Contents") + EMPTY_TOC + heading(1, "Introduction") + p("Text one.") + heading(2, "Background")

function tables(editor: PagedEditor): string[][] {
	const session = editor.session

	return tocRegions(session).map(region => entries(session.state.doc, region.from, region.to))
}

describe("an empty table of contents read as one field", () => {
	it("is a table: a press finds it whole, the snapshot offers Update Table, an unedited save is unchanged", async () => {
		const bytes = buildDocx(spec(BODY))
		const opened = await open(bytes)
		const editor = opened.editor
		const doc = editor.session.state.doc
		const [region] = tocRegions(editor.session)

		expect(region).toMatchObject({ from: doc.child(0).nodeSize, atom: doc.child(0).nodeSize + 1, updatable: true })
		caretAt(editor, region?.atom ?? 0)
		expect(editor.snapshot().context.field).toMatchObject({ type: "TOC", toc: true, refusal: null })
		expect(editor.snapshot().enabled["toc.update"]).toBe(true)
		expect(await read(await save(opened), "word/document.xml")).toBe(await read(bytes, "word/document.xml"))
	})

	it("is rebuilt by Update Table in one undo step, saved through Gate 2 and laid out as reopened", async () => {
		const opened = await open(spec(BODY))
		const editor = opened.editor
		const session = editor.session
		const original = session.state.doc

		caretAt(editor, find(original, "Contents"))
		// Page numbers only: an empty table has none.
		expect(await editor.execute({ type: "toc.update", pagesOnly: true })).toBe(true)
		expect(session.state.doc).toBe(original)
		expect(await editor.execute({ type: "toc.update", pagesOnly: false })).toBe(true)
		await settleAll(editor)
		expect(tables(editor)).toEqual([["Introduction\t[1]", "Background\t[1]"]])

		const bytes = await save(opened)
		const xml = await read(bytes, "word/document.xml")

		expect(xml).not.toContain("No table of contents entries found.")
		expect(xml).toMatch(/<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u <\/w:instrText>/)
		// The editor lays it out as a fresh open of what it saved does.
		expect(pageDigest(editor.layout)).toEqual(await savedDigest(bytes))
		expect(pageDigest((await open(bytes)).editor.layout)).toEqual(await savedDigest(bytes))
		expect(tables((await open(bytes)).editor)).toEqual([["Introduction\t[1]", "Background\t[1]"]])

		expect(await editor.execute({ type: "history.undo" })).toBe(true)
		expect(session.state.doc.eq(original)).toBe(true)
	})

	it("is rebuilt while tracking: the old field deleted whole, the new table inserted; rejected, as it was", async () => {
		const opened = await open(spec(BODY))
		const editor = opened.editor
		const session = editor.session

		opened.context.identity.author = "Ann Lee"
		opened.context.review.tracking = true
		caretAt(editor, find(session.state.doc, "Contents"))
		expect(await editor.execute({ type: "toc.update", pagesOnly: false })).toBe(true)
		await settleAll(editor)

		let deleted = ""
		let inserted = ""

		session.state.doc.descendants(node => {
			if (node.type === nodes.field && docxSchema.marks.deletion.isInSet(node.marks) !== undefined)
				deleted += node.attrs["result"] as string
			if (node.isText && docxSchema.marks.insertion.isInSet(node.marks) !== undefined) inserted += node.text ?? ""

			return true
		})

		expect(deleted).toBe("No table of contents entries found.")
		expect(inserted).toBe("IntroductionBackground")

		const bytes = await save(opened)

		expect(await read(bytes, "word/document.xml")).toMatch(/<w:del w:id="\d+" w:author="Ann Lee"/)
		expect(pageDigest(editor.layout)).toEqual(await savedDigest(bytes))
		expect(pageDigest((await open(bytes)).editor.layout)).toEqual(await savedDigest(bytes))

		for (const accept of [true, false]) {
			const copy = await open(bytes)

			expect(await copy.editor.execute({ type: accept ? "review.acceptAll" : "review.rejectAll" })).toBe(true)

			const resolved = await save(copy)
			const xml = await read(resolved, "word/document.xml")

			expect(xml).not.toMatch(/<w:(del|ins) /)
			expect(xml.includes("No table of contents entries found.")).toBe(!accept)
			expect(xml.includes('<w:hyperlink w:anchor="_Toc')).toBe(accept)
			expect(pageDigest((await open(resolved)).editor.layout)).toEqual(await savedDigest(resolved))
		}
	})
})
