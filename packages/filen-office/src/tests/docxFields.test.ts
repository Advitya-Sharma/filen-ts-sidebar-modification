import type { Node as ProseMirrorNode } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { docxSchema, type DocxEditContext } from "@filen/office/core/docx/edit"
import { prepareDocxSave, saveDocx, serializeDocx } from "@filen/office/core/docx/write"
import { parseXml } from "@filen/office/core/xml"
import type { CheckInput } from "../core/docx/write/checks/index"
import { fieldChecks } from "../core/docx/write/checks/fields"
import type { DocxPackage } from "@filen/office/core/docx"
import { hitObject, type PagedEditor } from "@filen/office/core/docx/paged"
import { markInternal } from "../core/docx/edit/invariants/index"
import { sessionBookmarks } from "../core/docx/fields/bookmarks"
import { datePicture, numericPicture, parseFieldCode } from "../core/docx/fields/evaluate"
import { NAMESPACES, buildDocx, p, type SyntheticDocx } from "./docxSynthetic"
import { pageLines } from "./layoutFixtures"
import { pageDigest, pagedEditor, savedDigest, settleAll } from "./pagedFixtures"

const decoder = new TextDecoder()

const encoder = new TextEncoder()
const nodes = docxSchema.nodes

beforeEach(() => {
	vi.stubEnv("TZ", "UTC")
})

interface Opened {
	readonly editor: PagedEditor
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
}

async function open(source: SyntheticDocx | Uint8Array): Promise<Opened> {
	const opened = await pagedEditor(source instanceof Uint8Array ? source : buildDocx(source))

	opened.context.identity.now = () => Date.UTC(2026, 9, 12, 14, 5, 9)
	await settleAll(opened.editor)

	return opened
}

async function save(opened: Opened): Promise<Uint8Array> {
	return (await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: opened.editor.sessions })).bytes as Uint8Array
}

async function refusal(opened: Opened): Promise<string | undefined> {
	try {
		await prepareDocxSave({ pkg: opened.pkg, context: opened.context, sessions: opened.editor.sessions })

		return undefined
	} catch (error) {
		return (error as { reason: string }).reason
	}
}

async function read(bytes: Uint8Array, name = "word/document.xml"): Promise<string> {
	const pkg = await OpcPackage.open(bytesSource(bytes))

	return decoder.decode(await pkg.read(name))
}

function select(editor: PagedEditor, from: number, to: number = from): void {
	const session = editor.session

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, from, to)))
}

function find(doc: ProseMirrorNode, text: string): number {
	let found = -1

	doc.descendants((node, pos) => {
		if (found >= 0) return false
		if (node.isText && (node.text ?? "").includes(text)) found = pos + (node.text ?? "").indexOf(text)

		return true
	})

	return found
}

function fieldsOf(doc: ProseMirrorNode): { node: ProseMirrorNode; pos: number }[] {
	const out: { node: ProseMirrorNode; pos: number }[] = []

	doc.descendants((node, pos) => {
		if (node.type === nodes.field) out.push({ node, pos })

		return true
	})

	return out
}

// The fields of a document: type and result.
function fields(doc: ProseMirrorNode): string[] {
	const out: string[] = []

	doc.descendants(node => {
		if (node.type === nodes.field) out.push(`${node.attrs["type"] as string}=${node.attrs["result"] as string}`)

		return true
	})

	return out
}

const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

// A complex field as Word writes it.
function complex(instruction: string, result: string, rPr = "", options: { dirty?: boolean; lock?: boolean } = {}): string {
	const begin = `<w:fldChar w:fldCharType="begin"${options.dirty === true ? ' w:dirty="true"' : ""}${options.lock === true ? ' w:fldLock="true"' : ""}/>`

	return (
		`<w:r>${rPr}${begin}</w:r><w:r>${rPr}<w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>` +
		`<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>${result}</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
	)
}

describe("field codes and pictures", () => {
	it("parses types, arguments and switches (ECMA-376 §17.16.1)", () => {
		expect(parseFieldCode(" REF _Ref1 \\h \\* MERGEFORMAT ")).toEqual({
			type: "REF",
			args: ["_Ref1"],
			switches: [
				{ name: "h", arg: undefined },
				{ name: "*", arg: "MERGEFORMAT" }
			]
		})
		expect(parseFieldCode(' DATE \\@ "d \\"MMMM\\" yyyy" ').switches).toEqual([{ name: "@", arg: 'd "MMMM" yyyy' }])
		expect(parseFieldCode(" SEQ Figure \\r 4 \\* roman ").switches).toEqual([
			{ name: "r", arg: "4" },
			{ name: "*", arg: "roman" }
		])
	})

	it("formats dates with \\@ pictures (§17.16.4.1)", () => {
		const date = new Date(2026, 9, 2, 14, 5, 9)

		expect(datePicture(date, "dddd, d MMMM yyyy")).toBe("Friday, 2 October 2026")
		expect(datePicture(date, "dd/MM/yy")).toBe("02/10/26")
		expect(datePicture(date, "h:mm am/pm")).toBe("2:05 pm")
		expect(datePicture(date, "HH:mm:ss 'o''clock'")).toBe("14:05:09 oclock")
		expect(datePicture(date, "MMM d, yyyy h:mm AM/PM")).toBe("Oct 2, 2026 2:05 PM")
	})

	it("formats numbers with \\# pictures (§17.16.4.2)", () => {
		expect(numericPicture(1234.5, "#,##0.00")).toBe("1,234.50")
		expect(numericPicture(7, "00")).toBe("07")
		expect(numericPicture(-3, "0;(0)")).toBe("(3)")
		expect(numericPicture(0, "0;-0;'none'")).toBe("none")
		expect(numericPicture(42, "'No. '0")).toBe("No. 42")
	})
})

describe("updating fields", () => {
	it("recomputes page numbers, page counts, sequences, references and dates, and saves only what changed", async () => {
		const opened = await open({
			body:
				'<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>Old heading</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
				`<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${complex(" REF Target \\h ", "stale")}<w:r><w:t xml:space="preserve"> on page </w:t></w:r>${complex(" PAGEREF Target \\h ", "9")}</w:p>` +
				pageBreak +
				`<w:p>${complex(" SEQ Figure \\* ARABIC ", "7")}${complex(" SEQ Figure \\* ARABIC ", "7")}${complex(" SEQ Figure \\c ", "7")}</w:p>` +
				`<w:p>${complex(" PAGE ", "5", "<w:rPr><w:b/></w:rPr>")} of ${complex(" NUMPAGES \\* roman ", "x")}</w:p>` +
				`<w:p>${complex(' DATE \\@ "d MMMM yyyy" ', "1 January 2000")}${complex(" PAGE ", "8", "", { lock: true })}</w:p>`
		})
		const editor = opened.editor
		const session = editor.session

		// The bookmarked text changes; the reference shows it once updated.
		const at = find(session.state.doc, "Old heading")

		session.dispatch(session.state.tr.insertText("New heading", at, at + "Old heading".length))
		expect(await editor.execute({ type: "field.updateAll" })).toBe(true)

		expect(fields(session.state.doc)).toEqual([
			"REF=New heading",
			"PAGEREF=1",
			"SEQ=1",
			"SEQ=2",
			"SEQ=2",
			"PAGE=2",
			"NUMPAGES=ii",
			"DATE=12 October 2026",
			// Locked (w:fldLock): kept.
			"PAGE=8"
		])

		const bytes = await save(opened)
		const xml = await read(bytes)

		// The result written anew, the code and its run formatting kept; the locked field as it was.
		expect(xml).toContain(
			'<w:r><w:rPr><w:b/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
		)
		expect(xml).toContain('<w:fldChar w:fldCharType="begin" w:fldLock="true"/>')
		expect(xml).toContain("<w:t>8</w:t>")

		const reopened = await open(bytes)

		expect(fields(reopened.editor.session.state.doc)).toEqual(fields(session.state.doc))
		expect(pageDigest(reopened.editor.layout)).toEqual(await savedDigest(bytes))
	})

	it("updates the field at the caret only, as one undo step", async () => {
		const opened = await open({
			body: `<w:p>${complex(" PAGE ", "4")}<w:r><w:t xml:space="preserve"> and </w:t></w:r>${complex(" PAGE ", "4")}</w:p>`
		})
		const editor = opened.editor
		const doc = editor.session.state.doc

		select(editor, 1)
		expect(editor.canExecute("field.update")).toBe(true)
		expect(await editor.execute({ type: "field.update" })).toBe(true)
		expect(fields(editor.session.state.doc)).toEqual(["PAGE=1", "PAGE=4"])
		expect(await editor.execute({ type: "history.undo" })).toBe(true)
		expect(editor.session.state.doc.eq(doc)).toBe(true)
	})

	it("keeps w:dirty, markers in the old result and the field's own runs", async () => {
		const opened = await open({
			body:
				'<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
				'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:bookmarkStart w:id="3" w:name="inside"/><w:r><w:t>7</w:t></w:r><w:bookmarkEnd w:id="3"/>' +
				'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
		})

		expect(await opened.editor.execute({ type: "field.updateAll" })).toBe(true)

		const xml = await read(await save(opened))

		expect(xml).toContain('<w:fldChar w:fldCharType="begin" w:dirty="true"/>')
		expect(xml).toContain('<w:bookmarkStart w:id="3" w:name="inside"/><w:bookmarkEnd w:id="3"/><w:r><w:t>1</w:t></w:r>')
	})
})

describe("inserting fields", () => {
	it("inserts page numbers, page counts and dates as Word's complex fields", async () => {
		const opened = await open({ body: p("Page ") + pageBreak + p("End") })
		const editor = opened.editor

		select(editor, 6)
		expect(editor.execute({ type: "insert.pageNumber" })).toBe(true)
		expect(editor.execute({ type: "insert.date", format: "d/M/yyyy" })).toBe(true)
		await settleAll(editor)
		expect(editor.execute({ type: "insert.pageCount" })).toBe(true)
		expect(fields(editor.session.state.doc)).toEqual(["PAGE=1", "DATE=12/10/2026", "NUMPAGES=2"])

		const bytes = await save(opened)
		const xml = await read(bytes)

		expect(xml).toContain(
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
		)
		expect(pageDigest((await open(bytes)).editor.layout)).toEqual(await savedDigest(bytes))
	})

	it("inserts cross-references to bookmarks it has, and no others", async () => {
		const opened = await open({
			body:
				'<w:p><w:bookmarkStart w:id="1" w:name="Figure1"/><w:r><w:t>Figure 1</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' + p("See ")
		})
		const editor = opened.editor
		const end = editor.session.state.doc.content.size - 1

		select(editor, end)
		expect(editor.execute({ type: "insert.field", instruction: "REF Missing \\h" })).toBe(false)
		expect(editor.execute({ type: "insert.field", instruction: "REF Figure1 \\h" })).toBe(true)
		expect(fields(editor.session.state.doc)).toEqual(["REF=Figure 1"])
		expect(await refusal(opened)).toBeUndefined()
		// Word keeps a reference whose bookmark goes (an update shows its error text).
		expect(editor.execute({ type: "bookmark.delete", name: "Figure1" })).toBe(true)
		expect(await refusal(opened)).toBeUndefined()
		expect(await editor.execute({ type: "field.updateAll" })).toBe(true)
		expect(fields(editor.session.state.doc)).toEqual(["REF=Error! Bookmark not defined."])
	})
})

describe("tracked updates", () => {
	it("records an update while tracking as Word does (old result deleted, new inserted), accepted or rejected whole", async () => {
		const body = `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${complex(" PAGE ", "9", "<w:rPr><w:b/></w:rPr>")}</w:p>`
		const opened = await open({ body })
		const editor = opened.editor
		const session = editor.session

		opened.context.identity.author = "Ann Lee"
		opened.context.review.tracking = true
		expect(await editor.execute({ type: "field.updateAll" })).toBe(true)

		const [field] = fieldsOf(session.state.doc)

		expect(field?.node.attrs["result"]).toBe("1")
		expect(field?.node.attrs["update"]).toMatchObject({ result: "9", author: "Ann Lee" })
		expect(session.snapshot().review.revisionCount).toBe(1)

		// Updated again: still one change, from the result before the first.
		expect(await editor.execute({ type: "field.updateAll" })).toBe(true)
		expect(fieldsOf(session.state.doc)[0]?.node.attrs["update"]).toMatchObject({ result: "9" })

		const bytes = await save(opened)
		const xml = await read(bytes)

		expect(xml).toMatch(
			/<w:fldChar w:fldCharType="separate"\/><\/w:r><w:del w:id="\d+" w:author="Ann Lee" w:date="[^"]+"><w:r><w:rPr><w:b\/><\/w:rPr><w:delText>9<\/w:delText><\/w:r><\/w:del><w:ins w:id="\d+" w:author="Ann Lee" w:date="[^"]+"><w:r><w:rPr><w:b\/><\/w:rPr><w:t>1<\/w:t><\/w:r><\/w:ins><w:r><w:rPr><w:b\/><\/w:rPr><w:fldChar w:fldCharType="end"\/>/
		)

		// Reopened, the update is the field's again, and resolves as one.
		const reopened = await open(bytes)
		const again = reopened.editor.session

		expect(pageDigest(reopened.editor.layout)).toEqual(await savedDigest(bytes))
		expect(fieldsOf(again.state.doc)[0]?.node.attrs).toMatchObject({ result: "1", update: { result: "9", author: "Ann Lee" } })
		expect(again.snapshot().review.revisionCount).toBe(1)
		expect(await reopened.editor.execute({ type: "review.rejectAll" })).toBe(true)
		expect(fieldsOf(again.state.doc)[0]?.node.attrs).toMatchObject({ result: "9", update: null })
		// As the file had it before the update.
		expect(await read(await save(reopened))).toContain(body)

		expect(await reopened.editor.execute({ type: "history.undo" })).toBe(true)
		expect(await reopened.editor.execute({ type: "review.acceptAll" })).toBe(true)
		expect(fieldsOf(again.state.doc)[0]?.node.attrs).toMatchObject({ result: "1", update: null })

		const accepted = await read(await save(reopened))

		expect(accepted).not.toContain("<w:del ")
		expect(accepted).toContain("<w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r>")
	})

	it("leaves fields deleted while tracking as they are when every field is updated", async () => {
		const figure = (result: string, text: string): string =>
			`<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r>${complex(" SEQ Figure ", result)}<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>`
		const page = `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${complex(" PAGE ", "7")}</w:p>`
		const opened = await open({ body: page + figure("2", "second") + figure("3", "third") })
		const editor = opened.editor
		const session = editor.session

		opened.context.identity.author = "Ann Lee"
		opened.context.review.tracking = true

		const [first] = fieldsOf(session.state.doc)

		select(editor, first?.pos ?? 0, (first?.pos ?? 0) + 1)
		session.dispatch(session.state.tr.deleteSelection())
		expect(await editor.execute({ type: "field.updateAll" })).toBe(true)

		const fields = fieldsOf(session.state.doc)

		// The deleted one keeps its result and records no update.
		expect(fields.map(field => field.node.attrs["result"])).toEqual(["7", "1", "2"])
		expect(fields[0]?.node.attrs["update"]).toBeNull()
		expect(docxSchema.marks.deletion.isInSet(fields[0]?.node.marks ?? [])).toBeDefined()
		expect(await refusal(opened)).toBeUndefined()

		const bytes = await save(opened)
		const reopened = await open(bytes)

		expect(pageDigest(reopened.editor.layout)).toEqual(await savedDigest(bytes))
		expect(fieldsOf(reopened.editor.session.state.doc).map(field => field.node.attrs["result"])).toEqual(["7", "1", "2"])
	})

	it("leaves a field inserted while tracking to its insertion, and needs the author's name", async () => {
		const opened = await open({ body: p("Text") })
		const editor = opened.editor

		opened.context.review.tracking = true
		select(editor, 3)
		expect(await editor.execute({ type: "field.updateAll" })).toBe(false)

		opened.context.identity.author = "Ann Lee"
		expect(await editor.execute({ type: "insert.pageCount" })).toBe(true)
		expect(await editor.execute({ type: "field.updateAll" })).toBe(true)
		expect(fieldsOf(editor.session.state.doc)[0]?.node.attrs["update"]).toBeNull()
		expect(await refusal(opened)).toBeUndefined()
	})
})

describe("field codes view", () => {
	const text = (editor: PagedEditor): string[] => pageLines(editor.layout, 0).map(line => line.text)

	it("shows every field as its code (Alt+F9) and one field at a time (Shift+F9), a view the file never sees", async () => {
		const opened = await open({
			body:
				`<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${complex(" PAGE \\* MERGEFORMAT ", "1")}<w:r><w:t xml:space="preserve"> of </w:t></w:r>` +
				'<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> see </w:t></w:r>' +
				`${complex(" REF Target \\h ", "the target")}</w:p>` +
				'<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>the target</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>'
		})
		const editor = opened.editor
		const before = pageDigest(editor.layout)

		expect(text(editor)[0]).toBe("Page 1 of 1 see the target")
		expect(editor.canExecute("view.fieldCodes")).toBe(true)
		expect(await editor.execute({ type: "view.fieldCodes" })).toBe(true)
		await settleAll(editor)
		expect(editor.fieldCodes).toBe(true)
		expect(text(editor)[0]).toBe("Page { PAGE \\* MERGEFORMAT } of { NUMPAGES } see { REF Target \\h }")

		// Typing after a field shown as its code lands after the field.
		const [page, , ref] = fieldsOf(editor.session.state.doc)

		select(editor, (ref?.pos ?? 0) + 1)
		editor.session.dispatch(editor.session.state.tr.insertText("!"))
		await settleAll(editor)
		expect(text(editor)[0]).toBe("Page { PAGE \\* MERGEFORMAT } of { NUMPAGES } see { REF Target \\h }!")

		// Shift+F9 on the page number shows its result again.
		select(editor, page?.pos ?? 0)
		expect(await editor.execute({ type: "view.fieldCode" })).toBe(true)
		await settleAll(editor)
		expect(text(editor)[0]).toBe("Page 1 of { NUMPAGES } see { REF Target \\h }!")

		expect(await editor.execute({ type: "view.fieldCodes" })).toBe(true)
		await settleAll(editor)
		expect(editor.fieldCodes).toBe(false)
		expect(text(editor)[0]).toBe("Page 1 of 1 see the target!")

		// One field alone, in the results view.
		select(editor, (fieldsOf(editor.session.state.doc)[2]?.pos ?? 0) + 1)
		expect(await editor.execute({ type: "view.fieldCode" })).toBe(true)
		await settleAll(editor)
		expect(text(editor)[0]).toBe("Page 1 of 1 see { REF Target \\h }!")

		// What a save writes and its fresh layout are the results view's.
		expect(editor.setFieldCodes(false)).toBe(true)
		await settleAll(editor)

		const bytes = await save(opened)

		expect(pageDigest(editor.layout)).toEqual(await savedDigest(bytes))
		expect(before).not.toEqual(pageDigest(editor.layout))
		expect(await read(bytes)).not.toContain("{")
	})

	it("lays out again from the page a field switched one by one is on, as switching every field would show it", async () => {
		const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
		const body =
			["One", "Two", "Three", "Four"].map(text => p(text) + pageBreak).join("") +
			`<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${complex(" PAGE ", "3")}<w:r><w:t xml:space="preserve"> of the text</w:t></w:r></w:p>`
		const switched = async (whole: boolean): Promise<{ digest: string[]; kept: boolean[] }> => {
			const opened = await open({ body })
			const editor = opened.editor
			const before = [...editor.layout.pages]
			const [field] = fieldsOf(editor.session.state.doc)

			select(editor, field?.pos ?? 0)

			if (whole) expect(editor.toggleFieldCodes(field === undefined ? [] : [field.node])).toBe(true)
			else expect(await editor.execute({ type: "view.fieldCode" })).toBe(true)

			await settleAll(editor)

			return { digest: pageDigest(editor.layout), kept: before.map((page, index) => editor.layout.pages[index] === page) }
		}
		const one = await switched(false)
		const all = await switched(true)

		expect(one.digest).toEqual(all.digest)
		expect(one.digest[4]).toContain("{ PAGE }")
		// Laid out again from the page the field is on (or the one before, which it might now fit at the end of).
		expect(one.kept).toEqual([true, true, true, false, false])
		expect(all.kept).toEqual([false, false, false, false, false])
	})

	it("shows a table of contents spanning paragraphs as its code alone", async () => {
		const opened = await open({
			body:
				'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>' +
				'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>One</w:t></w:r></w:p>' +
				'<w:p><w:r><w:t>Two</w:t></w:r></w:p><w:p><w:r><w:t>Three</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
				p("After")
		})
		const editor = opened.editor

		expect(text(editor)).toEqual(["One", "Two", "Three", "After"])
		await editor.execute({ type: "view.fieldCodes" })
		await settleAll(editor)
		expect(text(editor)).toEqual(['{ TOC \\o "1-3" }', "After"])

		// Its paragraphs edited (laid out from the edit) stay under the code, the first and last too.
		for (const entry of ["Two", "One", "Three"]) {
			select(editor, find(editor.session.state.doc, entry) + 1)
			editor.session.dispatch(editor.session.state.tr.insertText("x"))
			await settleAll(editor)
			expect(text(editor)).toEqual(['{ TOC \\o "1-3" }', "After"])
		}

		await editor.execute({ type: "view.fieldCodes" })
		await settleAll(editor)
		expect(text(editor)).toEqual(["Oxne", "Txwo", "Txhree", "After"])
	})
})

describe("field hits", () => {
	it("finds a field under a press, and a table of contents whole for its Update Table control", async () => {
		const opened = await open({
			body:
				'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>' +
				'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Entry one</w:t></w:r></w:p>' +
				'<w:p><w:r><w:t>Entry two</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
				`<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${complex(" PAGE ", "1")}<w:r><w:t xml:space="preserve"> of the text</w:t></w:r></w:p>`
		})
		const editor = opened.editor
		const doc = editor.session.state.doc
		const center = (from: number, to: number): [number, number] => {
			const rect = editor.rangeRects(from, to, 0)[0]

			return [(rect?.x ?? 0) + (rect?.width ?? 0) / 2, (rect?.y ?? 0) + (rect?.height ?? 0) / 2]
		}
		const [field] = fieldsOf(doc)
		const at = field?.pos ?? 0

		expect(hitObject(editor, 0, ...center(at, at + 1))).toMatchObject({ kind: "field", pos: at, story: editor.activeStory })

		const entry = find(doc, "Entry two")
		const toc = hitObject(editor, 0, ...center(entry, entry + 5))

		expect(toc).toMatchObject({ kind: "field", part: "toc", pos: 0 })
		// The box holds both entries' lines.
		expect(toc?.box?.height ?? 0).toBeGreaterThan(20)
		expect(hitObject(editor, 0, ...center(find(doc, "of the text") + 3, find(doc, "of the text") + 6))).toBeUndefined()
	})
})

describe("bookmarks", () => {
	it("adds, moves, renames and deletes bookmarks, keeping names unique", async () => {
		const opened = await open({
			body: '<w:p><w:bookmarkStart w:id="0" w:name="Kept"/><w:r><w:t>Alpha beta gamma</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>'
		})
		const editor = opened.editor
		const session = editor.session
		const ranges = (): string[] =>
			sessionBookmarks(session).list.map(bookmark => `${bookmark.name}:${session.state.doc.textBetween(bookmark.from, bookmark.to)}`)
		const beta = find(session.state.doc, "beta")

		select(editor, beta, beta + 4)
		expect(editor.execute({ type: "bookmark.add", name: "Beta" })).toBe(true)
		expect(ranges()).toEqual(["Kept:Alpha beta gamma", "Beta:beta"])

		// Word's Add with a name the document has moves that bookmark.
		const gamma = find(session.state.doc, "gamma")

		select(editor, gamma, gamma + 5)
		expect(editor.execute({ type: "bookmark.add", name: "kept" })).toBe(true)
		expect(ranges()).toEqual(["Beta:beta", "Kept:gamma"])
		expect(editor.execute({ type: "bookmark.rename", name: "Beta", to: "Delta" })).toBe(true)
		expect(editor.execute({ type: "bookmark.rename", name: "Delta", to: "Kept" })).toBe(false)

		let xml = await read(await save(opened))

		expect(xml).toMatch(/<w:bookmarkStart w:id="1" w:name="Delta"\/><w:r><w:t>beta<\/w:t><\/w:r><w:bookmarkEnd w:id="1"\/>/)
		expect(xml).toMatch(/<w:bookmarkStart w:id="0" w:name="Kept"\/><w:r><w:t>gamma<\/w:t><\/w:r><w:bookmarkEnd w:id="0"\/>/)
		expect(editor.execute({ type: "bookmark.delete", name: "Kept" })).toBe(true)
		xml = await read(await save(opened))
		expect(xml).not.toContain("Kept")
		expect(xml).toContain('w:name="Delta"')
	})
})

describe("snapshot", () => {
	it("tells the field at the caret, the same object while it is the same field", async () => {
		const opened = await open({
			body: `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${complex(" PAGE ", "1")}<w:r><w:t xml:space="preserve"> and </w:t></w:r>${complex(" HYPERLINK \\l x ", "link")}</w:p>`
		})
		const editor = opened.editor
		const field = fieldsOf(editor.session.state.doc)

		select(editor, field[0]?.pos ?? 0)

		const first = editor.snapshot().context.field

		expect(first).toEqual({ type: "PAGE", updatable: true, toc: false, instruction: " PAGE ", refusal: null })
		select(editor, (field[0]?.pos ?? 0) + 1)
		// Right after the field: the same one.
		expect(editor.snapshot().context.field).toBe(first)
		select(editor, (field[1]?.pos ?? 0) + 1)
		expect(editor.snapshot().context.field).toEqual({
			type: "HYPERLINK",
			updatable: false,
			toc: false,
			instruction: " HYPERLINK \\l x ",
			refusal: null
		})
		select(editor, 2)
		expect(editor.snapshot().context.field).toBeNull()
	})

	it("tells why a table of contents kept as the file's locked blocks cannot be updated", async () => {
		const begin =
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
		const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
		// A field inside it spanning two of its paragraphs: the table stays locked blocks.
		const spanning =
			`<w:p>${begin}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "x" </w:instrText></w:r>` +
			'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>One</w:t></w:r></w:p>' +
			`<w:p><w:r><w:t>Two</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>${end}</w:p>`
		const table =
			`<w:p>${begin}<w:r><w:t>One</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${p("Cell")}</w:tc></w:tr></w:tbl>` +
			`<w:p><w:r><w:t>Two</w:t></w:r>${end}</w:p>`
		const cases = [
			{ body: spanning, untracked: null, tracked: "trackedLockedTable" },
			{ body: spanning.replace('TOC \\o "1-3"', "TOC \\f"), untracked: "unsupportedTable", tracked: "unsupportedTable" },
			{ body: table, untracked: "tableWithOtherContent", tracked: "tableWithOtherContent" }
		] as const

		for (const entry of cases) {
			const opened = await open({ body: entry.body + p("Heading", '<w:outlineLvl w:val="0"/>') })
			const editor = opened.editor

			opened.context.identity.author = "Ann Lee"
			select(editor, 0)
			expect(editor.session.state.doc.child(0).type).toBe(nodes.locked_block)

			const untracked = editor.snapshot().context.field

			expect(untracked).toMatchObject({ type: "TOC", toc: true, updatable: entry.untracked === null, refusal: entry.untracked })
			opened.context.review.tracking = true
			editor.session.dispatch(editor.session.state.tr)
			expect(editor.snapshot().context.field).toMatchObject({ updatable: false, refusal: entry.tracked })
			expect(await editor.execute({ type: "toc.update", pagesOnly: false })).toBe(false)
			opened.context.review.tracking = false
			editor.session.dispatch(editor.session.state.tr)
			expect(editor.snapshot().context.field).toBe(untracked)
			expect(await editor.execute({ type: "toc.update", pagesOnly: false })).toBe(entry.untracked === null)
		}
	})
})

describe("gate 2", () => {
	it("refuses field characters without partners and duplicate bookmark names", async () => {
		const cases: { readonly label: string; readonly node: () => ProseMirrorNode; readonly reason: string }[] = [
			{ label: "stray begin", node: () => nodes.field_char.create({ kind: "begin", instruction: " TOC " }), reason: "content" },
			{
				label: "duplicate name",
				node: () => nodes.marker.create({ kind: "bookmarkStart", id: "77", name: "Kept" }),
				reason: "duplicateId"
			}
		]

		for (const entry of cases) {
			const opened = await open({
				body: '<w:p><w:bookmarkStart w:id="0" w:name="Kept"/><w:r><w:t>Alpha</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>'
			})
			const session = opened.editor.session
			const tr = session.state.tr.insert(find(session.state.doc, "Alpha") + 2, entry.node())

			if (entry.label === "duplicate name") tr.insert(tr.doc.content.size - 1, nodes.marker.create({ kind: "bookmarkEnd", id: "77" }))

			session.dispatch(markInternal(tr))
			expect(await refusal(opened), entry.label).toBe(entry.reason)
		}
	})

	it("refuses a save whose writer dropped a kept field's w:dirty", async () => {
		const opened = await open({
			body: `<w:p>${complex(" PAGE ", "1", "", { dirty: true })}<w:r><w:t xml:space="preserve"> x</w:t></w:r></w:p>`
		})
		const session = opened.editor.session
		const at = find(session.state.doc, " x")

		session.dispatch(session.state.tr.insertText("y", at + 2))

		expect(await refusal(opened)).toBeUndefined()

		// The same save with the flag taken out of the written markup, as SuperDoc's writer does.
		const [check] = fieldChecks
		const changes = await serializeDocx({ pkg: opened.pkg, context: opened.context, sessions: opened.editor.sessions })
		const story = changes.stories[0]

		if (story === undefined || check?.part === undefined || check.done === undefined) throw new Error("No story written")

		const bytes = story.output.bytes
		const joined = new Uint8Array(bytes.length)

		if (bytes instanceof Uint8Array) joined.set(bytes)
		else
			bytes.visit(0, bytes.length, (piece, from, count, at) => {
				joined.set(piece.subarray(from, from + count), at)

				return true
			})

		const markup = decoder.decode(joined)
		const input = {
			original: opened.pkg,
			context: opened.context,
			changes,
			ids: () => 1,
			added: () => []
		} as unknown as CheckInput

		await check.part(input, {
			name: "word/document.xml",
			part: story.part,
			document: parseXml(encoder.encode(`<w:body ${NAMESPACES}>${markup.replace(' w:dirty="true"', "")}</w:body>`)),
			skeleton: true
		})
		expect(() => check.done?.(input)).toThrow("w:dirty")
	})
})
