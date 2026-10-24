import { strFromU8, unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { readRtf } from "../core/rtf/index"
import { encodeText } from "../core/rtf/write/encode"
import { proveRtf, writeRtf } from "../core/rtf/write/writer"
import { editRtf, latin1, paragraphs, saveRtf, scriptedEdits } from "./rtfWriteFixtures"

async function bodyXml(bytes: Uint8Array): Promise<string> {
	const read = await readRtf(bytes)

	if (!read.ok) throw new Error(read.reason)

	return strFromU8(unzipSync(read.docx)["word/document.xml"] ?? new Uint8Array())
}

function text(xml: string): string {
	return [...xml.matchAll(/<w:p[ >].*?<\/w:p>|<w:p\/>/gu)]
		.map(paragraph => [...paragraph[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)].map(match => match[1]).join(""))
		.join("\n")
}

const DOC = latin1(
	"{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}{\\f1\\fswiss Arial;}}{\\colortbl;\\red255\\green0\\blue0;}" +
		"\\pard\\plain\\f0\\fs24 First paragraph here\\par\n" +
		"\\pard\\plain\\f0\\fs24 Second {\\b bold} paragraph\\par\n" +
		"\\pard\\plain\\f1\\fs20\\cf1 Third one in red\\par\n" +
		"\\pard\\plain\\f0\\fs24 Fourth paragraph text\\par\n" +
		"\\trowd\\trgaph108\\trleft-108\\cellx3000\\cellx6000\\pard\\plain\\intbl A1 cell\\cell B1 cell\\cell\\row\n" +
		"\\pard\\plain\\f0\\fs24 Fifth caf\\'e9 paragraph\\par\n" +
		"}"
)

// Enough paragraphs and a table for every scripted edit.
const SCRIPT = latin1(
	"{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}}" +
		["one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven"]
			.map(word => `\\pard\\plain\\f0\\fs24 Paragraph ${word} of the script\\par\n`)
			.join("") +
		"\\trowd\\trgaph108\\trleft-108\\cellx3000\\cellx6000\\pard\\plain\\intbl A1 cell\\cell B1 cell\\cell\\row\n" +
		"\\trowd\\trgaph108\\trleft-108\\cellx3000\\cellx6000\\pard\\plain\\intbl A2 cell\\cell B2 cell\\cell\\row\n" +
		"\\pard\\plain\\f0\\fs24 Paragraph twelve stays as it was\\par\n}"
)

describe("RTF text encoding", () => {
	it("escapes RTF syntax and writes other characters as \\uN with an ANSI fallback, or ? where the code page has none", () => {
		expect(encodeText("a\\b{c}", 1252).rtf).toBe("a\\\\b\\{c\\}")
		expect(encodeText("é€", 1252)).toEqual({ rtf: "\\u233\\'e9\\u8364\\'80", unicode: true })
		expect(encodeText("Ж", 1252).rtf).toBe("\\u1046?")
		expect(encodeText("Ж", 1251).rtf).toBe("\\u1046\\'c6")
		// Past U+7FFF a signed value; past U+FFFF the UTF-16 pair.
		expect(encodeText("", 1252).rtf).toBe("\\u-3913?")
		expect(encodeText("😀", 1252).rtf).toBe("\\u-10179?\\u-8704?")
	})
})

describe("RTF writer", () => {
	it("writes an untouched document byte for byte", async () => {
		const editing = await editRtf(DOC)
		const written = await saveRtf(editing)

		expect(written.bytes).toEqual(DOC)
	})

	it("keeps unchanged paragraphs' bytes and writes an edited one in full that reads back as edited", async () => {
		const editing = await editRtf(DOC)
		const session = editing.session
		const first = paragraphs(session.state.doc)[0]

		if (first === undefined) throw new Error("no paragraph")

		expect(session.dispatch(session.state.tr.insertText("New ", first.pos + 1))).toBe(true)

		const written = await saveRtf(editing)
		const out = strFromU8(written.bytes, true)

		expect(written.written).toBe(1)
		expect(out).toContain("\\pard\\plain\\f0\\fs24 Second {\\b bold} paragraph\\par\n")
		expect(out).toContain("{\\colortbl;\\red255\\green0\\blue0;}")
		expect(text(await bodyXml(written.bytes))).toBe(
			"New First paragraph here\nSecond bold paragraph\nThird one in red\nFourth paragraph text\nA1 cell\nB1 cell\nFifth café paragraph"
		)
	})

	it("writes the scripted edits so that the file reads back as the edit made it", async () => {
		const editing = await editRtf(SCRIPT)
		const done = await scriptedEdits(editing)

		expect(done).toEqual([
			"type",
			"format",
			"delete",
			"row",
			"list",
			"heading",
			"table",
			"merge",
			"nested",
			"section",
			"page",
			"picture"
		])

		const written = await saveRtf(editing)
		const out = strFromU8(written.bytes, true)

		expect(written.rounds).toBeGreaterThanOrEqual(1)
		// A heading style added to the style sheet; a table with borders, cells merged (one cell over the edge the row
		// below keeps), a table nested in a cell; a section break and the page set up again; a PNG in \shppict with its
		// metafile copy.
		expect(out).toMatch(/\{\\s\d+\\sbasedon0\\snext0[^}]*heading 1;\}/u)
		expect(out).toContain("\\trbrdrt\\brdrs\\brdrw10")
		expect(out).toContain("\\clwWidth6108\\cellx6000{")
		expect(out).not.toContain("\\clmrg")
		expect(out).toContain("{\\*\\nesttableprops\\trowd")
		expect(out).toContain("\\sect")
		expect(out).toContain("\\lndscpsxn")
		expect(out).toContain("\\cols2")
		expect(out).toContain("{\\*\\shppict{\\pict{\\*\\picprop{\\sp{\\sn wzDescription}{\\sv Three dots}}}\\pngblip")
		expect(out).toContain("{\\nonshppict{\\pict\\wmetafile8")
		// Untouched paragraphs keep their bytes.
		expect(out).toContain("\\pard\\plain\\f0\\fs24 Paragraph twelve stays as it was\\par\n")
	})

	it("adds a colour the source has not to the end of its colour table", async () => {
		const editing = await editRtf(DOC)
		const session = editing.session
		const first = paragraphs(session.state.doc)[0]

		if (first === undefined) throw new Error("no paragraph")

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, first.pos + 1, first.pos + 6)))
		expect(session.execute({ type: "color.text", color: { kind: "rgb", value: "00B050" } })).toBe(true)

		const out = strFromU8((await saveRtf(editing)).bytes, true)

		expect(out).toContain("{\\colortbl;\\red255\\green0\\blue0;\\red0\\green176\\blue80;}")
		expect(out).toContain("\\cf2")
	})

	it("proves every paragraph of a plain document writable", async () => {
		const editing = await editRtf(DOC)
		const proof = await proveRtf(editing.source)

		expect(proof.locked).toEqual([])
		expect(proof.written).toBeGreaterThan(0)
	})

	it("writes the track-changes switch as \\revisions after the header, and removes every one when it is turned off", async () => {
		const editing = await editRtf(DOC)

		editing.opened.context.review.tracking = true

		const on = (await saveRtf(editing)).bytes
		const onText = strFromU8(on, true)

		expect(onText).toContain("{\\colortbl;\\red255\\green0\\blue0;}\\revisions")
		expect(onText.replace("\\revisions", "")).toBe(strFromU8(DOC, true))

		const reread = await editRtf(on)

		expect(reread.opened.context.review.tracking).toBe(true)
		reread.opened.context.review.tracking = false
		expect((await saveRtf(reread)).bytes).toEqual(DOC)
	})

	it("writes tracked typing and a tracked paragraph split as revision marks, the new mark's at its \\par", async () => {
		const editing = await editRtf(DOC)
		const session = editing.session
		const first = paragraphs(session.state.doc)[0]

		if (first === undefined) throw new Error("no paragraph")

		editing.opened.context.review.tracking = true
		expect(session.dispatch(session.state.tr.insertText("New ", first.pos + 1))).toBe(true)
		expect(session.dispatch(session.state.tr.split(first.pos + 5))).toBe(true)

		const written = await saveRtf(editing)
		const out = strFromU8(written.bytes, true)
		const xml = await bodyXml(written.bytes)

		expect(out).toContain("\\revisions")
		expect(out).toMatch(/\\revised\\revauth\d+\\revdttm\d+ New\b/u)
		expect(xml).toMatch(/<w:pPr><w:rPr><w:ins w:id="\d+" w:author="Filen" w:date="[^"]+"\/>/u)
		expect(text(xml).split("\n").slice(0, 3)).toEqual(["New ", "First paragraph here", "Second bold paragraph"])
	})

	it("writes tracked formatting, paragraph, cell and page changes as old properties that read back as the changes", async () => {
		const editing = await editRtf(DOC)
		const session = editing.session
		const body = paragraphs(session.state.doc)
		const second = body[1]
		const third = body[2]
		const cell = paragraphs(session.state.doc, true)[0]

		if (second === undefined || third === undefined || cell === undefined) throw new Error("no paragraphs")

		editing.opened.context.review.tracking = true
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, second.pos + 1, second.pos + 7)))
		expect(session.execute({ type: "color.text", color: { kind: "rgb", value: "00B050" } })).toBe(true)
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, third.pos + 1)))
		expect(session.execute({ type: "align.center" })).toBe(true)
		expect(await editing.editor.execute({ type: "page.setup", orientation: "landscape" })).toBe(true)
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, cell.pos + 1)))
		expect(session.execute({ type: "table.cell.vAlign", align: "center" })).toBe(true)

		const written = await saveRtf(editing)
		const out = strFromU8(written.bytes, true)
		const xml = await bodyXml(written.bytes)

		expect(out).toMatch(/\{\\\*\\oldcprops\\plain[^}]*\\crauth\d+\\crdate\d+\}/u)
		expect(out).toMatch(/\{\\\*\\oldpprops\\pard[^}]*\\prauth\d+\\prdate\d+\}/u)
		expect(out).toMatch(/\{\\\*\\oldtprops\\trowd[^}]*\\trauth\d+\\trdate\d+\}/u)
		expect(out).toMatch(/\{\\\*\\oldsprops\\sectd[^}]*\\srauth\d+\\srdate\d+\}/u)
		expect(xml).toMatch(/<w:rPrChange w:id="\d+" w:author="Filen" w:date="[^"]+"><w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr><\/w:rPrChange>/u)
		expect(xml).toMatch(/<w:jc w:val="center"\/>(?:<w:rPr>.*?<\/w:rPr>)?<w:pPrChange w:id="\d+" w:author="Filen"/u)
		expect(xml).toMatch(/<w:vAlign w:val="center"\/><w:tcPrChange w:id="\d+" w:author="Filen" w:date="[^"]+"><w:tcPr>/u)
		expect(xml).toMatch(/<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>.*<w:sectPrChange w:id="\d+" w:author="Filen"/u)
		// Untouched paragraphs keep their bytes.
		expect(out).toContain("\\pard\\plain\\f0\\fs24 Fourth paragraph text\\par\n")
	})

	it("writes a tracked change of a table's grid as each row's old definition", async () => {
		const uneven = latin1(
			"{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0 Times New Roman;}}" +
				"\\trowd\\trgaph108\\trleft-108\\cellx2000\\cellx6000\\pard\\plain\\intbl A1 cell\\cell B1 cell\\cell\\row\n" +
				"\\trowd\\trgaph108\\trleft-108\\cellx2000\\cellx6000\\pard\\plain\\intbl A2 cell\\cell B2 cell\\cell\\row\n" +
				"\\pard\\plain\\f0\\fs24 After the table\\par\n}"
		)
		const editing = await editRtf(uneven)
		const session = editing.session
		const cell = paragraphs(session.state.doc, true)[0]

		if (cell === undefined) throw new Error("no cell")

		editing.opened.context.review.tracking = true
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, cell.pos + 1)))
		expect(session.execute({ type: "table.distributeColumns" })).toBe(true)

		const written = await saveRtf(editing)
		const out = strFromU8(written.bytes, true)
		const xml = await bodyXml(written.bytes)

		expect(
			out.match(/\{\\\*\\oldtprops\\trowd[^}]*\\cellx2000\\clftsWidth3\\clwWidth4000\\cellx6000[^}]*\\trauth\d+\\trdate\d+\}/gu)
				?.length
		).toBe(2)
		expect(xml).toMatch(
			/<w:tblGridChange w:id="\d+"><w:tblGrid><w:gridCol w:w="2108"\/><w:gridCol w:w="4000"\/><\/w:tblGrid><\/w:tblGridChange>/u
		)
		expect(out).toContain("\\trleft-108\\clftsWidth3\\clwWidth3054\\cellx2946\\clftsWidth3\\clwWidth3054\\cellx6000")
		expect(text(xml)).toBe("A1 cell\nB1 cell\nA2 cell\nB2 cell\nAfter the table")
	})

	it("reads old properties as the change they revert, with the author and date around them", async () => {
		const xml = await bodyXml(
			latin1(
				"{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times New Roman;}}{\\*\\revtbl {Unknown;}{Ann;}}" +
					"\\pard\\plain\\qc{\\*\\oldpprops\\pard\\qr\\prauth1\\prdate131474048}\\f0 {\\b\\i\\crauth1\\crdate131474048{\\*\\oldcprops \\b0\\i0\\crauth1\\crdate131474048}is a} test\\par}"
			)
		)

		expect(xml).toContain(
			'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:i/><w:rPrChange w:id="1" w:author="Ann" w:date="2025-06-04T10:00:00Z"><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b w:val="0"/><w:i w:val="0"/></w:rPr></w:rPrChange></w:rPr>'
		)
		expect(xml).toMatch(
			/<w:jc w:val="center"\/>.*<w:pPrChange w:id="\d+" w:author="Ann" w:date="2025-06-04T10:00:00Z"><w:pPr><w:jc w:val="right"\/><\/w:pPr><\/w:pPrChange>/u
		)
	})

	it("refuses a tracked row insertion: RTF 1.9.1 tracks rows only by \\trauthN and \\trdateN", async () => {
		const editing = await editRtf(DOC)
		const session = editing.session
		const cell = paragraphs(session.state.doc, true)[0]

		if (cell === undefined) throw new Error("no cell")

		editing.opened.context.review.tracking = true
		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, cell.pos + 1)))
		expect(session.execute({ type: "table.row.insertBelow" })).toBe(true)

		await expect(saveRtf(editing)).rejects.toThrow(/RTF 1\.9\.1/u)
	})

	it("writes a cell's tracked insertion and deletion as \\clins and \\cldel with their author and date", async () => {
		const tracked = latin1(
			"{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0 Times New Roman;}}{\\*\\revtbl {Unknown;}{Bo;}}" +
				"\\trowd\\trgaph108\\trleft-108\\clins\\clinsauth1\\clinsdttm131474048\\cellx3000\\cldel\\cldelauth1\\cldeldttm131474048\\cellx6000" +
				"\\pard\\plain\\intbl A1 cell\\cell B1 cell\\cell\\row\n" +
				"\\pard\\plain\\f0\\fs24 After the table\\par\n}"
		)
		const editing = await editRtf(tracked)
		const session = editing.session
		const cell = paragraphs(session.state.doc, true)[0]

		if (cell === undefined) throw new Error("no cell")

		expect(session.dispatch(session.state.tr.insertText("Edited ", cell.pos + 1))).toBe(true)

		const written = await saveRtf(editing)
		const out = strFromU8(written.bytes, true)

		expect(written.written).toBeGreaterThan(0)
		expect(out).toMatch(/\\clins\\clinsauth1\\clinsdttm\d+/u)
		expect(out).toMatch(/\\cldel\\cldelauth1\\cldeldttm\d+/u)

		const xml = await bodyXml(written.bytes)

		expect(xml).toMatch(/<w:cellIns w:id="\d+" w:author="Bo" w:date="[^"]+"\/>/u)
		expect(xml).toMatch(/<w:cellDel w:id="\d+" w:author="Bo" w:date="[^"]+"\/>/u)
		expect(text(xml)).toBe("Edited A1 cell\nB1 cell\nAfter the table")
	})

	it("refuses an edit part it has no RTF for, writing nothing", async () => {
		const editing = await editRtf(DOC)

		await expect(writeRtf(editing.source, { parts: new Map([["word/settings.xml", latin1("<x/>")]]) })).rejects.toThrow(/settings/u)
	})

	it("keeps what an edited paragraph sets to itself: a following paragraph without \\pard reads as it did", async () => {
		const shared = latin1(
			"{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times New Roman;}}\\pard\\plain\\qc\\f0 {Centred one.\n\\par Centred two.\n\\par Centred three.}\\par\n}"
		)
		const editing = await editRtf(shared)
		const session = editing.session
		const second = paragraphs(session.state.doc)[1]

		if (second === undefined) throw new Error("no paragraph")

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, second.pos + 1)))
		expect(session.execute({ type: "align.end" })).toBe(true)

		const out = strFromU8((await saveRtf(editing)).bytes, true)

		// The edited paragraph in a group of its own; the third keeps its bytes and its centring.
		expect(out).toMatch(/\{\\pard\\plain\\qr[^}]*Centred two\.\}[^{]*\\par\}/u)
		expect(out).toContain("Centred three.}")
		expect((await bodyXml(new TextEncoder().encode(out))).match(/<w:jc w:val="center"\/>/gu)?.length).toBe(2)
	})

	it("writes a link the edit made as a HYPERLINK field", async () => {
		const editing = await editRtf(DOC)
		const session = editing.session
		const first = paragraphs(session.state.doc)[0]

		if (first === undefined) throw new Error("no paragraph")

		session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, first.pos + 1, first.pos + 6)))
		expect(session.execute({ type: "link.set", target: "https://filen.io/" })).toBe(true)

		const out = strFromU8((await saveRtf(editing)).bytes, true)

		expect(out).toContain('{\\field\\plain{\\*\\fldinst{\\plain HYPERLINK "https://filen.io/"}}')
		expect(text(await bodyXml(new TextEncoder().encode(out)))).toContain("First paragraph here")
	})

	it("locks a paragraph holding source the reader keeps only as bytes, and keeps the rest editable", async () => {
		const shape = latin1(
			"{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times New Roman;}}\\pard\\plain Plain paragraph\\par\n" +
				"\\pard\\plain Before {\\*\\unknowndest kept bytes} after\\par\n}"
		)
		const editing = await editRtf(shape)

		expect(editing.locked).toBe(1)

		const proof = await proveRtf(editing.source)

		expect(proof.locked.map(block => block.reason)).toEqual(["preserved"])
	})
})
