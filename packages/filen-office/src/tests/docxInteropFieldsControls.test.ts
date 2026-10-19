import { strToU8, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { summarizeDocument } from "./fuzz/saves"
import { acceptedReason } from "./interop/accepted"
import type { EntryResult } from "./interop/interop"
import { summarizeFodt } from "./interop/odfSummary"

// The LibreOffice leg's readings of tables of contents, macro buttons, rows inserted as tracked changes and framed
// cell paragraphs, each as LibreOffice 7.5 and 26.8 were seen to read the minimal file named.

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const AUTHOR = 'w:author="Ada" w:date="2026-01-01T00:00:00Z"'

function docx(body: string): Uint8Array {
	return zipSync({
		"[Content_Types].xml": strToU8(
			'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
		),
		"_rels/.rels": strToU8(
			'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
		),
		"word/document.xml": strToU8(
			`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`
		)
	})
}

const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const char = (kind: string): string => `<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`
const instr = (code: string, deleted = false): string =>
	deleted
		? `<w:r><w:delInstrText xml:space="preserve">${code}</w:delInstrText></w:r>`
		: `<w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r>`
const TOC = ' TOC \\o "1-3" \\h '
const block = (content: string): string =>
	`<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`

async function lines(body: string): Promise<{
	final: string[]
	original: string[]
	counts: Readonly<Record<string, number>>
}> {
	const summary = await summarizeDocument(docx(body))

	if (summary === undefined) throw new Error("unreadable")

	const nonEmpty = (text: string): string[] => text.split("\n").filter(line => line.trim() !== "")

	return {
		final: nonEmpty(summary.final),
		original: nonEmpty(summary.original),
		counts: summary.counts
	}
}

function fodt(text: string): Uint8Array {
	return strToU8(
		`<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:text>${text}</office:text></office:body></office:document>`
	)
}

describe("tables of contents as LibreOffice imports them", () => {
	it("leaves out a building block holding one, title and text after the field included", async () => {
		const toc = block(
			p("Contents") +
				`<w:p>${char("begin")}${instr(TOC)}${char("separate")}${r("Entry one")}</w:p>` +
				`<w:p>${char("end")}<w:ins w:id="1" ${AUTHOR}>${r("Inserted after")}</w:ins></w:p>`
		)

		expect((await lines(p("Before") + toc + p("After"))).final).toEqual(["Before", "After"])
	})

	it("reads the entries as body text when the field begins after text and holds none of its own result there", async () => {
		const pageRef = `${char("begin")}${instr(" PAGEREF _Toc1 \\h ")}${char("separate")}${r("7")}${char("end")}`
		const toc = (head: string): string =>
			block(
				`<w:p>${head}${char("begin")}${instr(TOC)}${char("separate")}</w:p><w:p>${r("Entry one")}${pageRef}</w:p>${p("Entry two")}<w:p>${char("end")}</w:p>`
			)

		expect((await lines(toc(r("Table of Contents")))).final).toEqual(["Entry one7", "Entry two"])
		expect((await lines(toc(r(" ")))).final).toEqual(["Entry one7", "Entry two"])
		// The field first in its paragraph, or its result in it: all index.
		expect((await lines(toc(""))).final).toEqual([])
		expect(
			(
				await lines(
					block(
						`<w:p>${r("Typed")}${char("begin")}${instr(TOC)}${char("separate")}${r("Entry one")}</w:p>${p("Entry two")}<w:p>${char("end")}</w:p>`
					)
				)
			).final
		).toEqual([])
		// Outside a building block the field's result is the index.
		expect(
			(
				await lines(
					`<w:p>${r("Table of Contents")}${char("begin")}${instr(TOC)}${char("separate")}</w:p>${p("Entry one")}<w:p>${char("end")}</w:p>`
				)
			).final
		).toEqual(["Table of Contents"])
	})

	it("keeps a block without its field as content", async () => {
		expect((await lines(block(p("Body in a block")))).final).toEqual(["Body in a block"])
	})

	it("knows a deleted field by its deleted instruction and one in a deleted row", async () => {
		const deleted =
			`<w:p><w:del w:id="1" ${AUTHOR}>${char("begin")}${instr(TOC, true)}${char("separate")}<w:r><w:delText>Entry one</w:delText></w:r></w:del></w:p>` +
			`<w:p><w:del w:id="2" ${AUTHOR}>${char("end")}</w:del></w:p>`
		const grid = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
		const row = `<w:tbl>${grid}<w:tr><w:trPr><w:del w:id="3" ${AUTHOR}/></w:trPr><w:tc><w:p>${char("begin")}${instr(" REF Mark \\r \\h ")}${char("separate")}${r("2.2")}${char("end")}</w:p></w:tc></w:tr></w:tbl>`
		const summary = await lines(p("Before") + deleted + row + p("After"))

		expect(summary.original).toEqual(["Before", "After"])
		expect(summary.final).toEqual(["Before", "After"])
	})

	it("mutes a deletion's text that comes back inside an index (LibreOffice 26.8 keeps it as a point change there)", () => {
		const summary = summarizeFodt(
			fodt(
				'<text:tracked-changes><text:changed-region text:id="d1"><text:deletion><text:p>.2 Entry</text:p></text:deletion></text:changed-region></text:tracked-changes>' +
					'<text:p>Before</text:p><text:table-of-content><text:index-body><text:p><text:change text:change-id="d1"/>Entry two</text:p></text:index-body></text:table-of-content><text:p>After</text:p>'
			)
		)

		expect(summary.original.split("\n").filter(line => line !== "")).toEqual(["Before", "After"])
	})
})

describe("other fields and changes as LibreOffice holds them", () => {
	it("reads no text of a macro button, ours or LibreOffice's display text", async () => {
		const button = `<w:p>${char("begin")}${instr(' MACROBUTTON  synchronize "Sync contents" ')}${char("separate")}${r("*")}${char("end")}</w:p>`

		expect((await lines(button + p("Body"))).final).toEqual(["Body"])
		expect(
			summarizeFodt(fodt('<text:p><text:execute-macro>"Sync contents"</text:execute-macro></text:p><text:p>Body</text:p>')).final
		).toBe("\nBody")
	})

	it("keeps rows inserted as tracked changes with every change rejected", async () => {
		const grid = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
		const inserted = `<w:pPr><w:rPr><w:ins w:id="1" ${AUTHOR}/></w:rPr></w:pPr>`
		const table = `<w:tbl>${grid}<w:tr><w:trPr><w:ins w:id="2" ${AUTHOR}/></w:trPr><w:tc><w:p>${inserted}</w:p></w:tc></w:tr></w:tbl>`

		// LibreOffice 7.5 and 26.8: "Mak" | "e the rest." (the row stays, so the paragraphs cannot join across it).
		expect((await lines(`<w:p>${inserted}${r("Mak")}</w:p>${table}${p("e the rest.")}`)).original).toEqual(["Mak", "e the rest."])
	})

	it("counts framed paragraphs in table cells, and accepts 7.5's lost changes only where one is", async () => {
		const grid = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
		const framed = `<w:tbl>${grid}<w:tr><w:tc><w:p><w:pPr><w:framePr w:wrap="around"/></w:pPr>${r("Framed")}</w:p></w:tc></w:tr></w:tbl>`

		expect((await lines(framed + p("After"))).counts["framedCellParagraphs"]).toBe(1)
		expect((await lines(`<w:p><w:pPr><w:framePr w:wrap="around"/></w:pPr>${r("Framed")}</w:p>`)).counts["framedCellParagraphs"]).toBe(0)

		const entry = (count: number): EntryResult =>
			({
				source: "docx/x.docx",
				witness: {
					before: { framedCellParagraphs: count },
					after: { framedCellParagraphs: count }
				}
			}) as unknown as EntryResult

		expect(acceptedReason(entry(1), 'L2 final: [["If",1]]', "lo-7.5.0.3")).toContain("framed paragraph")
		expect(acceptedReason(entry(1), 'L2 final: [["If",1]]', "lo-26.8.1.1")).toBeUndefined()
		expect(acceptedReason(entry(0), 'L2 final: [["If",1]]', "lo-7.5.0.3")).toBeUndefined()
	})
})
