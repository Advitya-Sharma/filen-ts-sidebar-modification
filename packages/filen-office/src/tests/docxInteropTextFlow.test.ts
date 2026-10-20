import { strToU8, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { summarizeDocument } from "./fuzz/saves"
import { summarizeFodt } from "./interop/odfSummary"
import { ZIP_ENTRY_TIME } from "./zipTime"

// The LibreOffice leg's readings of drop caps, deletions inside an index, field codes
// deleted as a tracked change, and the witnesses of LibreOffice's own readings (breaks before floating drawings,
// tables of contents in another one's result).

const W =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const AUTHOR = 'w:author="Ada" w:date="2026-01-01T00:00:00Z"'

function docx(body: string): Uint8Array {
	return zipSync(
		{
			"[Content_Types].xml": strToU8(
				'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
			),
			"_rels/.rels": strToU8(
				'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
			),
			"word/document.xml": strToU8(
				`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`
			)
		},
		{ mtime: ZIP_ENTRY_TIME }
	)
}

const p = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const dropCap = (letter: string): string =>
	`<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr><w:r><w:t>${letter}</w:t></w:r></w:p>`
const table =
	'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

function field(code: string, result: string): string {
	return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r>`
}

const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

async function read(body: string): Promise<{
	final: string[]
	original: string[]
	counts: Readonly<Record<string, number>>
}> {
	const summary = await summarizeDocument(docx(body))

	if (summary === undefined) throw new Error("unreadable")

	return {
		final: summary.final.split("\n"),
		original: summary.original.split("\n"),
		counts: summary.counts
	}
}

function fodt(text: string): Uint8Array {
	return strToU8(
		`<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:body><office:text>${text}</office:text></office:body></office:document>`
	)
}

describe("our reading as an ODF body holds it", () => {
	it("starts the next paragraph with a drop cap's text, as ODF's style:drop-cap", async () => {
		// LibreOffice 7.5 and 26.8: "Lorem ipsum." | "Next."; before an empty paragraph "L" | "orem"; before a table
		// 26.8 starts the first cell with it.
		expect((await read(dropCap("L") + p("orem ipsum.") + p("Next."))).final).toEqual(["Lorem ipsum.", "Next."])
		expect((await read(dropCap("L") + "<w:p/>" + p("orem"))).final).toEqual(["L", "orem"])

		const before = await read(dropCap("L") + table + p("orem"))

		expect(before.final).toEqual(["LCell", "orem"])
		expect(before.counts["bodyParagraphs"]).toBe(2)
		// Not a framed paragraph LibreOffice may anchor (interop/accepted.ts framedParagraphs).
		expect(before.counts["framedParagraphs"]).toBe(0)
		// A drop cap ending the story stays a paragraph of its own.
		expect((await read(p("A") + dropCap("Z"))).final).toEqual(["A", "Z"])
	})

	it("leaves out the result of a field whose code was deleted only where the code still counts", async () => {
		const deleted = `<w:p><w:del w:id="1" ${AUTHOR}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText xml:space="preserve"> TOC \\o "1-3" </w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText>Heading 1</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del></w:p>`
		const summary = await read(p("A") + deleted + p("B"))

		// Rejected, the table of contents is back, and LibreOffice's index text is left out on its side.
		expect(summary.original).toEqual(["A", "", "B"])
	})

	it("counts the instances of LibreOffice's own readings", async () => {
		const anchor =
			'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
		const pageBreak = '<w:r><w:br w:type="page"/></w:r>'

		// A break before plain text is not one.
		expect((await read(`<w:p><w:r><w:t>A</w:t></w:r>${pageBreak}${anchor}<w:r><w:t>B</w:t></w:r></w:p>`)).counts["breakAnchors"]).toBe(
			1
		)
		expect((await read(`<w:p><w:r><w:t>A</w:t></w:r>${pageBreak}<w:r><w:t>B</w:t></w:r></w:p>`)).counts["breakAnchors"]).toBe(0)

		const nested = `<w:p>${field('TOC \\o "1-4"', "Outer")}</w:p><w:p>${field('TOC \\o "1-2"', "Inner")}</w:p><w:p>${END}</w:p><w:p><w:r><w:t>Rest</w:t></w:r></w:p><w:p>${END}</w:p>`
		const plain = `<w:p>${field('TOC \\o "1-4"', "Outer")}</w:p><w:p>${END}</w:p>${p("After")}<w:p>${field('TOC \\o "1-2"', "Other")}</w:p><w:p>${END}</w:p>`

		expect((await read(nested)).counts["nestedTocs"]).toBe(1)
		expect((await read(plain)).counts["nestedTocs"]).toBe(0)
	})
})

describe("LibreOffice's reading of indexes", () => {
	it("leaves out a deletion inside an index's entries, as the entries' text", () => {
		const summary = summarizeFodt(
			fodt(
				'<text:tracked-changes><text:changed-region text:id="c1"><text:deletion><office:change-info><dc:creator>Ada</dc:creator></office:change-info><text:p>ack</text:p></text:deletion></text:changed-region></text:tracked-changes><text:table-of-content><text:index-body><text:p>P<text:change text:change-id="c1"/>age</text:p></text:index-body></text:table-of-content><text:p>Body</text:p>'
			)
		)

		expect(summary.original.split("\n")).toEqual(["", "Body"])
	})
})
