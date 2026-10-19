import { strToU8, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { summarizeDocument } from "./fuzz/saves"
import { acceptedCount, acceptedReason } from "./interop/accepted"
import type { EntryResult } from "./interop/interop"
import { summarizeFodt } from "./interop/odfSummary"

// The LibreOffice leg's readings of both sides (fuzz/saves.ts, interop/odfSummary.ts) and its accepted
// differences (interop/accepted.ts), each on the construct LibreOffice was seen to read its own way.

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

const p = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const grid = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
const deletedMark = `<w:pPr><w:rPr><w:del w:id="1" ${AUTHOR}/></w:rPr></w:pPr>`

async function read(body: string): Promise<{ final: string[]; original: string[]; counts: Readonly<Record<string, number>> }> {
	const summary = await summarizeDocument(docx(body))

	if (summary === undefined) throw new Error("unreadable")

	return { final: summary.final.split("\n"), original: summary.original.split("\n"), counts: summary.counts }
}

function fodt(text: string): string {
	return `<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:text>${text}</office:text></office:body></office:document>`
}

describe("our reading as an ODF body holds it", () => {
	it("leaves out a cell continuing a vertical merge, as LibreOffice's covered cell", async () => {
		const merged = `<w:tbl>${grid}<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${p("A")}</w:tc><w:tc>${p("X")}</w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>${p("B")}</w:tc><w:tc>${p("Y")}</w:tc></w:tr></w:tbl>`
		const summary = await read(p("Before") + merged + p("After"))

		// LibreOffice 7.5 and 26.8: 5 paragraphs, 3 cells (the continuation's text dropped).
		expect(summary.counts["bodyParagraphs"]).toBe(5)
		expect(summary.counts["cells"]).toBe(3)
	})

	it("ends a paragraph whose deleted mark meets a table or a cell end, the rows being kept", async () => {
		const table = `<w:tbl>${grid}<w:tr><w:tc>${p("X")}</w:tc><w:tc><w:p>${deletedMark}<w:r><w:t>Y</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
		const summary = await read(`<w:p>${deletedMark}<w:r><w:t>A</w:t></w:r></w:p>${table}${p("B")}`)

		// LibreOffice: A | X | Y | B, never "AX" or "YB".
		expect(summary.final).toEqual(["A", "X", "Y", "B"])
		expect(summary.counts["bodyParagraphs"]).toBe(4)
	})

	it("keeps rows deleted as tracked changes in the text with changes accepted, so a paragraph does not run across them", async () => {
		const row = `<w:tr><w:trPr><w:del w:id="5" ${AUTHOR}/></w:trPr><w:tc><w:p>${deletedMark}<w:del w:id="6" ${AUTHOR}><w:r><w:delText>X</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>`
		const summary = await read(`<w:p>${deletedMark}<w:r><w:t>A</w:t></w:r></w:p><w:tbl>${grid}${row}</w:tbl>${p("B")}`)

		expect(summary.final).toEqual(["A", "", "", "B"])
	})

	it("counts a page break before the document's first text as a paragraph of its own, and only that one", async () => {
		const first = await read(`<w:p><w:r><w:br w:type="page"/><w:t>A</w:t></w:r></w:p>${p("B")}`)
		const second = await read(`${p("A")}<w:p><w:r><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>`)

		expect(first.counts["bodyParagraphs"]).toBe(3)
		expect(second.counts["bodyParagraphs"]).toBe(2)
	})

	it("counts a break before the text of any section's first paragraph as a paragraph of its own", async () => {
		const ended = await read(
			'<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>'
		)
		const inSection = await read(`${p("A")}<w:p><w:r><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>${p("C")}`)

		expect(ended.counts["bodyParagraphs"]).toBe(3)
		expect(inSection.counts["bodyParagraphs"]).toBe(3)
	})

	it("counts the breaks it counts as paragraph starts that a floating drawing follows", async () => {
		const anchor =
			'<w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="9525" cy="9525"/><wp:wrapNone/><wp:docPr id="1" name="d"/></wp:anchor></w:drawing>'
		const between = await read(`<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r><w:r>${anchor}<w:t>B</w:t></w:r></w:p>`)
		const atStart = await read(`${p("A")}<w:p><w:r><w:br w:type="page"/></w:r><w:r>${anchor}<w:t>B</w:t></w:r></w:p>`)
		const textFirst = await read(`<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>C</w:t></w:r><w:r>${anchor}<w:t>B</w:t></w:r></w:p>`)

		expect(between.counts["breakAnchors"]).toBe(1)
		expect(between.counts["bodyParagraphs"]).toBe(2)
		// Not a paragraph start: not counted either.
		expect(atStart.counts["breakAnchors"]).toBe(0)
		expect(textFirst.counts["breakAnchors"]).toBe(0)
	})

	it("counts the paragraph layouts LibreOffice makes its own way", async () => {
		const table = (inner = "<w:p/>", floating = false): string =>
			`<w:tbl><w:tblPr>${floating ? '<w:tblpPr w:tblpY="1"/>' : ""}<w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`
		const note = '<w:r><w:footnoteReference w:id="1"/></w:r>'
		const startTable = await read(table() + `<w:p><w:r><w:t>A</w:t></w:r>${note}</w:p>`)
		const laterTable = await read(`<w:p>${note}</w:p>` + table())
		const plainStart = await read(table() + p("A"))
		const floatingThenTable = await read(table("<w:p/>", true) + table() + p("A"))
		const floatingThenParagraph = await read(table("<w:p/>", true) + p("A") + table())
		const custom = '<w:r><w:endnoteReference w:customMarkFollows="1" w:id="2"/><w:t>a)</w:t></w:r>'
		const customThenBreak = await read(
			`<w:p>${custom}</w:p><w:p><w:r><w:br w:type="page"/><w:t>B</w:t></w:r>${note}</w:p><w:p><w:r><w:br w:type="page"/><w:t>C</w:t></w:r></w:p>`
		)
		const begin =
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
		const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
		const sdt = (inner: string): string => `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
		const tocLeading = await read(sdt(`<w:p><w:r><w:t xml:space="preserve"> </w:t></w:r>${begin}</w:p><w:p>${end}</w:p>`) + p("A"))
		const tocBare = await read(sdt(`<w:p>${begin}</w:p><w:p>${end}</w:p>`) + p("A"))
		const tocOutside = await read(`<w:p><w:r><w:t xml:space="preserve"> </w:t></w:r>${begin}</w:p><w:p>${end}</w:p>` + p("A"))
		const tocTail = await read(`<w:p>${begin}<w:r><w:t>E</w:t></w:r></w:p><w:p>${deletedMark}${end}</w:p>` + p("A"))
		const row = (mark: string): string =>
			`<w:tbl><w:tblPr><w:tblpPr w:tblpY="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr>${mark}</w:trPr><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>${p("A")}`
		const deletedRow = await read(row(`<w:del w:id="1" ${AUTHOR}/>`))
		const insertedRow = await read(row(`<w:ins w:id="1" ${AUTHOR}/>`))

		expect(startTable.counts["startTableAnnotated"]).toBe(1)
		expect(laterTable.counts["startTableAnnotated"]).toBe(0)
		expect(plainStart.counts["startTableAnnotated"]).toBe(0)
		expect(floatingThenTable.counts["floatingTableBeforeTable"]).toBe(1)
		expect(floatingThenParagraph.counts["floatingTableBeforeTable"]).toBe(0)
		expect(customThenBreak.counts["customMarkBreaks"]).toBe(1)
		expect(tocLeading.counts["tocSdtLeadingText"]).toBe(1)
		expect(tocBare.counts["tocSdtLeadingText"]).toBe(0)
		expect(tocOutside.counts["tocSdtLeadingText"]).toBe(0)
		expect(tocTail.counts["indexTailDeletedMarks"]).toBe(1)
		expect(tocBare.counts["indexTailDeletedMarks"]).toBe(0)
		expect(deletedRow.counts["floatingTrackedRows"]).toBe(1)
		expect(insertedRow.counts["floatingTrackedRows"]).toBe(1)
		expect(plainStart.counts["floatingTrackedRows"]).toBe(0)
	})

	it("counts the constructs LibreOffice reads its own way", async () => {
		const section = (type: string, content = ""): string =>
			`<w:p><w:pPr><w:sectPr><w:type w:val="${type}"/></w:sectPr></w:pPr>${content}</w:p>`
		const empty = await read(p("A") + section("oddPage") + p("B"))
		const deleted = await read(`${p("A")}<w:p><w:pPr><w:rPr><w:del w:id="1" ${AUTHOR}/></w:rPr><w:sectPr/></w:pPr></w:p>${p("B")}`)
		const last = await read(p("A") + section("nextPage", "<w:r><w:t>B</w:t></w:r>"))
		const framed = await read(
			`${p("A")}<w:p><w:pPr><w:framePr w:w="5760" w:h="245" w:hAnchor="page" w:vAnchor="page"/></w:pPr><w:r><w:t>F</w:t></w:r></w:p>`
		)
		const nested = await read(
			`<w:p><w:ins w:id="1" ${AUTHOR}><w:del w:id="2" ${AUTHOR}><w:r><w:delText>gone</w:delText></w:r></w:del></w:ins></w:p>`
		)
		const formatted = await read(
			`<w:p><w:ins w:id="1" ${AUTHOR}><w:r><w:rPr><w:b/><w:rPrChange w:id="2" ${AUTHOR}><w:rPr/></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:ins></w:p>`
		)
		const moved = await read(
			`<w:p><w:pPr><w:rPr><w:moveFrom w:id="1" ${AUTHOR}/></w:rPr></w:pPr><w:moveFrom w:id="2" ${AUTHOR}><w:r><w:t>M</w:t></w:r></w:moveFrom><w:ins w:id="3" ${AUTHOR}><w:r><w:t>N</w:t></w:r></w:ins></w:p>${p("K")}`
		)

		expect(empty.counts["emptySectionParagraphs"]).toBe(1)
		expect(deleted.counts["deletedSectionMarks"]).toBe(1)
		expect(deleted.counts["emptySectionParagraphs"]).toBe(0)
		expect(last.counts["emptyLastSection"]).toBe(1)
		expect(empty.counts["emptyLastSection"]).toBe(0)
		expect(framed.counts["framedParagraphs"]).toBe(1)
		expect(nested.counts["nestedDeletions"]).toBe(1)
		expect(formatted.counts["formattedChangeRuns"]).toBe(1)
		expect(moved.counts["movedMarksInserted"]).toBe(1)
	})
})

describe("LibreOffice's reading of its flat ODF", () => {
	it("reads a text frame anchored to a paragraph (a framed paragraph or floating table) as body text after it", () => {
		const summary = summarizeFodt(
			strToU8(
				fodt(
					'<text:p>A<draw:frame text:anchor-type="paragraph"><draw:text-box><table:table><table:table-row><table:table-cell><text:p>X</text:p></table:table-cell></table:table-row></table:table></draw:text-box></draw:frame></text:p><text:p>B<draw:frame text:anchor-type="char"><draw:text-box><text:p>box</text:p></draw:text-box></draw:frame></text:p>'
				)
			)
		)

		expect(summary.final.split("\n")).toEqual(["A", "X", "B"])
		expect(summary.counts["bodyParagraphs"]).toBe(3)
	})

	it("counts an illustration index's paragraphs but not its text", () => {
		const summary = summarizeFodt(
			strToU8(
				fodt(
					"<text:p>A</text:p><text:illustration-index><text:index-body><text:p>Figure 1\t3</text:p></text:index-body></text:illustration-index>"
				)
			)
		)

		expect(summary.final.split("\n")).toEqual(["A", ""])
	})
})

describe("accepted differences", () => {
	const entry = (before: Record<string, number>, after: Record<string, number>): EntryResult => ({
		path: "x/y.docx",
		source: "docx/y.docx",
		family: "x",
		l1: [],
		l1Original: [],
		l2: {},
		l3: {},
		g7: {},
		l4: {},
		witness: { before, after }
	})

	it("accepts a count residual only within what the witnesses of both files allow", () => {
		// Two empty section paragraphs added: LibreOffice may drop both, one or none, never add one.
		const before = { emptySectionParagraphs: 0 }
		const after = { emptySectionParagraphs: 2 }

		expect(acceptedCount("bodyParagraphs", 3, 5, before, after, "lo-26.8.1.1")).toBeDefined()
		expect(acceptedCount("bodyParagraphs", 4, 5, before, after, "lo-26.8.1.1")).toBeDefined()
		expect(acceptedCount("bodyParagraphs", 2, 5, before, after, "lo-26.8.1.1")).toBeUndefined()
		expect(acceptedCount("bodyParagraphs", 6, 5, before, after, "lo-26.8.1.1")).toBeUndefined()
		expect(acceptedCount("bodyParagraphs", 3, 5, {}, {}, "lo-26.8.1.1")).toBeUndefined()
		expect(acceptedCount("tables", 3, 5, before, after, "lo-26.8.1.1")).toBeUndefined()
	})

	it("accepts a text residual only where the construct is in one of the files, and only for the offices named", () => {
		const failure = 'L2 original: [["Durlening",1],["During",-1]]'

		expect(acceptedReason(entry({ formattedChangeRuns: 0 }, { formattedChangeRuns: 1 }), failure, "lo-26.8.1.1")).toBeDefined()
		expect(acceptedReason(entry({ formattedChangeRuns: 0 }, { formattedChangeRuns: 1 }), failure, "lo-7.5.0.3")).toBeUndefined()
		expect(acceptedReason(entry({}, {}), failure, "lo-26.8.1.1")).toBeUndefined()
	})
})
