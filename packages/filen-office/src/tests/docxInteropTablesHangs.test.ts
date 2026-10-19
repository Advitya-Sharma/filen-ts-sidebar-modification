import { strToU8, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { summarizeDocument } from "./fuzz/saves"
import { acceptedCount, acceptedReason } from "./interop/accepted"
import type { EntryResult } from "./interop/interop"
import { summarizeFodt } from "./interop/odfSummary"
import { buildDocx } from "./docxSynthetic"

// The LibreOffice leg's tables are counted in the body's flow on both sides (not in notes or a chart's data on
// LibreOffice's), and tables in DrawingML text boxes, which LibreOffice cannot hold, are witnessed; so are the
// constructs LibreOffice hangs or runs away on, and the failures that are no difference an edit made.

const W =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

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
const table =
	'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
// A DrawingML text box and a VML one holding `content`.
const drawingBox = (content: string): string =>
	`<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2000000" cy="1000000"/><wp:wrapNone/><wp:docPr id="2" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
const vmlBox = (content: string): string =>
	`<w:p><w:r><w:pict><v:shape id="Box" style="position:absolute;width:150pt;height:80pt"><v:textbox><w:txbxContent>${content}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`

function fodt(text: string): Uint8Array {
	return strToU8(
		`<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:text>${text}</office:text></office:body></office:document>`
	)
}

const odfTable = "<table:table><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table>"

describe("LibreOffice's tables", () => {
	it("counts the body's tables, not a note's or a chart's data", () => {
		const summary = summarizeFodt(
			fodt(
				`${odfTable}<text:p>A<text:note text:note-class="footnote"><text:note-body>${odfTable}</text:note-body></text:note></text:p>` +
					`<text:p><draw:frame><draw:object><office:document><office:body><office:chart><chart:chart>${odfTable}</chart:chart></office:chart></office:body></office:document></draw:object></draw:frame></text:p>`
			)
		)

		expect(summary.counts["tables"]).toBe(1)
		expect(summary.counts["rows"]).toBe(1)
		expect(summary.counts["cells"]).toBe(1)
	})
})

describe("our tables", () => {
	it("witnesses tables and rows in DrawingML text boxes, not in VML ones", async () => {
		const summary = await summarizeDocument(
			docx(p("Before") + drawingBox(table + p("In")) + vmlBox(table + p("In")) + table + p("After"))
		)

		if (summary === undefined) throw new Error("unreadable")

		expect(summary.counts["tables"]).toBe(3)
		expect(summary.counts["drawingTextboxTables"]).toBe(1)
		expect(summary.counts["rows"]).toBe(6)
		expect(summary.counts["drawingTextboxRows"]).toBe(2)
	})

	it("accepts a table LibreOffice lost from a DrawingML text box, and no more", () => {
		const before = { drawingTextboxTables: 0, drawingTextboxRows: 0 }
		const after = { drawingTextboxTables: 1, drawingTextboxRows: 2 }

		// One table added in a text box: LibreOffice counts none.
		expect(acceptedCount("tables", 0, 1, before, after, "lo-26.8.1.1")).toBeDefined()
		expect(acceptedCount("rows", 0, 2, before, after, "lo-7.5.0.3")).toBeDefined()
		// More than the boxes hold is not LibreOffice's.
		expect(acceptedCount("tables", 0, 2, before, after, "lo-26.8.1.1")).toBeUndefined()
		expect(acceptedCount("tables", 0, 1, before, before, "lo-26.8.1.1")).toBeUndefined()
	})
})

describe("notes that start with a table", () => {
	const note = (id: number, content: string): string => `<w:footnote w:id="${String(id)}">${content}</w:footnote>`
	const separators =
		'<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
	const body = '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>'

	it("witnesses a note whose first block is a table, not one with a paragraph before it", async () => {
		const first = await summarizeDocument(
			buildDocx({ body, footnotes: separators + note(1, table + p("After")) + note(2, p("Plain")) })
		)
		const later = await summarizeDocument(
			buildDocx({ body, footnotes: separators + note(1, p("Before") + table) + note(2, p("Plain")) })
		)

		expect(first?.counts["notesStartingWithTable"]).toBe(1)
		expect(later?.counts["notesStartingWithTable"]).toBe(0)
	})
})

describe("LibreOffice's runaways and its own round trip", () => {
	const entry = (l1: string[], l1Original: string[], witness: Record<string, number> = {}): EntryResult => ({
		path: "base/x.docx",
		source: "docx/x.docx",
		family: "base",
		l1,
		l1Original,
		l2: {},
		l3: {},
		g7: {},
		l4: {},
		witness: { before: {}, after: witness }
	})

	it("witnesses a body that ends with a table, not one with a paragraph after it", async () => {
		const ends = await summarizeDocument(docx(p("Before") + table))
		const after = await summarizeDocument(docx(p("Before") + table + p("")))

		expect(ends?.counts["bodyEndsWithTable"]).toBe(1)
		expect(after?.counts["bodyEndsWithTable"]).toBe(0)
		expect(acceptedReason(entry(["txt: timed out"], [], { bodyEndsWithTable: 1 }), "L1 txt: timed out", "lo-26.8.1.1")).toBeDefined()
		expect(acceptedReason(entry(["txt: timed out"], []), "L1 txt: timed out", "lo-26.8.1.1")).toBeUndefined()
		expect(acceptedReason(entry(["txt: timed out"], [], { bodyEndsWithTable: 1 }), "L1 txt: timed out", "lo-7.5.0.3")).toBeUndefined()
	})

	it("leaves out a format the original does not convert to either, and a DOCX LibreOffice wrote none of from its ODT", () => {
		expect(acceptedReason(entry(["txt: no output"], ["txt"]), "L1 txt: no output", "lo-7.5.0.3")).toBeDefined()
		expect(acceptedReason(entry(["pdf: timed out"], ["txt"]), "L1 pdf: timed out", "lo-7.5.0.3")).toBeUndefined()
		expect(acceptedReason(entry([], []), "G7 back gate: no file", "lo-7.5.0.3")).toBeDefined()
		// Without an ODT there is nothing LibreOffice failed on its own.
		expect(acceptedReason(entry(["odt: timed out"], []), "G7 back gate: no file", "lo-7.5.0.3")).toBeUndefined()
	})
})

describe("floating drawings in notes under tall footers", () => {
	const inline = (cy: number): string =>
		`<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="${String(cy)}"/><wp:docPr id="3" name="F"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:inline></w:drawing></w:r>`
	const anchored =
		'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="character"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="A"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:anchor></w:drawing></w:r>'
	const notes = `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p>${anchored}</w:p></w:footnote>`
	const read = (footerHeight: number) =>
		summarizeDocument(
			buildDocx({
				body: '<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
				footnotes: notes,
				footers: { rIdF: `<w:p>${inline(footerHeight)}</w:p>` }
			})
		)

	it("counts a note's floating drawings once the footer's drawings take more than half the page", async () => {
		expect((await read(7305403))?.counts["noteFloatsUnderTallFooters"]).toBe(1)
		expect((await read(457200))?.counts["noteFloatsUnderTallFooters"]).toBe(0)
	})
})
