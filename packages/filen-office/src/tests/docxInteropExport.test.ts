import { describe, expect, it } from "vitest"
import { buildDocx, p } from "./docxSynthetic"
import { summarizeDocument } from "./fuzz/saves"
import { normalizeLines } from "./interop/textDelta"

// What LibreOffice writes back differently in its DOCX of a document it read as Word does (the G7 leg), counted on
// our side so interop/accepted.ts accepts only where the construct is: links to "_Toc" bookmarks, Word's watermark
// shape in the body, a link before a drawing group with a text box, tracked changes in a placeholder control.

const AUTHOR = 'w:author="Ada" w:date="2026-01-01T00:00:00Z"'

async function counts(body: string): Promise<Readonly<Record<string, number>>> {
	const summary = await summarizeDocument(buildDocx({ body }))

	if (summary === undefined) throw new Error("unreadable")

	return summary.counts
}

function link(anchor: string, text: string): string {
	return `<w:p><w:hyperlink w:anchor="${anchor}"><w:r><w:t>${text}</w:t></w:r></w:hyperlink></w:p>`
}

function watermark(id: string): string {
	return (
		"<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:pict>" +
		`<v:shape id="${id}" o:spid="_x0000_s3073" type="#_x0000_t136" style="position:absolute;width:494pt;height:164pt;rotation:315" fillcolor="#e4e4e4" stroked="f">` +
		'<v:textpath style="font-family:&quot;Arial&quot;;font-size:1pt" string="DEMO"/></v:shape></w:pict></w:r></w:p>'
	)
}

const GROUP =
	'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
	'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
	'<wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="1" name="Group"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
	'<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr/><wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></wpg:wgp>' +
	"</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"

function placeholder(content: string, showing = true): string {
	return `<w:sdt><w:sdtPr><w:id w:val="5"/>${showing ? "<w:showingPlcHdr/>" : ""}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`
}

describe("constructs LibreOffice writes back differently", () => {
	it("counts links to a _Toc bookmark outside an index", async () => {
		expect((await counts(link("_Toc996", "c") + link("Other", "d")))["tocLinks"]).toBe(1)

		// Entries of a table of contents (its field's result) are left out on both sides already.
		const toc =
			'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
			'<w:hyperlink w:anchor="_Toc1"><w:r><w:t>Heading</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

		expect((await counts(toc))["tocLinks"]).toBe(0)
	})

	it("counts Word's watermark shape in the body, not other VML shapes", async () => {
		expect((await counts(watermark("PowerPlusWaterMarkObject1288920517")))["bodyWatermarks"]).toBe(1)
		expect((await counts(watermark("Shape1")))["bodyWatermarks"]).toBe(0)
	})

	it("counts a drawing group with a text box after a link in its paragraph", async () => {
		const linked = `<w:p><w:hyperlink w:anchor="Other"><w:r><w:t>Filen</w:t></w:r></w:hyperlink>${GROUP}</w:p>`

		expect((await counts(linked))["linkBeforeTextboxGroup"]).toBe(1)
		expect(
			(await counts(`<w:p>${GROUP}<w:hyperlink w:anchor="Other"><w:r><w:t>Filen</w:t></w:r></w:hyperlink></w:p>`))[
				"linkBeforeTextboxGroup"
			]
		).toBe(0)
	})

	it("counts tracked changes in a block content control showing its placeholder", async () => {
		const changed = `<w:p><w:r><w:t>[</w:t></w:r><w:del w:id="1" ${AUTHOR}><w:r><w:delText>Ins</w:delText></w:r></w:del><w:r><w:t>ert text.</w:t></w:r></w:p>`
		const mark = `<w:p><w:pPr><w:rPr><w:del w:id="2" ${AUTHOR}/></w:rPr></w:pPr><w:r><w:t>[Degree]</w:t></w:r></w:p>`

		expect((await counts(p("Before") + placeholder(changed) + p("After")))["placeholderChanges"]).toBe(1)
		expect((await counts(p("Before") + placeholder(mark) + p("Cli")))["placeholderChanges"]).toBe(1)
		// Not showing its placeholder, or no change: LibreOffice reads it as text.
		expect((await counts(p("Before") + placeholder(changed, false) + p("After")))["placeholderChanges"]).toBe(0)
		expect((await counts(p("Before") + placeholder(p("[Degree]")) + p("Cli")))["placeholderChanges"]).toBe(0)
	})
})

describe("text deltas", () => {
	it("leave out C1 controls, which LibreOffice reads a symbol-font w:sym as", () => {
		// oxsdk__equation-1: w:sym w:font="UniversalMath1 BT" w:char="0081"
		expect(normalizeLines("故结论正确． \u0081 \nNext")).toEqual(["故结论正确．", "Next"])
	})
})
