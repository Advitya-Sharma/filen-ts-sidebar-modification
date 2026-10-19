import { describe, expect, it } from "vitest"
import { buildDocx } from "./docxSynthetic"
import { summarizeDocument } from "./fuzz/saves"
import { normalizeLines } from "./interop/textDelta"

// The LibreOffice leg's readings of Microsoft Equation 3.0 objects, rows' grid gaps, symbol characters,
// and the witnesses of LibreOffice's own readings of comments, tracked paragraph marks and content controls.

const AUTHOR = 'w:author="Ada" w:date="2026-01-01T00:00:00Z"'
const INSERTED = `<w:rPr><w:ins w:id="90" ${AUTHOR}/></w:rPr>`
const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const p = (text: string): string => `<w:p>${r(text)}</w:p>`
const MATH = "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"
const ANCHOR =
	'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="7" name="Shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'

async function counts(spec: Parameters<typeof buildDocx>[0]): Promise<Readonly<Record<string, number>>> {
	const summary = await summarizeDocument(buildDocx(spec))

	if (summary === undefined) throw new Error("unreadable")

	return summary.counts
}

function cell(text: string): string {
	return `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`
}

describe("our reading as an ODF body holds it", () => {
	it("counts a Microsoft Equation 3.0 object as an equation, as LibreOffice reads it", async () => {
		const object = (progId: string): string =>
			`<w:p><w:r><w:object><v:shape id="s1" style="width:20pt;height:10pt"><v:imagedata r:id="rIdImg" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="${progId}" ShapeID="s1" DrawAspect="Content" ObjectID="_1" r:id="rIdOle"/></w:object></w:r></w:p>`
		const spec = (progId: string) => ({
			body: p("Before") + object(progId),
			relationships: [
				{ id: "rIdImg", type: "image", target: "media/image1.wmf" },
				{ id: "rIdOle", type: "oleObject", target: "embeddings/oleObject1.bin" }
			],
			parts: [
				{ name: "word/media/image1.wmf", data: new Uint8Array([1, 2, 3]) },
				{ name: "word/embeddings/oleObject1.bin", data: new Uint8Array([1, 2, 3]) }
			]
		})

		expect((await counts(spec("Equation.3")))["equations"]).toBe(1)
		expect((await counts(spec("Excel.Sheet.12")))["equations"]).toBe(0)
	})

	it("fills a row's grid gaps with an empty cell each, as an ODF row holds them", async () => {
		const table = (trPr: string, cells: number): string =>
			`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${trPr}${cell("X").repeat(cells)}</w:tr></w:tbl>`
		const plain = await counts({ body: p("A") + table("", 3) + p("B") })
		const gaps = await counts({
			body: p("A") + table('<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>', 1) + p("B")
		})

		expect(gaps["cells"]).toBe(plain["cells"])
		expect(gaps["bodyParagraphs"]).toBe(plain["bodyParagraphs"])
	})

	it("leaves out symbol characters LibreOffice keeps as C1 control code points", () => {
		// w:sym w:char="0081" (the corpus's oxsdk__equation-1, "UniversalMath1 BT").
		expect(normalizeLines("text\u0081 ")).toEqual(["text"])
	})
})

describe("witnesses of LibreOffice's own readings", () => {
	it("counts the equations of comments", async () => {
		const found = await counts({
			body: `<w:p><w:commentRangeStart w:id="0"/>${r("Body")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`,
			comments: `<w:comment w:id="0" ${AUTHOR}><w:p>${r("Note ")}${MATH}</w:p><w:p><m:oMathPara>${MATH}${MATH}</m:oMathPara></w:p></w:comment>`
		})

		expect(found["commentEquations"]).toBe(3)
		expect(found["equations"]).toBe(3)
	})

	it("counts empty comments after the end of a range holding a shape, before its reference", async () => {
		const comments = [0, 1, 2].map(id => `<w:comment w:id="${String(id)}" ${AUTHOR}>${p(`C${String(id)}`)}</w:comment>`).join("")
		const S = (id: number): string => `<w:commentRangeStart w:id="${String(id)}"/>`
		const E = (id: number): string => `<w:commentRangeEnd w:id="${String(id)}"/>`
		const R = (id: number): string => `<w:r><w:commentReference w:id="${String(id)}"/></w:r>`
		const of = async (body: string): Promise<number> =>
			(await counts({ body: `<w:p>${body}</w:p>`, comments }))["commentsAfterShapeRange"] ?? 0

		expect(await of(`${r("A ")}${S(0)}${ANCHOR}${E(0)}${S(1)}${E(1)}${R(1)}${R(0)}`)).toBe(1)
		expect(await of(`${r("A ")}${S(0)}${ANCHOR}${E(0)}${S(1)}${E(1)}${R(1)}${S(2)}${r("q")}${E(2)}${R(2)}${R(0)}`)).toBe(1)
		expect(await of(`${r("A ")}${S(0)}${r("G")}${E(0)}${S(1)}${E(1)}${R(1)}${R(0)}`)).toBe(0)
	})

	it("counts inserted paragraph marks ending a section with its own headers, or in a paragraph with a floating drawing", async () => {
		const section = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:type w:val="continuous"/></w:sectPr>'
		const found = await counts({
			body:
				`<w:p><w:pPr>${INSERTED}${section}</w:pPr>${r("you mu")}</w:p>${p("st ensure")}` +
				`<w:p><w:pPr>${INSERTED}</w:pPr>${ANCHOR}${r("Social and emo")}</w:p>${p("tional")}` +
				`<w:p><w:pPr>${INSERTED}</w:pPr>${r("plain")}</w:p>${p("next")}`,
			headers: { rIdH1: p("Header") }
		})

		expect(found["insertedSectionMarks"]).toBe(1)
		expect(found["insertedAnchorMarks"]).toBe(1)
	})

	it("counts date controls bound to custom XML, and tracked changes inside block-level controls", async () => {
		const bound =
			'<w:p><w:sdt><w:sdtPr><w:id w:val="1"/><w:dataBinding w:xpath="/a[1]/b[1]" w:storeItemID="{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}"/><w:date><w:dateFormat w:val="M/d/yyyy"/><w:storeMappedDataAs w:val="dateTime"/></w:date></w:sdtPr><w:sdtContent>' +
			`${r("[Pick the date]")}</w:sdtContent></w:sdt></w:p>`
		const block =
			'<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>' +
			`<w:p>${r("[")}<w:del w:id="91" ${AUTHOR}><w:r><w:delText>Deg</w:delText></w:r></w:del>${r("ree]")}</w:p>` +
			`<w:p><w:pPr>${INSERTED}</w:pPr>${r("Cli")}</w:p>${p("ck here")}</w:sdtContent></w:sdt>`
		const found = await counts({ body: bound + block + p("After") })

		expect(found["boundDateControls"]).toBe(1)
		expect(found["blockControlRevisions"]).toBe(2)
		expect(found["blockControlInsertedMarks"]).toBe(1)
	})
})
