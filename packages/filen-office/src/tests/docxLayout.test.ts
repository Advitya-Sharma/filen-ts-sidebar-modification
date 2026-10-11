import { setImmediate } from "node:timers/promises"
import { describe, expect, it } from "vitest"
import {
	CLIP_INTS,
	DocumentLayout,
	GRAPHIC_FLOATS,
	GRAPHIC_INTS,
	OBJECT_FLOATS,
	OBJECT_INTS,
	PageGraphic,
	Region,
	RevisionView,
	TrackedLook,
	layoutDocument,
	type LayoutEvent,
	type LayoutOptions
} from "@filen/office/core/docx"
import { buildDocx, type SyntheticDocx } from "./docxSynthetic"
import { HmtxMeasurer, layOut, loadDocx, pageLines, type LineInfo } from "./layoutFixtures"
import { wordLikeDocument } from "./wordLike"

// Letter paper, 1 in margins: a 624 x 864 px text area. Without a styles part text is Times New Roman 10 pt
// (Tinos metrics: 15.33 px lines), so a page holds 56 lines.
const SECTION =
	'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
const LINE = 15.33
const WORDS = "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua"

function p(text: string, pPr = "", rPr = ""): string {
	return `<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r>${rPr === "" ? "" : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
}

function lines(count: number): string {
	return Array.from({ length: count }, (_, k) => p(`line ${k + 1}`)).join("")
}

function words(count: number): string {
	const all = WORDS.split(" ")

	return Array.from({ length: count }, (_, k) => all[k % all.length]).join(" ")
}

async function layout(spec: SyntheticDocx & { section?: string }, options: LayoutOptions = {}): Promise<DocumentLayout> {
	const body = `${spec.body ?? ""}${spec.section ?? SECTION}`

	return layOut(buildDocx({ ...spec, body }), options)
}

function body(layout: DocumentLayout, page: number): LineInfo[] {
	return pageLines(layout, page).filter(line => line.region === Region.BODY)
}

function texts(infos: readonly LineInfo[]): string[] {
	return infos.map(info => info.text)
}

function graphics(layout: DocumentLayout, page: number, kind: number): { x: number; y: number; width: number; height: number }[] {
	const list = layout.pages[page]
	const out: { x: number; y: number; width: number; height: number }[] = []

	if (list === undefined) return out

	for (let g = 0; g < list.graphicCount; g++) {
		if (list.graphicInts[g * GRAPHIC_INTS] !== kind) continue

		const f = g * GRAPHIC_FLOATS

		out.push({
			x: list.graphicFloats[f] ?? 0,
			y: list.graphicFloats[f + 1] ?? 0,
			width: list.graphicFloats[f + 2] ?? 0,
			height: list.graphicFloats[f + 3] ?? 0
		})
	}

	return out
}

const NUMBERING = {
	relationships: [{ id: "rId50", type: "numbering", target: "numbering.xml" }],
	parts: [
		{
			name: "word/numbering.xml",
			contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
			data:
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
				'<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>' +
				'<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
				'<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1)"/><w:lvlJc w:val="right"/>' +
				'<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
				'<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
		}
	]
}

describe("pages", () => {
	it("fills pages line by line and delivers them incrementally", async () => {
		const loaded = await loadDocx(buildDocx({ body: lines(150) + SECTION }))
		const events: LayoutEvent[] = []
		const result = layoutDocument(loaded.model, loaded.resolver, loaded.fonts, new HmtxMeasurer(), { store: loaded.store, sliceMs: 0 })

		result.subscribe(event => events.push(event))
		await result.done

		expect(result.final).toBe(true)
		expect(result.pages.length).toBe(3)
		expect(texts(body(result, 0)).length).toBe(56)
		expect(texts(body(result, 1))[0]).toBe("line 57")
		expect(body(result, 0)[0]?.top).toBe(96)
		expect(body(result, 0)[1]?.top).toBeCloseTo(96 + LINE, 1)
		expect(events.filter(event => !event.final).map(event => event.from)).toEqual([0, 1, 2])
		expect(events[events.length - 1]).toMatchObject({ type: "pages", from: 0, to: 3, final: true })
		expect(result.pages[0]?.width).toBe(816)
		expect(result.pages[0]?.bodyFirst).toBeGreaterThanOrEqual(0)
	})

	it("lays out a 300-page document quickly, the first page first", async () => {
		const loaded = await loadDocx(buildDocx({ document: new TextDecoder().decode(wordLikeDocument(4000)) }))
		const started = performance.now()
		let first = 0
		const result = layoutDocument(loaded.model, loaded.resolver, loaded.fonts, new HmtxMeasurer(), {
			store: loaded.store,
			sliceMs: 8,
			yield: () => setImmediate()
		})

		result.subscribe(() => {
			if (first === 0) first = performance.now() - started
		})
		await result.done

		const total = performance.now() - started

		expect(result.pages.length).toBeGreaterThan(300)
		expect(first).toBeLessThan(300)
		// About 0.4 s on a laptop; generous for shared CI machines.
		expect(total).toBeLessThan(3000)
	})

	it("reports the style, decoration and object entries each event adds", async () => {
		const loaded = await loadDocx(buildDocx({ body: p("plain") + lines(60) + p("bold", "", '<w:b/><w:u w:val="single"/>') + SECTION }))
		const events: LayoutEvent[] = []
		const result = layoutDocument(loaded.model, loaded.resolver, loaded.fonts, new HmtxMeasurer(), { store: loaded.store })

		result.subscribe(event => events.push(event))
		await result.done

		const [first, second] = events

		expect(first?.styles.from).toBe(0)
		expect(second?.styles.from).toBe(first?.styles.to)
		expect((second?.styles.to ?? 0) > (second?.styles.from ?? 0)).toBe(true)
		expect(second?.decorations.to).toBe(result.decorations.length)
		expect(events[events.length - 1]?.styles.to).toBe(result.styleInfo.length)
	})

	it("anchors pages at the block, table row and text offset where they start", async () => {
		const row = (k: number): string => `<w:tr><w:tc>${p(`r${k}`)}</w:tc></w:tr>`
		const result = await layout({
			body:
				p(words(30)) +
				`<w:tbl><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>${Array.from({ length: 70 }, (_, k) => row(k)).join("")}</w:tbl>` +
				lines(50)
		})
		const [first, second, third] = result.anchors

		expect(first).toMatchObject({ block: 1, row: -1, offset: 0 })
		expect(result.model.parts[0]?.kinds[second?.block ?? 0]).toBe(11)
		expect(second?.row).toBeGreaterThan(40)
		expect(third?.row).toBe(-1)
		const split = await layout({ body: lines(50) + p(words(200)) })

		// The 51st paragraph (three items each: paragraph, run, text).
		expect(split.anchors[1]).toMatchObject({ block: 1 + 50 * 3, row: -1 })
		expect(split.anchors[1]?.offset).toBeGreaterThan(0)
		expect(texts(body(split, 1))[0]).toBe(
			words(200)
				.slice(split.anchors[1]?.offset ?? 0)
				.slice(0, (body(split, 1)[0]?.text ?? "").length)
		)
	})

	it("can be cancelled between slices", async () => {
		const loaded = await loadDocx(buildDocx({ body: lines(2000) + SECTION }))
		const result = layoutDocument(loaded.model, loaded.resolver, loaded.fonts, new HmtxMeasurer(), { store: loaded.store, sliceMs: 0 })

		result.subscribe(() => result.cancel())
		await result.done

		expect(result.cancelled).toBe(true)
		expect(result.final).toBe(false)
		expect(result.pages.length).toBeLessThan(36)
	})

	it("keeps display lists compact", async () => {
		const result = await layout({ body: Array.from({ length: 200 }, () => p(words(60))).join("") })
		const lineCount = result.pages.reduce((sum, page) => sum + page.text.lineCount, 0)

		expect(result.pages.length).toBeGreaterThan(8)
		// One fragment per line of plain text: 48 bytes of line and 36 of fragment, plus the page's sources.
		expect(result.byteLength / lineCount).toBeLessThan(90)
	})

	it("breaks pages and columns where the text asks", async () => {
		const result = await layout({
			body:
				p("one") +
				'<w:p><w:r><w:t>two</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>' +
				p("three") +
				p("four", "<w:pageBreakBefore/>") +
				'<w:p><w:r><w:t>left</w:t></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t>right</w:t></w:r></w:p>',
			section:
				'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="720"/></w:sectPr>'
		})

		expect(texts(body(result, 0))).toEqual(["one", "two\u000b"])
		expect(texts(body(result, 1))).toEqual(["three"])
		expect(texts(body(result, 2))).toEqual(["four", "left\u000b", "right"])
		// The second column starts at the top, half the text width plus half the gap to the right.
		expect(body(result, 2)[2]?.top).toBe(96)
		expect(body(result, 2)[2]?.x).toBeCloseTo(96 + (624 + 48) / 2, 0)
	})
})

describe("paragraph spacing", () => {
	const spaced = p("a", '<w:spacing w:before="0" w:after="240"/>') + p("b", '<w:spacing w:before="480" w:after="0"/>')

	it("takes the larger of space after and space before (HTML paragraph auto spacing)", async () => {
		const result = await layout({ body: spaced })
		const [a, b] = body(result, 0)

		expect((b?.top ?? 0) - (a?.bottom ?? 0)).toBeCloseTo(32, 1)
	})

	it("adds them when w:doNotUseHTMLParagraphAutoSpacing is set", async () => {
		const result = await layout({ body: spaced, settings: "<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>" })
		const [a, b] = body(result, 0)

		expect((b?.top ?? 0) - (a?.bottom ?? 0)).toBeCloseTo(48, 1)
	})

	it("drops spacing between paragraphs of one style with contextual spacing", async () => {
		const result = await layout({
			body: p("a", '<w:contextualSpacing/><w:spacing w:after="240"/>') + p("b", '<w:contextualSpacing/><w:spacing w:before="240"/>')
		})
		const [a, b] = body(result, 0)

		expect((b?.top ?? 0) - (a?.bottom ?? 0)).toBeCloseTo(0, 1)
	})

	it("drops space before at the top of a page the text flowed onto", async () => {
		const result = await layout({ body: lines(56) + p("next", '<w:spacing w:before="240"/>') })

		expect(body(result, 1)[0]?.top).toBe(96)
	})

	it("uses HTML auto spacing: nothing before, 12 pt after, fixed 5 pt and 10 pt under the compatibility switch", async () => {
		const auto =
			p("a", '<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/>') +
			p("b", '<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/>')
		const html = await layout({ body: auto })
		const fixed = await layout({ body: auto, settings: "<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>" })
		const gap = (result: DocumentLayout): number => (body(result, 0)[1]?.top ?? 0) - (body(result, 0)[0]?.bottom ?? 0)

		expect(body(html, 0)[0]?.top).toBe(96)
		expect(gap(html)).toBeCloseTo(16, 1)
		expect(body(fixed, 0)[0]?.top).toBeCloseTo(96 + 20 / 3, 1)
		expect(gap(fixed)).toBeCloseTo(40 / 3 + 20 / 3, 1)
	})

	it("applies exact and at-least line spacing", async () => {
		const result = await layout({
			body:
				p(words(40), '<w:spacing w:line="480" w:lineRule="exact"/>') +
				p(words(40), '<w:spacing w:line="100" w:lineRule="atLeast"/>')
		})
		const all = body(result, 0)

		expect((all[0]?.bottom ?? 0) - (all[0]?.top ?? 0)).toBeCloseTo(32, 1)
		expect((all[all.length - 1]?.bottom ?? 0) - (all[all.length - 1]?.top ?? 0)).toBeCloseTo(LINE, 1)
	})
})

describe("keep and widow rules", () => {
	it("moves a heading kept with the next paragraph", async () => {
		const result = await layout({ body: lines(55) + p("Heading", "<w:keepNext/>") + p("text") })

		expect(texts(body(result, 1)).slice(0, 2)).toEqual(["Heading", "text"])
		expect(body(result, 0).length).toBe(55)
	})

	it("keeps a heading with a following paragraph that splits", async () => {
		const result = await layout({ body: lines(40) + p("Heading", "<w:keepNext/>") + p(words(400)) })
		const first = texts(body(result, 0))

		expect(first[40]).toBe("Heading")
		expect(first.length).toBe(56)
	})

	it("drops the keep when a run of keep-with-next paragraphs is longer than a page", async () => {
		const result = await layout({ body: Array.from({ length: 200 }, (_, k) => p(`kept ${k}`, "<w:keepNext/>")).join("") })

		expect(result.pages.length).toBe(4)
		expect(body(result, 0).length).toBe(56)
	})

	it("takes back a moved paragraph's anchored objects", async () => {
		const anchor =
			'<w:r><w:drawing><wp:anchor behindDoc="0" allowOverlap="1" layoutInCell="1" relativeHeight="1" simplePos="0" locked="0"><wp:simplePos x="0" y="0"/>' +
			'<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
			'<wp:extent cx="952500" cy="952500"/><wp:wrapSquare wrapText="right"/><wp:docPr id="2" name="f"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:anchor></w:drawing></w:r>'
		const result = await layout({
			body: lines(55) + `<w:p><w:pPr><w:keepLines/></w:pPr>${anchor}<w:r><w:t>${words(60)}</w:t></w:r></w:p>`
		})

		expect(result.pages.map(page => page.objectCount)).toEqual([0, 1])
	})

	it("never leaves a single line of a paragraph at the bottom or top of a page", async () => {
		const long = p(words(60))
		const orphan = await layout({ body: lines(55) + long })
		const widow = await layout({ body: lines(53) + long })

		// Only the first line would fit: the whole paragraph moves.
		expect(body(orphan, 0).length).toBe(55)
		// Three lines would fit but the last would be alone: two lines move.
		expect(body(widow, 0).length).toBe(55)
		expect(body(widow, 1).length).toBe(2)
	})

	it("keeps a paragraph's lines together with keepLines", async () => {
		const result = await layout({ body: lines(54) + p(words(60), "<w:keepLines/>") })

		expect(body(result, 0).length).toBe(54)
	})

	it("lets widowed lines stay without widow control", async () => {
		const result = await layout({ body: lines(55) + p(words(60), '<w:widowControl w:val="0"/>') })

		expect(body(result, 0).length).toBe(56)
	})
})

describe("sections", () => {
	const section = (type: string, extra = ""): string =>
		`<w:p><w:pPr><w:sectPr><w:type w:val="${type}"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${extra}</w:sectPr></w:pPr></w:p>`

	it("starts sections on new pages, continuously, and on odd pages with a blank page between", async () => {
		// A section's type says how it starts: the second and third sections continue on the first page.
		const result = await layout({
			body: p("one") + section("nextPage") + p("two") + section("continuous") + p("three") + section("continuous") + p("four"),
			section:
				'<w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		})

		expect(result.pages.map(page => page.number)).toEqual([1, 2, 3])
		expect(texts(body(result, 0))).toEqual(["one", "", "two", "", "three", ""])
		expect(texts(body(result, 1))).toEqual([])
		expect(texts(body(result, 2))).toEqual(["four"])
		expect(result.pages[2]?.section).toBe(3)

		const next = await layout({ body: p("one") + section("nextPage") + p("two") + section("nextPage") + p("three") })

		expect(next.pages.map(page => texts(body(next, page.index)))).toEqual([["one", ""], ["two", ""], ["three"]])
	})

	it("draws a page with the vertical alignment and borders of its own section", async () => {
		const border = '<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="24" w:space="24"/></w:pgBorders>'
		const result = await layout({
			body: p("Cover") + section("nextPage", `<w:vAlign w:val="center"/>${border}`) + p("Body"),
			section: SECTION
		})

		expect(body(result, 0)[0]?.top).toBeGreaterThan(400)
		expect(body(result, 1)[0]?.top).toBe(96)
		expect(graphics(result, 0, PageGraphic.BORDER).length).toBe(1)
		expect(graphics(result, 1, PageGraphic.BORDER).length).toBe(0)
	})

	it("lays out landscape pages, columns with a separator and gutters", async () => {
		const result = await layout({
			body: lines(120),
			section:
				'<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="720"/><w:cols w:num="2" w:space="720" w:sep="1"/></w:sectPr>'
		})
		const first = body(result, 0)

		expect(result.pages[0]?.width).toBe(1056)
		// 624 px tall text area: 40 lines per column.
		expect(first.length).toBe(80)
		expect(first[0]?.x).toBe(144)
		expect(first[40]?.top).toBe(96)
		expect(graphics(result, 0, PageGraphic.COLUMN_SEPARATOR).length).toBe(1)
	})

	it("numbers pages in fields, restarts and formats page numbers", async () => {
		const footer =
			'<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
		const result = await layout({
			body: lines(60),
			footers: { rId20: footer },
			section:
				'<w:sectPr><w:footerReference w:type="default" r:id="rId20"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="3"/></w:sectPr>'
		})
		const footers = [0, 1].map(page =>
			pageLines(result, page)
				.filter(line => line.region === Region.FOOTER)
				.map(line => line.text)
		)

		expect(footers).toEqual([["Page iii of 2"], ["Page iv of 2"]])
		expect(result.pages[1]?.numberText).toBe("iv")
		// The footer sits above its distance from the page bottom.
		expect(pageLines(result, 0).find(line => line.region === Region.FOOTER)?.bottom).toBeCloseTo(1056 - 48, 1)
	})

	it("uses the first-page header and even-page headers", async () => {
		const header = (text: string): string => p(text)
		const result = await layout({
			body: lines(170),
			headers: { rId21: header("default"), rId22: header("first"), rId23: header("even") },
			settings: "<w:evenAndOddHeaders/>",
			section:
				'<w:sectPr><w:headerReference w:type="default" r:id="rId21"/><w:headerReference w:type="first" r:id="rId22"/><w:headerReference w:type="even" r:id="rId23"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:titlePg/></w:sectPr>'
		})
		const headers = result.pages.map((_, page) => pageLines(result, page).find(line => line.region === Region.HEADER)?.text)

		expect(headers).toEqual(["first", "even", "default", "even"])
	})

	it("pushes the body down below a tall header", async () => {
		const result = await layout({
			body: p("body"),
			headers: { rId24: lines(8) },
			section:
				'<w:sectPr><w:headerReference w:type="default" r:id="rId24"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
		})

		expect(body(result, 0)[0]?.top).toBeCloseTo(48 + 8 * LINE, 1)
	})

	it("draws page borders measured from the page edge", async () => {
		const result = await layout({
			body: p("text"),
			section:
				'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
				'<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="12" w:space="24"/><w:left w:val="single" w:sz="12" w:space="24"/><w:bottom w:val="single" w:sz="12" w:space="24"/><w:right w:val="single" w:sz="12" w:space="24"/></w:pgBorders></w:sectPr>'
		})
		const lines = graphics(result, 0, PageGraphic.BORDER)

		expect(lines.length).toBe(4)
		expect(lines[0]).toMatchObject({ x: 32, y: 32, width: 816 - 64, height: 2 })
	})

	it("numbers lines in the margin", async () => {
		const result = await layout({
			body: lines(10),
			section:
				'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:lnNumType w:countBy="5" w:distance="360"/></w:sectPr>'
		})
		const numbers = pageLines(result, 0).filter(line => line.region === Region.LINE_NUMBER)

		expect(numbers.map(line => line.text)).toEqual(["5", "10"])
		expect((numbers[0]?.x ?? 0) + (numbers[0]?.width ?? 0)).toBeCloseTo(96 - 24, 1)
	})
})

describe("lists, fields and hidden text", () => {
	it("draws list markers with their suffix tab and right-aligned markers left of the indent", async () => {
		const result = await layout({
			body:
				p("first", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
				p("second", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
				p("roman", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'),
			...NUMBERING
		})
		const all = body(result, 0)

		expect(texts(all)).toEqual(["1.\tfirst", "2.\tsecond", "i)\troman"])
		expect(all[0]?.x).toBe(120)
		// A right-aligned marker ends at the hanging indent's start.
		expect((all[2]?.x ?? 0) + 0).toBeLessThan(120)
	})

	it("shows field results, hides field codes and hidden text", async () => {
		const result = await layout({
			body:
				'<w:p><w:r><w:t xml:space="preserve">see </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF bookmark </w:instrText></w:r>' +
				'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
				'<w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r><w:r><w:t xml:space="preserve"> page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
				'<w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
		})

		expect(texts(body(result, 0))).toEqual(["see result page 1"])
	})

	it("joins a paragraph whose mark is hidden to the next one", async () => {
		const result = await layout({ body: p("first", "<w:rPr><w:vanish/></w:rPr>") + p("second") })

		expect(texts(body(result, 0))).toEqual(["firstsecond"])
	})

	it("shows tracked changes per view", async () => {
		const tracked =
			'<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r><w:ins w:id="1" w:author="A"><w:r><w:t>added</w:t></w:r></w:ins>' +
			'<w:del w:id="2" w:author="A"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>'
		const final = await layout({ body: tracked })
		const original = await layout({ body: tracked }, { view: RevisionView.Original })
		const markup = await layout({ body: tracked }, { view: RevisionView.Markup })

		expect(texts(body(final, 0))).toEqual(["keep added"])
		expect(texts(body(original, 0))).toEqual(["keep removed"])
		expect(texts(body(markup, 0))).toEqual(["keep addedremoved"])
		expect(markup.styleInfo.map(info => info?.revision)).toEqual(expect.arrayContaining([TrackedLook.INSERTED, TrackedLook.DELETED]))
	})

	it("draws grouped paragraph borders as one box with a rule between", async () => {
		const borders =
			'<w:pBdr><w:top w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:between w:val="single" w:sz="8"/></w:pBdr>'
		const result = await layout({ body: p("a", borders) + p("b", borders) + p("c") })
		const lines = graphics(result, 0, PageGraphic.BORDER)

		// Top, bottom and one rule between: one box for both paragraphs.
		expect(lines.length).toBe(3)
		expect(lines.every(line => line.height < 2)).toBe(true)
	})
})

describe("tables", () => {
	const cell = (text: string, tcPr = ""): string => `<w:tc>${tcPr === "" ? "" : `<w:tcPr>${tcPr}</w:tcPr>`}${p(text)}</w:tc>`
	const grid = '<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>'
	const borders =
		'<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
		'<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>'
	const table = (rows: string, tblPr = ""): string => `<w:tbl><w:tblPr>${borders}${tblPr}</w:tblPr>${grid}${rows}</w:tbl>`

	it("places cells on the grid with spans and draws their borders", async () => {
		const result = await layout({
			body:
				table(
					`<w:tr>${cell("a")}${cell("b")}${cell("c")}</w:tr><w:tr>${cell("wide", '<w:gridSpan w:val="2"/>')}${cell("d")}</w:tr>`
				) + p("after")
		})
		const all = body(result, 0)

		expect(texts(all)).toEqual(["a", "b", "c", "wide", "d", "after"])
		expect(all.map(line => line.x)).toEqual([96, 288, 480, 96, 480, 96])
		expect(all[3]?.top).toBeCloseTo(96 + LINE, 1)
		expect(all[5]?.top).toBeCloseTo(96 + 2 * LINE, 1)
		// Two horizontal lines per cell edge row plus the vertical edges.
		expect(graphics(result, 0, PageGraphic.BORDER).length).toBe(3 + 3 + 2 + 4 + 3)
	})

	it("applies row heights, vertical alignment and vertical merges", async () => {
		const result = await layout({
			body: table(
				`<w:tr><w:trPr><w:trHeight w:val="1440" w:hRule="exact"/></w:trPr>${cell("top")}${cell("middle", '<w:vAlign w:val="center"/>')}${cell("merged", '<w:vMerge w:val="restart"/>')}</w:tr>` +
					`<w:tr>${cell("x")}${cell("y")}${cell("", "<w:vMerge/>")}</w:tr>`
			)
		})
		const all = body(result, 0)

		expect(all.find(line => line.text === "top")?.top).toBe(96)
		expect(all.find(line => line.text === "middle")?.top).toBeCloseTo(96 + (96 - LINE) / 2, 1)
		expect(all.find(line => line.text === "x")?.top).toBe(192)
		expect(all.filter(line => line.text === "").length).toBe(0)
	})

	it("repeats header rows and splits rows across pages", async () => {
		const header = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell("H1")}${cell("H2")}${cell("H3")}</w:tr>`
		const rows = Array.from({ length: 60 }, (_, k) => `<w:tr>${cell(`r${k}`)}${cell("")}${cell("")}</w:tr>`).join("")
		const result = await layout({ body: table(header + rows) })

		expect(result.pages.length).toBe(2)
		expect(body(result, 1)[0]?.text).toBe("H1")
		expect(body(result, 1)[3]?.text).toBe("r55")

		const tall = `<w:tr>${cell("")}${cell("")}<w:tc>${lines(80)}</w:tc></w:tr>`
		const split = await layout({ body: p("before") + table(tall) })

		// "before", the empty paragraphs of the two short cells and 55 lines of the tall one.
		expect(body(split, 0).length).toBe(58)
		expect(texts(body(split, 1)).slice(0, 2)).toEqual(["line 56", "line 57"])
	})

	it("lays out thousands of vertically merged rows", async () => {
		const rows = Array.from(
			{ length: 4000 },
			(_, k) =>
				`<w:tr>${cell(`a${k}`, k % 2 === 0 ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>")}${cell(`b${k}`)}${cell("")}</w:tr>`
		).join("")
		const result = await layout({ body: table(rows) })

		expect(result.pages.length).toBeGreaterThan(30)
	})

	it("clips the cells of a row of exact height", async () => {
		const result = await layout({
			body: table(
				`<w:tr><w:trPr><w:trHeight w:val="300" w:hRule="exact"/></w:trPr>${cell("a")}<w:tc>${lines(5)}</w:tc>${cell("c")}</w:tr>`
			)
		})
		const page = result.pages[0]

		expect(page?.clipCount).toBe(3)
		expect(page?.clipFloats[3]).toBeCloseTo(20, 1)
		// The middle cell's five lines are drawn through its clip.
		expect((page?.clipInts[CLIP_INTS + 1] ?? 0) - (page?.clipInts[CLIP_INTS] ?? 0)).toBe(5)
	})

	it("resolves shared cell edges by weight, then style, then the darker colour", async () => {
		const red = '<w:tcBorders><w:right w:val="single" w:sz="4" w:color="FF0000"/></w:tcBorders>'
		const thick = '<w:tcBorders><w:left w:val="double" w:sz="4" w:color="00FF00"/></w:tcBorders>'
		const result = await layout({ body: table(`<w:tr>${cell("a", red)}${cell("b", thick)}${cell("c")}</w:tr>`) })
		const shared = graphics(result, 0, PageGraphic.BORDER).filter(
			line => line.height > 2 && Math.abs(line.x + line.width / 2 - 288) < 2
		)
		const values = result.decorations

		expect(shared.length).toBe(1)
		// double (number 3) outweighs single (1) at the same width.
		expect(values.some(value => value.kind === "border" && value.style === "double")).toBe(true)
		// Three strokes of half a point each.
		expect(shared[0]?.width).toBeCloseTo(2, 1)
	})

	it("moves a row that cannot split", async () => {
		const row = `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell("")}${cell("")}<w:tc>${lines(10)}</w:tc></w:tr>`
		const result = await layout({ body: lines(50) + table(row) })

		expect(body(result, 0).length).toBe(50)
		expect(body(result, 1)[2]?.text).toBe("line 1")
	})

	it("lays out nested tables inside cells", async () => {
		const inner = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid><w:tr>${cell("x")}${cell("y")}</w:tr></w:tbl>`
		const result = await layout({ body: table(`<w:tr><w:tc>${inner}${p("")}</w:tc>${cell("b")}${cell("c")}</w:tr>`) })
		const all = body(result, 0)

		expect(texts(all)).toEqual(["x", "y", "", "b", "c"])
		expect(all[1]?.x).toBe(96 + 96)
		expect(all[2]?.top).toBeCloseTo(96 + LINE, 1)
	})

	it("offsets a floating table horizontally", async () => {
		const result = await layout({
			body: table(`<w:tr>${cell("a")}${cell("b")}${cell("c")}</w:tr>`, '<w:tblpPr w:horzAnchor="margin" w:tblpX="720"/>')
		})

		expect(body(result, 0)[0]?.x).toBe(96 + 48)
	})

	it("starts a table with a page break before its first paragraph on a new page", async () => {
		const result = await layout({
			body:
				p("before") +
				table(
					`<w:tr><w:tc><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>cell</w:t></w:r></w:p></w:tc>${cell("")}${cell("")}</w:tr>`
				)
		})

		expect(texts(body(result, 1))[0]).toBe("cell")
	})
})

describe("notes", () => {
	const footnotes =
		'<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
		'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
		'<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> first note</w:t></w:r></w:p></w:footnote>' +
		`<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> second note</w:t></w:r></w:p>${lines(3)}</w:footnote>`
	const reference = (id: number): string => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`

	it("numbers references and places notes at the bottom of their page", async () => {
		const result = await layout({
			body: `<w:p><w:r><w:t>text</w:t></w:r>${reference(1)}<w:r><w:t xml:space="preserve"> more</w:t></w:r>${reference(2)}</w:p>`,
			footnotes,
			settings: '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
		})
		const notes = pageLines(result, 0).filter(line => line.region === Region.FOOTNOTE)

		expect(texts(body(result, 0))).toEqual(["text1 more2"])
		expect(notes.map(line => line.text)).toEqual(["￼", "1 first note", "2 second note", "line 1", "line 2", "line 3"])
		expect(notes[notes.length - 1]?.bottom).toBeCloseTo(960, 1)
	})

	it("moves a line whose note does not fit to the next page", async () => {
		const result = await layout({
			body: lines(54) + `<w:p><w:r><w:t>ref</w:t></w:r>${reference(2)}</w:p>`,
			footnotes,
			settings: '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
		})

		expect(body(result, 0).length).toBe(54)
		expect(texts(body(result, 1))).toEqual(["ref1"])
		expect(pageLines(result, 1).filter(line => line.region === Region.FOOTNOTE).length).toBe(5)
	})

	it("continues a note that does not fit on the next page", async () => {
		const long = `<w:footnote w:id="3"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> long</w:t></w:r></w:p>${lines(80)}</w:footnote>`
		const result = await layout({
			body: `<w:p><w:r><w:t>ref</w:t></w:r>${reference(3)}</w:p>` + lines(10),
			footnotes: footnotes + long,
			settings: '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
		})
		const first = pageLines(result, 0).filter(line => line.region === Region.FOOTNOTE)
		const second = pageLines(result, 1).filter(line => line.region === Region.FOOTNOTE)

		expect(result.pages.length).toBe(2)
		expect(first.length + second.length).toBe(83)
		expect(first.length).toBeGreaterThan(20)
		expect(second[0]?.text).toBe("\ufffc")
		expect(Math.max(...first.map(line => line.bottom))).toBeLessThanOrEqual(960 + 0.01)
	})

	it("places endnotes at the end of their section when asked to", async () => {
		const sectionBreak =
			'<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>'
		const result = await layout({
			body: `<w:p><w:r><w:t>one</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>${sectionBreak}${p("two")}`,
			endnotes:
				'<w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> end of section</w:t></w:r></w:p></w:endnote>',
			settings: '<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr>'
		})

		expect(
			pageLines(result, 0)
				.filter(line => line.region === Region.ENDNOTE)
				.map(line => line.text)
		).toEqual(["i end of section"])
		expect(texts(body(result, 1))).toEqual(["two"])
	})

	it("restarts footnote numbers on every page where asked", async () => {
		const result = await layout({
			body: `<w:p><w:r><w:t>one</w:t></w:r>${reference(1)}</w:p>` + lines(56) + `<w:p><w:r><w:t>two</w:t></w:r>${reference(2)}</w:p>`,
			footnotes,
			settings: '<w:footnotePr><w:numRestart w:val="eachPage"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
		})

		expect(texts(body(result, 0))[0]).toBe("one1")
		expect(texts(body(result, 1)).at(-1)).toBe("two1")
	})

	it("collects endnotes at the end of the document", async () => {
		const result = await layout({
			body: '<w:p><w:r><w:t>text</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>',
			endnotes:
				'<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
				'<w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> the note</w:t></w:r></w:p></w:endnote>',
			settings: '<w:endnotePr><w:endnote w:id="-1"/></w:endnotePr>'
		})
		const notes = pageLines(result, 0).filter(line => line.region === Region.ENDNOTE)

		expect(texts(body(result, 0))).toEqual(["texti"])
		expect(notes.map(line => line.text)).toEqual(["￼", "i the note"])
	})
})

describe("objects", () => {
	const inline = (cx: number, cy: number): string =>
		`<w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:inline></w:drawing></w:r>`
	const anchor = (wrap: string, x: number, y: number, cx: number, cy: number): string =>
		'<w:r><w:drawing><wp:anchor behindDoc="0" allowOverlap="1" layoutInCell="1" relativeHeight="1" simplePos="0" locked="0"><wp:simplePos x="0" y="0"/>' +
		`<wp:positionH relativeFrom="column"><wp:posOffset>${x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="${cx}" cy="${cy}"/>${wrap}<wp:docPr id="2" name="f"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:anchor></w:drawing></w:r>`

	it("sizes lines around inline pictures", async () => {
		const result = await layout({ body: `<w:p>${inline(952500, 952500)}<w:r><w:t>caption</w:t></w:r></w:p>` })
		const line = body(result, 0)[0]

		expect((line?.bottom ?? 0) - (line?.top ?? 0)).toBeGreaterThan(100)
		expect(result.objects.count).toBe(1)
	})

	it("wraps text beside a square-wrapped picture and below a top-and-bottom one", async () => {
		const square = await layout({
			body: `<w:p>${anchor('<wp:wrapSquare wrapText="right"/>', 0, 0, 2857500, 952500)}<w:r><w:t>${words(200)}</w:t></w:r></w:p>`
		})
		const lines = body(square, 0)
		const page = square.pages[0]

		expect(page?.objectCount).toBe(1)
		expect(page?.objectFloats[OBJECT_FLOATS + 0] ?? page?.objectFloats[0]).toBe(96)
		expect(page?.objectInts[OBJECT_INTS * 0]).toBe(0)
		// Lines beside the picture start right of it; lines below it use the whole width again.
		expect(lines[0]?.x).toBeGreaterThanOrEqual(96 + 300)
		expect(lines[lines.length - 1]?.x).toBe(96)

		const band = await layout({
			body: `<w:p>${anchor("<wp:wrapTopAndBottom/>", 0, 0, 2857500, 952500)}<w:r><w:t>below</w:t></w:r></w:p>`
		})

		expect(body(band, 0)[0]?.top).toBeGreaterThanOrEqual(96 + 100)
	})

	it("lays out a text box's story inside its frame", async () => {
		const box =
			'<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" allowOverlap="1" layoutInCell="1" relativeHeight="1" simplePos="0" locked="0"><wp:simplePos x="0" y="0"/>' +
			'<wp:positionH relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>4762500</wp:posOffset></wp:positionV>' +
			'<wp:extent cx="1905000" cy="952500"/><wp:wrapNone/><wp:docPr id="3" name="t"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
			'<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>inside the box</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp>' +
			"</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>"
		const result = await layout({ body: `<w:p>${box}<w:r><w:t>body</w:t></w:r></w:p>` })
		const inside = pageLines(result, 0).find(line => line.region === Region.TEXTBOX)

		expect(texts(body(result, 0))).toEqual(["body"])
		expect(inside?.text).toBe("inside the box")
		expect(inside?.x).toBe(100)
		expect(inside?.top).toBe(500)
	})

	it("sizes VML pictures from their style", async () => {
		const result = await layout({
			body: '<w:p><w:r><w:pict><v:shape style="width:72pt;height:144pt" type="#_x0000_t75"><v:imagedata r:id="rId9"/></v:shape></w:pict></w:r></w:p>'
		})

		expect(result.objects.widths[0]).toBe(96)
		expect(result.objects.heights[0]).toBe(192)
	})

	it("sets a drop cap beside the following lines", async () => {
		const result = await layout({
			body:
				p(
					"D",
					'<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>',
					'<w:sz w:val="72"/>'
				) + p(words(120))
		})
		const all = body(result, 0)

		expect(all[0]?.text).toBe("D")
		expect(all[1]?.top).toBe(all[0]?.top)
		// Lines beside the letter start right of it, later lines at the margin.
		expect(all[1]?.x).toBeGreaterThan(96 + 20)
		expect(all[all.length - 1]?.x).toBe(96)
	})
})
