import { describe, expect, it } from "vitest"
import { Region, type DocumentLayout } from "../core/docx/layout/index"
import { buildDocx } from "./docxSynthetic"
import { layOut, pageLines } from "./layoutFixtures"
import { partText, translated } from "./odtFixtures"

// ODF 1.3 Part 3 §20.267 style:flow-with-text: a drawing shape that does not flow with the text "can be laid out
// anywhere on the page its anchor is on", so it stays on that page. LibreOffice 7.5 (measured over UNO on copies of
// the NSI notification forms, which leave the attribute out, read as false) moves a paragraph-anchored frame that
// would cross the bottom edge of the page up until it does not; where that would take it above its paragraph, the
// paragraph moves to the next page with it. WordprocessingML has no such rule (Word lets the object run off the
// page), so the translation asks for it in a compatibility setting the layout honours only there.

const FLOW_SETTING = '<w:compatSetting w:name="flow-with-text" w:uri="urn:oasis:names:tc:opendocument:xmlns:style:1.0" w:val="false"/>'

const FRAME_STYLE = (flow: string): string =>
	`<style:style style:name="fr" style:family="graphic"><style:graphic-properties style:wrap="parallel"${flow} style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="left" style:horizontal-rel="paragraph"/></style:style>`
const FRAME_BODY =
	'<text:p>Anchor<draw:frame draw:style-name="fr" text:anchor-type="paragraph" svg:y="0.5cm" svg:width="5cm" svg:height="2cm"><draw:text-box><text:p/></draw:text-box></draw:frame></text:p>'

describe("ODT frames kept on their anchor's page", () => {
	it("asks for frames to stay on their anchor's page unless the document turns capture off", async () => {
		const settings = (translation: Awaited<ReturnType<typeof translated>>): string => partText(translation, "word/settings.xml")

		expect(settings(await translated({ automatic: FRAME_STYLE(""), body: FRAME_BODY }))).toContain(FLOW_SETTING)
		expect(
			settings(
				await translated({
					automatic: FRAME_STYLE(""),
					body: FRAME_BODY,
					settings: '<config:config-item config:name="DoNotCaptureDrawObjsOnPage" config:type="boolean">true</config:config-item>'
				})
			)
		).not.toContain("flow-with-text")
	})

	it("counts a frame that flows with the text as approximated: it is kept on the page, not in its text area", async () => {
		const translation = await translated({ automatic: FRAME_STYLE(' style:flow-with-text="true"'), body: FRAME_BODY })

		expect(translation.lossy.get("frameAnchor")).toBe(1)
	})
})

const SECTION =
	'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'

// A 4in by `height` EMU shape 0.5in below the top of its paragraph.
function frame(height: number): string {
	return (
		'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
		'<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>' +
		`<wp:extent cx="3657600" cy="${height}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="9" name="f"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
		`<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
	)
}

// Lines fill the page to near its bottom, then the frame's paragraph; `flow` adds the ODF setting.
async function nearBottom(
	height: number,
	flow: boolean
): Promise<{ layout: DocumentLayout; anchorPage: number; frameTop: number; frameBottom: number }> {
	const filler = Array.from({ length: 58 }, (_, k) => `<w:p><w:r><w:t>line ${k + 1}</w:t></w:r></w:p>`).join("")
	const layout = await layOut(
		buildDocx({
			body: `${filler}<w:p>${frame(height)}<w:r><w:t>Anchor</w:t></w:r></w:p>${SECTION}`,
			...(flow ? { settings: `<w:compat>${FLOW_SETTING}</w:compat>` } : {})
		})
	)
	const anchorPage = layout.pages.findIndex((_, k) =>
		pageLines(layout, k).some(line => line.region === Region.BODY && line.text === "Anchor")
	)
	const page = layout.pages[anchorPage]
	const frameTop = page?.objectFloats[1] ?? NaN

	return { layout, anchorPage, frameTop, frameBottom: frameTop + (page?.objectFloats[3] ?? NaN) }
}

describe("layout of floats kept on their anchor's page", () => {
	it("leaves a WordprocessingML document's float running off the page", async () => {
		const result = await nearBottom(1828800, false)

		expect(result.anchorPage).toBe(0)
		expect(result.frameBottom).toBeGreaterThan(1122.5)
	})

	it("moves the float up into the page where it fits below its paragraph's top", async () => {
		const result = await nearBottom(914400, true)
		const anchor = pageLines(result.layout, result.anchorPage).find(line => line.text === "Anchor")

		expect(result.anchorPage).toBe(0)
		expect(result.frameBottom).toBeCloseTo(1122.5, 0)
		expect(result.frameTop).toBeGreaterThanOrEqual(anchor?.top ?? Infinity)
	})

	it("moves the paragraph to the next page where the float would have to rise above it", async () => {
		const result = await nearBottom(1828800, true)
		const anchor = pageLines(result.layout, result.anchorPage).find(line => line.text === "Anchor")

		expect(result.anchorPage).toBe(1)
		expect(anchor?.top).toBeCloseTo(96, 0)
		expect(result.frameTop).toBeCloseTo(96 + 48, 0)
	})
})
