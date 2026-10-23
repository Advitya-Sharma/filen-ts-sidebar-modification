import { describe, expect, it } from "vitest"
import { Region, type DocumentLayout } from "../core/docx/layout/index"
import { buildDocx } from "./docxSynthetic"
import { layOut, pageLines, type LineInfo } from "./layoutFixtures"

// Text beside a square-wrapped float (ECMA-376 §20.4.2.17 wrapSquare) goes where it fits; with no room it moves below
// the float. The spec names no least width; LibreOffice 7.5 (the tie-break), opening these packages, puts a line
// beside a float when its first word fits there, and an empty paragraph's mark in any gap at all, and moves both
// below a float that leaves none.

const SECTION =
	'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'
// The column: A4 less two 1in margins, in EMU.
const COLUMN = 5760720

// A text box right-aligned in the margins, `gap` EMU narrower than the column, 1.535in tall, its top `dy` EMU below
// the top of its paragraph.
function box(gap: number, dy: number): string {
	const cx = COLUMN - gap

	return (
		'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
		`<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${dy}</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="${cx}" cy="1403988"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="9" name="Text Box 5"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
		`<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="1403988"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
		"<wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
	)
}

// "Your email", its box starting 200000 EMU (15.7 pt) below, so the next line meets it.
async function form(gap: number, after: string): Promise<{ layout: DocumentLayout; lines: LineInfo[]; boxTop: number; boxBottom: number }> {
	const layout = await layOut(
		buildDocx({ body: `<w:p>${box(gap, 200000)}<w:r><w:t>Your email</w:t></w:r></w:p><w:p/>${after}${SECTION}` })
	)
	const page = layout.pages[0]
	const boxTop = page?.objectFloats[1] ?? 0

	return {
		layout,
		lines: pageLines(layout, 0).filter(line => line.region === Region.BODY),
		boxTop,
		boxBottom: boxTop + (page?.objectFloats[3] ?? 0)
	}
}

const EMU_PER_CM = 360000

describe("text beside square-wrapped floats", () => {
	it("moves an empty paragraph below a float that leaves no room beside it", async () => {
		const { lines, boxTop, boxBottom } = await form(0, "")
		const empty = lines[1]

		expect(lines[0]?.text).toBe("Your email")
		expect(empty?.text).toBe("")
		// Without the move it would overlap the box's top.
		expect(lines[0]?.bottom ?? 0).toBeLessThan(boxTop)
		expect(empty?.top).toBeCloseTo(boxBottom, 1)
	})

	it("keeps an empty paragraph beside a float that leaves a gap, in the gap", async () => {
		const { lines, boxTop } = await form(Math.round(0.2 * EMU_PER_CM), "")
		const empty = lines[1]

		expect(empty?.text).toBe("")
		expect(empty?.bottom ?? 0).toBeGreaterThan(boxTop)
		expect(empty?.top ?? 0).toBeLessThan(boxTop)
		expect(empty?.x).toBeCloseTo(96, 1)
	})

	it("wraps words beside a float where the first fits, and moves them below where it does not", async () => {
		const words = "<w:p><w:r><w:t>Afterwards the form continues</w:t></w:r></w:p>"
		const narrow = await form(Math.round(0.2 * EMU_PER_CM), words)
		const wide = await form(Math.round(2.5 * EMU_PER_CM), words)
		const first = (result: { lines: LineInfo[] }): LineInfo | undefined => result.lines.find(line => line.text.startsWith("After"))

		expect(first(narrow)?.top).toBeCloseTo(narrow.boxBottom, 1)
		expect(first(wide)?.top ?? Infinity).toBeLessThan(wide.boxBottom)
	})
})
