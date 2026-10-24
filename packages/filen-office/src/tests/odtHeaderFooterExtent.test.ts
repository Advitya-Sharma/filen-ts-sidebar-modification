import { describe, expect, it } from "vitest"
import { TextSelection } from "prosemirror-state"
import type { DocxEditorCommand } from "@filen/office/core/docx/edit"
import { pageLines } from "./layoutFixtures"
import { odtBytes, odtLayout, partText } from "./odtFixtures"
import { editOdt, saveOdt } from "./odtWriteFixtures"
import { paragraphs } from "./rtfWriteFixtures"
import { bytesSource } from "@filen/office/core/opc"
import { translateOdt } from "../core/odt/index"

// Expected positions are LibreOffice 7.5's PDF export of the same documents (A4, 2 cm margins), in twips.
const STYLES =
	'<style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0in" fo:margin-bottom="0in"/><style:text-properties fo:font-family="&apos;Liberation Sans&apos;" fo:font-size="10pt"/></style:default-style>'
const BODY = Array.from({ length: 80 }, (_, index) => `<text:p>Body line ${index}</text:p>`).join("")

async function bodyLines(kind: "header" | "footer", properties: string, lines: number): Promise<{ count: number; top: number }> {
	const region = Array.from({ length: lines }, (_, index) => `<text:p>${kind} line ${index}</text:p>`).join("")
	const { layout } = await odtLayout(
		odtBytes({
			styles: STYLES,
			body: BODY,
			stylesAutomatic: `<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/><style:${kind}-style><style:header-footer-properties ${properties}/></style:${kind}-style></style:page-layout>`,
			masters: `<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:${kind}>${region}</style:${kind}></style:master-page>`
		})
	)
	const body = pageLines(layout, 0).filter(line => line.text.startsWith("Body"))

	return { count: body.length, top: (body[0]?.top ?? 0) * 15 }
}

describe("an ODT header or footer's minimum height", () => {
	it.each([
		["holds its spacing", 'fo:min-height="2in" fo:margin-top="0.2in"', 1, 50],
		["grows with content past it", 'fo:min-height="0.5in" fo:margin-top="0.3in"', 2, 59],
		["of zero is the spacing and content", 'fo:min-height="0in" fo:margin-top="0.3in"', 1, 60],
		["given as svg:height holds its spacing too", 'svg:height="2in" fo:margin-top="0.2in"', 1, 50]
	])("of a footer %s", async (_, properties, lines, expected) => {
		expect((await bodyLines("footer", properties, lines)).count).toBe(expected)
	})

	it("of a header holds its spacing above the body", async () => {
		expect((await bodyLines("header", 'fo:min-height="2in" fo:margin-bottom="0.2in"', 1)).top).toBeCloseTo(1134 + 2880, -1)
	})

	it("leaves the last line's proportional spacing out of a header and a footer", async () => {
		const spaced =
			'<style:style style:name="H" style:family="paragraph"><style:paragraph-properties fo:margin-top="0in" fo:margin-bottom="0in" fo:line-height="150%"/><style:text-properties fo:font-size="11pt"/></style:style>'
		const { layout } = await odtLayout(
			odtBytes({
				styles: STYLES,
				body: "<text:p>Body</text:p>",
				stylesAutomatic: `${spaced}<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/><style:header-style><style:header-footer-properties fo:min-height="0.1in"/></style:header-style><style:footer-style><style:header-footer-properties fo:min-height="0.1in"/></style:footer-style></style:page-layout>`,
				masters:
					'<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p text:style-name="H">Head</text:p></style:header><style:footer><text:p text:style-name="H">Foot</text:p></style:footer></style:master-page>'
			})
		)
		const page = layout.pages[0]
		const baseline = (text: string): number => {
			const index = pageLines(layout, 0).findIndex(line => line.text.startsWith(text))

			return (page?.text.lineFloats[index * 5 + 2] ?? 0) * 15
		}

		// 11 pt Liberation Sans: the 10 pt body line starts right under the header's 253 tw line, as with single spacing.
		expect(baseline("Body") - baseline("Head")).toBeCloseTo(234, -1)
		expect(baseline("Foot")).toBeCloseTo(15650, -1)
	})

	it.each([
		["above its spacing", 2268, "0.25in"],
		["below its spacing", 1300, "0.25in"]
	])("survives a page margin edit %s exactly", async (_, top, spacing) => {
		const bytes = odtBytes({
			styles: STYLES,
			body: "<text:p>One</text:p>",
			stylesAutomatic: `<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/><style:header-style><style:header-footer-properties fo:min-height="0.5in" fo:margin-bottom="${spacing}"/></style:header-style></style:page-layout>`,
			masters:
				'<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Head</text:p></style:header></style:master-page>'
		})
		const editing = await editOdt(bytes)
		const first = paragraphs(editing.session.state.doc)[0]

		if (first === undefined) throw new Error("no paragraph")

		editing.session.dispatch(editing.session.state.tr.setSelection(TextSelection.create(editing.session.state.doc, first.pos + 1)))
		expect(await editing.editor.execute({ type: "page.setup", margins: { top } } as DocxEditorCommand)).toBe(true)

		const reopened = await translateOdt(bytesSource((await saveOdt(editing)).bytes))

		if (!reopened.ok) throw new Error(reopened.message)

		expect(partText(reopened.translation, "word/document.xml")).toContain(`w:top="${top}"`)
	})
})
