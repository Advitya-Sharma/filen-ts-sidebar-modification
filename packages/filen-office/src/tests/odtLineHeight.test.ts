import { describe, expect, it } from "vitest"
import { LINE_FLOATS, LineFloat } from "../core/text/index"
import { pageLines } from "./layoutFixtures"
import { odtBytes, odtLayout, ODF_NS, type OdtParts } from "./odtFixtures"
import { text } from "./zipBuilder"

// Expected distances are LibreOffice 7.5's PDF export of the same documents, in points.
const PX = 0.75
const STYLES =
	'<style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0in" fo:margin-bottom="0in"/><style:text-properties fo:font-family="&apos;Liberation Sans&apos;" fo:font-size="10pt"/></style:default-style><style:style style:name="Mono" style:family="text"><style:text-properties fo:font-family="&apos;Liberation Mono&apos;" fo:font-size="10pt"/></style:style>'
const PLAIN = `<text:p>${["First", "Second", "Third"].map(word => `${word} line of plain text`).join("<text:line-break/>")}</text:p>`
const MIXED = `<text:p>${["First", "Second", "Third"].map(word => `${word} line with <text:span text:style-name="Mono">mono</text:span> text`).join("<text:line-break/>")}</text:p>`

// settings.xml with a configuration-settings set of these items, its prefix bound as LibreOffice reads it.
function settings(items: Readonly<Record<string, string>>): OdtParts["extra"] {
	const entries = Object.entries(items)
		.map(([name, value]) => {
			const type = value === "true" || value === "false" ? "boolean" : "string"

			return `<config:config-item config:name="${name}" config:type="${type}">${value}</config:config-item>`
		})
		.join("")

	return [
		{
			name: "settings.xml",
			data: text(
				`<?xml version="1.0" encoding="UTF-8"?><office:document-settings ${ODF_NS} xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:configuration-settings">${entries}</config:config-item-set></office:settings></office:document-settings>`
			)
		}
	]
}

interface Line {
	readonly top: number
	readonly baseline: number
}

async function lines(body: string, extra?: OdtParts["extra"], automatic?: string): Promise<Line[]> {
	const { layout } = await odtLayout(
		odtBytes({ styles: STYLES, body, ...(extra === undefined ? {} : { extra }), ...(automatic === undefined ? {} : { automatic }) })
	)
	const page = layout.pages[0]

	if (page === undefined) throw new Error("no page")

	return pageLines(layout, 0).map((line, index) => ({
		top: line.top * PX,
		baseline: (page.text.lineFloats[index * LINE_FLOATS + LineFloat.BASELINE] ?? 0) * PX
	}))
}

function pitch(all: readonly Line[]): number {
	const [first, second, third] = all

	if (first === undefined || second === undefined || third === undefined) throw new Error("fewer than three lines")

	expect(third.top - second.top).toBeCloseTo(second.top - first.top, 3)

	return second.top - first.top
}

describe("the height of an ODT line", () => {
	it("adds the font's line gap under the text of a document without settings", async () => {
		const all = await lines(PLAIN)

		expect(pitch(all)).toBeCloseTo(11.5, 1)

		const [first] = all

		expect((first?.baseline ?? 0) - (first?.top ?? 0)).toBeCloseTo(9.05, 1)
	})

	it.each([
		["AddExternalLeading missing", {}],
		["AddExternalLeading false", { AddExternalLeading: "false" }],
		["the zero leading of the high-resolution layout", { AddExternalLeading: "true", PrinterIndependentLayout: "high-resolution" }]
	])("leaves the line gap out with %s", async (_, items) => {
		const all = await lines(PLAIN, settings(items))

		expect(pitch(all)).toBeCloseTo(11.17, 1)

		const [first] = all

		expect((first?.baseline ?? 0) - (first?.top ?? 0)).toBeCloseTo(9.05, 1)
	})

	it.each([
		["printer metrics", { AddExternalLeading: "true", PrinterIndependentLayout: "disabled" }],
		["the low-resolution layout", { AddExternalLeading: "true", PrinterIndependentLayout: "low-resolution" }],
		[
			"UnxForceZeroExtLeading false",
			{ AddExternalLeading: "true", UnxForceZeroExtLeading: "false", PrinterIndependentLayout: "high-resolution" }
		]
	])("keeps the line gap with AddExternalLeading under %s", async (_, items) => {
		expect(pitch(await lines(PLAIN, settings(items)))).toBeCloseTo(11.5, 1)
	})

	it.each([
		["with", undefined, 13.43],
		["without", settings({ AddExternalLeading: "false" }), 11]
	])("in Carlito takes Calibri's ascent and descent %s the line gap", async (_, extra, expected) => {
		const all = await lines(
			`<text:p text:style-name="Carlito">${["First", "Second", "Third"].map(word => `${word} line`).join("<text:line-break/>")}</text:p>`,
			extra,
			'<style:style style:name="Carlito" style:family="paragraph"><style:text-properties fo:font-family="Carlito" fo:font-size="11pt"/></style:style>'
		)

		expect(pitch(all)).toBeCloseTo(expected, 1)
		expect((all[0]?.baseline ?? 0) - (all[0]?.top ?? 0)).toBeCloseTo(8.25, 1)
	})

	it.each([
		["with", undefined],
		["without", settings({ AddExternalLeading: "false" })]
	])("of mixed fonts is the tallest ascent over the deepest descent %s the line gap", async (_, extra) => {
		expect(pitch(await lines(MIXED, extra))).toBeCloseTo(12.05, 1)
	})

	it("leaves out the paragraph mark's font when the line has text", async () => {
		const all = await lines(
			["First", "Second", "Third"]
				.map(word => `<text:p text:style-name="Big"><text:span text:style-name="Small">${word}</text:span></text:p>`)
				.join(""),
			undefined,
			'<style:style style:name="Big" style:family="paragraph"><style:text-properties fo:font-family="Carlito" fo:font-size="20pt"/></style:style><style:style style:name="Small" style:family="text"><style:text-properties fo:font-family="&apos;Liberation Sans&apos;" fo:font-size="11pt"/></style:style>'
		)

		expect(pitch(all)).toBeCloseTo(12.65, 1)
	})

	it.each([
		["Symbol", "Symbol", "&#xF0B7;"],
		["Segoe UI Symbol", "&apos;Segoe UI Symbol&apos;", "&#x2610;"]
	])("measures a missing symbol font (%s) as LibreOffice's symbol font, not the face drawing its characters", async (_, family, mark) => {
		const all = await lines(
			`<text:p text:style-name="Carlito">${["First", "Second", "Third"].map(word => `<text:span text:style-name="Symbol">${mark}</text:span> ${word}`).join("<text:line-break/>")}</text:p>`,
			undefined,
			`<style:style style:name="Carlito" style:family="paragraph"><style:text-properties fo:font-family="Carlito" fo:font-size="11pt"/></style:style><style:style style:name="Symbol" style:family="text"><style:text-properties fo:font-family="${family}" fo:font-size="11pt"/></style:style>`
		)

		expect(pitch(all)).toBeCloseTo(13.43, 1)
	})
})

describe("the spacing of an ODT line", () => {
	it.each([
		["a proportional height's extra space under the text", 'fo:line-height="150%"', 17.25, 9.05],
		["the baseline of an exact height at four fifths of it", 'fo:line-height="18pt"', 18, 14.4],
		["a minimum height's extra space above the text", 'style:line-height-at-least="16pt"', 16, 13.55]
	])("puts %s", async (_, properties, expectedPitch, expectedBaseline) => {
		const all = await lines(
			PLAIN.replace("<text:p>", '<text:p text:style-name="S">'),
			undefined,
			`<style:style style:name="S" style:family="paragraph"><style:paragraph-properties ${properties}/></style:style>`
		)

		expect(pitch(all)).toBeCloseTo(expectedPitch, 1)
		expect((all[0]?.baseline ?? 0) - (all[0]?.top ?? 0)).toBeCloseTo(expectedBaseline, 1)
	})
})
