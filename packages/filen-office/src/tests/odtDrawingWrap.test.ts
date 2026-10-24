import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { convertText } from "../core/convert/text/index"
import { bytesSource, ZipArchive } from "../core/opc/index"
import { decodeUtf8 } from "../core/xml/index"
import { pageLines, type LineInfo } from "./layoutFixtures"
import { CORPUS_DIR } from "./corpusDir"
import { layOutOdt, odtBytes, odtLayout, partText, translated, type OdtParts } from "./odtFixtures"

// Expected distances are LibreOffice 7.5's PDF export of the same documents, in points.
const PX = 0.75
const STYLES =
	'<style:default-style style:family="paragraph"><style:text-properties style:font-name="Liberation Sans" fo:font-size="12pt"/></style:default-style>'
const LETTER = join(CORPUS_DIR, "odt", "govuk__UKHSA_MMRV_catch-up_letter_template_01_FINAL_.odt")
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

async function lines(parts: OdtParts): Promise<LineInfo[]> {
	const { layout } = await odtLayout(odtBytes({ styles: STYLES, ...parts }))

	return pageLines(layout, 0).filter(line => line.region === 0)
}

function top(all: readonly LineInfo[], text: string): number {
	const found = all.find(item => item.text.includes(text))

	if (found === undefined) throw new Error(`no line ${text}`)

	return found.top * PX
}

describe("text wrapped around an ODT drawing shape", () => {
	const shape = (path: string, mirror: boolean): OdtParts => ({
		automatic:
			'<style:style style:family="graphic" style:name="a1"><style:graphic-properties style:wrap="none" draw:fill="none" draw:stroke="solid" svg:stroke-width="0.007in" style:horizontal-rel="page" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:vertical-pos="from-top"/></style:style>',
		body: `<text:p>Anchor<draw:custom-shape svg:x="1in" svg:y="0.3in" svg:width="5in" svg:height="0.5in" draw:style-name="a1" text:anchor-type="paragraph"><draw:enhanced-geometry svg:viewBox="0 0 100 100" draw:enhanced-path="${path}"${mirror ? ' draw:mirror-vertical="true"' : ""}/></draw:custom-shape></text:p>${[0, 1, 2, 3].map(index => `<text:p>Line ${index}</text:p>`).join("")}`
	})

	it.each([
		["a line at the frame's top", "M 0 0 L 100 0 N", false, [21.9, 35.7, 49.5]],
		["a line mirrored to the frame's bottom", "M 0 0 L 100 0 N", true, [13.8, 27.6, 41.4]],
		["a line across the frame's middle", "M 0 50 L 100 50 N", false, [13.8, 39.9, 53.7]],
		["a rectangle filling the frame", "M 0 0 L 100 0 100 100 0 100 Z N", false, [58, 71.8, 85.6]]
	])("follows the drawn path and its stroke, not the frame: %s", async (_, path, mirror, expected) => {
		const all = await lines(shape(path, mirror))

		expect(expected.map((_, index) => top(all, `Line ${index}`) - top(all, "Anchor"))).toEqual(
			expected.map(value => expect.closeTo(value, 0))
		)
	})
})

describe("an as-char draw:g", () => {
	const group = (members: string): OdtParts => ({
		automatic: '<style:style style:family="graphic" style:name="g"><style:graphic-properties/></style:style>',
		body: `<text:p>Before</text:p><text:p><draw:g draw:style-name="g" text:anchor-type="as-char">${members}</draw:g></text:p><text:p>After</text:p>`,
		extra: [{ name: "Pictures/a.png", data: PNG }]
	})
	const picture = (x: string, y: string): string =>
		`<draw:frame svg:x="${x}" svg:y="${y}" svg:width="0.5in" svg:height="0.2in"><draw:image xlink:href="Pictures/a.png"/></draw:frame>`

	it("is one character as tall as its members, with no line below it", async () => {
		const all = await lines(group(picture("0in", "0in") + picture("0in", "0.3in")))

		expect(top(all, "After") - top(all, "Before")).toBeCloseTo(49.9, 0)
	})

	it("keeps a drawing in a member's text in that text", async () => {
		const box = `<draw:frame svg:x="0in" svg:y="0in" svg:width="1in" svg:height="0.5in"><draw:text-box><text:p>${picture("0in", "0in").replace("<draw:frame ", '<draw:frame text:anchor-type="as-char" ')}</text:p></draw:text-box></draw:frame>`
		const document = partText(await translated(group(box + picture("0in", "0.6in"))), "word/document.xml")
		const text = /<w:txbxContent>(.*?)<\/w:txbxContent>/su.exec(document)?.[1] ?? ""

		expect(text).toContain("<pic:pic>")
		expect(document.match(/<wp:inline/gu)?.length).toBe(2)
	})

	it("is written as one inline group of its members", async () => {
		const document = partText(await translated(group(picture("0.1in", "0in") + picture("0in", "0.3in"))), "word/document.xml")

		expect(document.match(/<wp:inline/gu)?.length).toBe(1)
		expect(document).not.toContain("<wp:anchor")
		expect(document).toContain('<wp:extent cx="548640" cy="457200"/>')
		expect(document.match(/<pic:pic>/gu)?.length).toBe(2)
	})
})

describe("a line of as-char objects", () => {
	const line = (content: string, height = "0.5in"): OdtParts => ({
		automatic: '<style:style style:family="graphic" style:name="g"><style:graphic-properties/></style:style>',
		body: `<text:p>Before</text:p><text:p>${content.replace("FRAME", `<draw:frame draw:style-name="g" text:anchor-type="as-char" svg:width="0.5in" svg:height="${height}"><draw:image xlink:href="Pictures/a.png"/></draw:frame>`)}</text:p><text:p>After</text:p>`,
		extra: [{ name: "Pictures/a.png", data: PNG }]
	})

	it.each([
		["a picture alone is as tall as the picture", "FRAME", "0.5in", 49.9],
		["a small picture alone is not raised to the font's height", "FRAME", "0.05in", 17.5]
	])("%s", async (_, content, height, expected) => {
		const all = await lines(line(content, height))

		expect(top(all, "After") - top(all, "Before")).toBeCloseTo(expected, 0)
	})

	it("a picture beside text keeps the text's descent below it (LibreOffice: 53 pt)", async () => {
		const all = await lines(line("text FRAME text"))

		expect(top(all, "After") - top(all, "Before")).toBeGreaterThan(49.9 + 2)
	})

	it("is left out of a converted .docx, which counts it where objects are inline", async () => {
		const converted = await convertText(odtBytes({ styles: STYLES, ...line("FRAME") }), "odt", "docx")

		if (!converted.ok) throw new Error(converted.refusal.message)

		const archive = await ZipArchive.open(bytesSource(converted.bytes))
		const settings = archive.get("word/settings.xml")

		expect(settings === undefined ? "" : decodeUtf8(await archive.read(settings))).not.toContain("object-line-height")
		expect(converted.lossy.features.map(feature => feature.kind)).toContain("objectLineHeight")
	})

	it("is asked for by the translation's settings", async () => {
		expect(partText(await translated(line("FRAME")), "word/settings.xml")).toContain(
			'<w:compatSetting w:name="object-line-height" w:uri="urn:oasis:names:tc:opendocument:xmlns:style:1.0" w:val="objects"/>'
		)
	})
})

// A letter with an as-char logo group and a rule drawn at the foot of its frame: one page in LibreOffice.
it.skipIf(!existsSync(LETTER))("lays out a letter with an as-char group and a drawn rule on LibreOffice's one page", async () => {
	expect((await layOutOdt(new Uint8Array(readFileSync(LETTER)))).pages).toBe(1)
})
