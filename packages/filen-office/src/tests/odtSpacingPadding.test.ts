import { describe, expect, it } from "vitest"
import { convertText } from "../core/convert/text/index"
import { bytesSource, ZipArchive } from "../core/opc/index"
import { decodeUtf8 } from "../core/xml/index"
import { TextSelection } from "prosemirror-state"
import { editOdt, saveOdt } from "./odtWriteFixtures"
import { pageLines, type LineInfo } from "./layoutFixtures"
import { odtBytes, odtLayout, partText, translated, type OdtParts } from "./odtFixtures"

// Expected distances are LibreOffice 7.5's PDF export of the same documents, in points.
const PX = 0.75
const STYLES =
	'<style:default-style style:family="paragraph"><style:text-properties style:font-name="Liberation Sans" fo:font-size="12pt"/></style:default-style>'

async function lines(parts: OdtParts): Promise<LineInfo[]> {
	const { layout } = await odtLayout(odtBytes({ styles: STYLES, ...parts }))

	return pageLines(layout, 0).filter(line => line.region === 0)
}

function line(all: readonly LineInfo[], text: string): LineInfo {
	const found = all.find(item => item.text.includes(text))

	if (found === undefined) throw new Error(`no line ${text}`)

	return found
}

const pitch = (all: readonly LineInfo[], from: string, to: string): number => (line(all, to).top - line(all, from).top) * PX
const shift = (all: readonly LineInfo[], from: string, to: string): number => (line(all, to).x - line(all, from).x) * PX

describe("space between ODT paragraphs", () => {
	const spaced =
		'<style:style style:name="S" style:family="paragraph"><style:paragraph-properties fo:margin-top="0.25in" fo:margin-bottom="0.25in"/></style:style>'
	const body = [0, 1, 2, 3].map(index => `<text:p text:style-name="S">Line ${index}</text:p>`).join("")

	it.each([
		["no settings", undefined],
		[
			"AddParaTableSpacing false",
			'<config:config-item config:name="AddParaTableSpacing" config:type="boolean">false</config:config-item>'
		],
		[
			"AddParaTableSpacing true",
			'<config:config-item config:name="AddParaTableSpacing" config:type="boolean">true</config:config-item>'
		]
	])("is the sum of space after and space before (%s)", async (_, settings) => {
		const parts = { styles: STYLES + spaced, body, ...(settings === undefined ? {} : { settings }) }

		expect(pitch(await lines(parts), "Line 0", "Line 1")).toBeCloseTo(49.8, 0)
		expect(partText(await translated(parts), "word/settings.xml")).toContain("<w:doNotUseHTMLParagraphAutoSpacing/>")
	})
})

describe("fo:padding of an ODT paragraph", () => {
	const padded = (properties: string): OdtParts => ({
		automatic: `<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties ${properties}/></style:style>`,
		body: '<text:p>Before</text:p><text:p text:style-name="P1">Padded</text:p><text:p>After</text:p>'
	})

	it.each([
		["no border", 'fo:padding="0.5in"'],
		["fo:border none", 'fo:border="none" fo:padding="0.5in"'],
		["a 0in border", 'fo:border="0in solid #FFFFFF" fo:padding="0.5in"']
	])("sets the text in from its margins with %s", async (_, properties) => {
		const all = await lines(padded(properties))

		expect(pitch(all, "Before", "Padded")).toBeCloseTo(49.8, 0)
		expect(pitch(all, "Padded", "After")).toBeCloseTo(49.8, 0)
		expect(shift(all, "Before", "Padded")).toBeCloseTo(36, 0)
	})

	it("adds to the margins it is set with", async () => {
		const all = await lines(padded('fo:margin-top="0.25in" fo:margin-bottom="0.25in" fo:margin-left="0.3in" fo:padding="0.25in"'))

		expect(pitch(all, "Before", "Padded")).toBeCloseTo(49.8, 0)
		expect(pitch(all, "Padded", "After")).toBeCloseTo(49.8, 0)
		expect(shift(all, "Before", "Padded")).toBeCloseTo(39.6, 0)
	})

	it("keeps a border's width and its padding in front of the text", async () => {
		const thin = await lines(padded('fo:border="0.01in solid #000000" fo:padding="0.5in"'))

		expect(pitch(thin, "Before", "Padded")).toBeCloseTo(50.5, 0)
		expect(pitch(thin, "Padded", "After")).toBeCloseTo(50.5, 0)

		const left = await lines(padded('fo:padding-left="0.5in" fo:border-left="0.02in solid #000000"'))

		expect(shift(left, "Before", "Padded")).toBeCloseTo(37.4, 0)
	})

	it("is not cut to 31 pt beside a border", async () => {
		const all = await lines(padded('fo:border="0.01in solid #000000" fo:padding="0.8in"'))

		expect(pitch(all, "Before", "Padded")).toBeCloseTo(72.1, 0)
		expect(shift(all, "Before", "Padded")).toBeCloseTo(58.3, 0)
	})

	it("pads each side by its own border", async () => {
		const all = await lines(padded('fo:padding="0.25in" fo:border-top="0.02in solid #000000"'))

		expect(pitch(all, "Before", "Padded")).toBeCloseTo(33.3, 0)
		expect(pitch(all, "Padded", "After")).toBeCloseTo(31.8, 0)
		expect(shift(all, "Before", "Padded")).toBeCloseTo(18, 0)
	})

	it("moves a list paragraph's label and text in", async () => {
		const list =
			'<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="*"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.5in" fo:text-indent="-0.25in" fo:margin-left="0.5in"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>'
		const all = await lines({
			automatic: `${list}<style:style style:name="P1" style:family="paragraph" style:list-style-name="L1"><style:paragraph-properties fo:padding-left="0.5in"/></style:style><style:style style:name="P2" style:family="paragraph" style:list-style-name="L1"/>`,
			body: '<text:list text:style-name="L1"><text:list-item><text:p text:style-name="P2">Plain</text:p></text:list-item><text:list-item><text:p text:style-name="P1">Padded</text:p></text:list-item></text:list>'
		})

		expect(shift(all, "Plain", "Padded")).toBeCloseTo(36, 0)
	})

	describe("joined with a neighbour's", () => {
		const three = (first: string, second: string): OdtParts => ({
			automatic: `<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties ${first}/></style:style><style:style style:name="P2" style:family="paragraph"><style:paragraph-properties ${second}/></style:style>`,
			body: '<text:p>Before</text:p><text:p text:style-name="P1">One</text:p><text:p text:style-name="P2">Two</text:p><text:p>After</text:p>'
		})

		it("leaves out the padding between paragraphs that agree, keeping their margins", async () => {
			const bare = await lines(three('fo:padding="0.5in"', 'fo:padding="0.5in" fo:text-align="center"'))

			expect(pitch(bare, "Before", "One")).toBeCloseTo(49.8, 0)
			expect(pitch(bare, "One", "Two")).toBeCloseTo(13.8, 0)
			expect(pitch(bare, "Two", "After")).toBeCloseTo(49.8, 0)

			const spaced = 'fo:padding="0.5in" fo:margin-top="0.1in" fo:margin-bottom="0.1in"'
			const margins = await lines(three(spaced, spaced))

			expect(pitch(margins, "One", "Two")).toBeCloseTo(28.2, 0)
			expect(pitch(margins, "Two", "After")).toBeCloseTo(57, 0)
		})

		it("joins whatever the top and bottom padding when no line differs", async () => {
			const all = await lines(three('fo:padding="0.5in"', 'fo:padding="0.5in" fo:padding-top="0.2in" fo:padding-bottom="0.3in"'))

			expect(pitch(all, "One", "Two")).toBeCloseTo(13.8, 0)
			expect(pitch(all, "Two", "After")).toBeCloseTo(35.4, 0)
		})

		it("so top and bottom padding alone joins a plain paragraph's", async () => {
			const all = await lines({
				automatic:
					'<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:padding-top="0.5in" fo:padding-bottom="0.5in"/></style:style>',
				body: '<text:p>Before</text:p><text:p text:style-name="P1">One</text:p><text:p>After</text:p>'
			})

			expect(pitch(all, "Before", "One")).toBeCloseTo(13.8, 0)
			expect(pitch(all, "One", "After")).toBeCloseTo(13.8, 0)
		})

		it.each([
			["other left margins", 'fo:padding="0.5in" fo:margin-left="0.2in"', 'fo:padding="0.5in" fo:margin-left="0.3in"'],
			["another shadow", 'fo:padding="0.5in"', 'fo:padding="0.5in" style:shadow="#808080 0.05in 0.05in"'],
			["style:join-border false on the first", 'fo:padding="0.5in" style:join-border="false"', 'fo:padding="0.5in"']
		])("keeps it apart with %s", async (_, first, second) => {
			expect(pitch(await lines(three(first, second)), "One", "Two")).toBeCloseTo(85.8, 0)
		})

		it("keeps bordered paragraphs apart with style:join-border false on the first", async () => {
			const bordered = 'fo:border="0.02in solid #000000" fo:padding="0.2in"'
			const all = await lines(three(`${bordered} style:join-border="false"`, bordered))

			expect(pitch(all, "One", "Two")).toBeCloseTo(45.5, 0)
		})

		it("keeps padding to the twip, not to the whole point", async () => {
			const all = await lines(three('fo:padding="0.2in"', 'fo:padding="0.2in" fo:margin-left="0.5in"'))

			expect(pitch(all, "One", "Two")).toBeCloseTo(42.6, 1)
			expect(shift(all, "Before", "One")).toBeCloseTo(14.4, 1)
		})

		it("joins when only the second sets style:join-border false", async () => {
			expect(
				pitch(await lines(three('fo:padding="0.5in"', 'fo:padding="0.5in" style:join-border="false"')), "One", "Two")
			).toBeCloseTo(13.8, 0)
		})

		it("leaves out a border between paragraphs that agree", async () => {
			const bordered = 'fo:padding="0.5in" fo:border="0.01in solid #000000"'
			const all = await lines(three(bordered, bordered))

			expect(pitch(all, "Before", "One")).toBeCloseTo(50.5, 0)
			expect(pitch(all, "One", "Two")).toBeCloseTo(13.8, 0)
			expect(pitch(all, "Two", "After")).toBeCloseTo(50.5, 0)
		})
	})

	it("is left out of a converted .docx with its setting, and counted", async () => {
		const kinds = async (properties: string): Promise<string[]> => {
			const converted = await convertText(odtBytes({ styles: STYLES, ...padded(properties) }), "odt", "docx")

			if (!converted.ok) throw new Error(converted.refusal.message)

			const archive = await ZipArchive.open(bytesSource(converted.bytes))
			const settings = archive.get("word/settings.xml")

			expect(settings === undefined ? "" : decodeUtf8(await archive.read(settings))).not.toContain("paragraph-padding")

			return converted.lossy.features.map(feature => feature.kind)
		}

		expect(await kinds('fo:padding-top="0.5in"')).toContain("paragraphPadding")
		expect(await kinds('fo:border-left="0.02in solid #000000"')).toContain("paragraphPadding")
		expect(await kinds('fo:border-top="0.02in solid #000000" fo:padding-top="0.1in"')).not.toContain("paragraphPadding")
	})

	it("is written in whole points to a converted .docx, and a line kept apart is counted", async () => {
		const converted = async (properties: string): Promise<{ document: string; kinds: string[] }> => {
			const result = await convertText(odtBytes({ styles: STYLES, ...padded(properties) }), "odt", "docx")

			if (!result.ok) throw new Error(result.refusal.message)

			const archive = await ZipArchive.open(bytesSource(result.bytes))
			const entry = archive.get("word/document.xml")

			return {
				document: entry === undefined ? "" : decodeUtf8(await archive.read(entry)),
				kinds: result.lossy.features.map(feature => feature.kind)
			}
		}
		const sides =
			'fo:border-top="0.02in solid #000000" fo:border-bottom="0.02in solid #000000" fo:padding-top="0.2in" fo:padding-bottom="0.2in"'
		const plain = await converted(sides)

		expect(plain.document).toContain('w:space="14"')
		expect(plain.document).not.toMatch(/w:space="\d+\./u)
		expect(plain.kinds).not.toContain("paragraphPadding")

		const apart = await converted(`${sides} style:join-border="false"`)

		expect(apart.document).not.toContain("w:frame")
		expect(apart.kinds).toContain("paragraphPadding")
	})
})

describe("a padded ODT paragraph edited", () => {
	const body =
		'<text:p>Before</text:p><text:p text:style-name="P1">One padded</text:p><text:p text:style-name="P1">Two padded</text:p><text:p>After</text:p>'

	it.each([
		["space before", { type: "paragraph.spacing", before: 240 }],
		["first-line indent", { type: "paragraph.indent", firstLine: 240 }],
		["split in two", undefined]
	] as const)("saves with its padding kept: %s", async (_, command) => {
		const editing = await editOdt(
			odtBytes({
				automatic:
					'<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:padding="0.5in"/></style:style>',
				body
			})
		)
		let at = -1
		let paragraphs = 0

		editing.session.state.doc.descendants((node, position) => {
			if (at >= 0) return false

			if (node.type.name === "paragraph") {
				if (paragraphs++ === 1) at = position

				return false
			}

			return true
		})

		if (command === undefined) editing.session.dispatch(editing.session.state.tr.split(at + 4))
		else {
			editing.session.dispatch(editing.session.state.tr.setSelection(TextSelection.create(editing.session.state.doc, at + 4)))
			editing.editor.execute(command)
		}

		const archive = await ZipArchive.open(bytesSource((await saveOdt(editing)).bytes))
		const entry = archive.get("content.xml")
		const content = entry === undefined ? "" : decodeUtf8(await archive.read(entry))

		expect(content.match(/fo:padding(?:-top)?="(?:0\.5in|36pt)"/gu)?.length ?? 0).toBeGreaterThan(0)
		expect(content).not.toMatch(/fo:padding-top="0/u)
	})
})
