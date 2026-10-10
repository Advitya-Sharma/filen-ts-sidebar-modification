import { readdirSync } from "node:fs"
import * as docx from "@filen/office/core/docx"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { NumberingCounters, parseNumbering } from "../core/docx/numbering/index"
import { PropsStore } from "../core/docx/props/index"
import { PropertyResolver, conditionBit, headingLevel, parseStyles } from "../core/docx/styles/index"
import { parseTheme } from "../core/docx/theme/index"
import { encodeUtf8, parseXml } from "../core/xml/index"
import { OFFICE_THEME, W, find, wordPart } from "./docxFixtures"
import { openDocx, walkAll, type WalkedParagraph } from "./docxWalk"

const DIRECTORY = join(import.meta.dirname, "fixtures", "docx-styles")
// Files the model's fixtures already carry are read from there rather than copied twice.
const SHARED = join(import.meta.dirname, "fixtures", "docx-corpus")
const fixture = (name: string): string => join(DIRECTORY, `${name}.docx`)
const shared = (name: string): string => join(SHARED, `${name}.docx`)

function byText(paragraphs: readonly WalkedParagraph[], text: string): WalkedParagraph {
	const found = paragraphs.find(paragraph => paragraph.text.startsWith(text))

	if (found === undefined) throw new Error(`No paragraph starting with ${text}`)

	return found
}

describe("real documents", () => {
	it.each(readdirSync(DIRECTORY).filter(name => name.endsWith(".docx")))("resolves every paragraph and run of %s", async name => {
		const opened = await openDocx(join(DIRECTORY, name))
		const paragraphs = walkAll(opened)
		const stats = opened.resolver.stats()

		expect(paragraphs.length).toBeGreaterThan(0)
		// Interning: far fewer distinct resolved sets than paragraphs and runs.
		expect(stats.paragraphs).toBeLessThanOrEqual(paragraphs.length)
		expect(stats.runs).toBeLessThanOrEqual(
			Math.max(
				1,
				paragraphs.reduce((sum, paragraph) => sum + paragraph.runs.length, 0)
			)
		)
	})

	it("finds localised German built-in styles by canonical name and resolves their theme formatting", async () => {
		const opened = await openDocx(fixture("lo__StyleRef-DE"))
		const paragraphs = walkAll(opened, true)
		const heading = opened.styles.byName("heading 1")
		const title = byText(paragraphs, "Heading 1")

		expect(heading?.id).toBe("berschrift1")
		expect(heading && headingLevel(heading)).toBe(1)
		expect(opened.styles.defaultOf("paragraph")?.id).toBe("Standard")
		expect(opened.styles.byName("Default Paragraph Font")?.id).toBe("Absatz-Standardschriftart")
		expect(opened.styles.byName("subtitle")?.id).toBe("Untertitel")
		expect(opened.settings.themeFontLanguage.latin).toBe("hu-HU")
		expect(title.props).toMatchObject({ styleId: "berschrift1", keepNext: true, outlineLevel: 0 })
		// Word's own value for Accent 1 darker 25% of the 2013 theme.
		expect(title.runs[0]).toMatchObject({ size: 32, color: "2E74B5", bold: false })
		expect(title.runs[0]?.fonts.ascii).toBe("Calibri Light")

		const body = paragraphs.find(paragraph => paragraph.props.styleId === "Standard")

		expect(body?.props.spacing).toMatchObject({ after: 160, line: 259, lineRule: "auto" })
		expect(body?.runs[0]).toMatchObject({ size: 22, styleId: "Absatz-Standardschriftart" })
		expect(body?.runs[0]?.fonts.ascii).toBe("Calibri")
		expect(paragraphs.find(paragraph => paragraph.props.styleId === "Untertitel")?.runs[0]?.color).toBe("5A5A5A")
		expect(paragraphs.flatMap(paragraph => paragraph.runs).find(run => run.styleId === "IntensiveHervorhebung")).toMatchObject({
			italic: true,
			color: "5B9BD5"
		})
	})

	it("reads a Strict OOXML package", async () => {
		const opened = await openDocx(shared("docx4j__strict-math"))

		expect(opened.styles.styles.length).toBe(33)
		expect(opened.styles.defaultOf("paragraph")?.id).toBe("Normal")
		expect(opened.theme?.colors.get("accent1")).toBe("4472C4")
		expect(opened.theme?.minor.latin).toBe("Calibri")
		expect(opened.settings.themeFontLanguage).toEqual({ latin: "en-AU", eastAsia: "ja-JP" })

		const normal = opened.resolver.paragraph(opened.store.paragraph(opened.document, -1))

		expect(normal.props.spacing).toMatchObject({ after: 160, line: 259 })
		expect(opened.resolver.run(normal, opened.store.run(opened.document, -1)).fonts).toMatchObject({
			ascii: "Calibri",
			eastAsia: "游明朝"
		})
	})

	it("applies table style regions under an attribute-form tblLook", async () => {
		const opened = await openDocx(fixture("lo__tdf138020_undefined_firstCol"))
		const cells = walkAll(opened).filter(paragraph => paragraph.cell !== undefined)
		const firstTable = cells.slice(0, 8)

		// firstColumn on, firstRow off, lastRow on, horizontal bands off, vertical bands on.
		expect(firstTable.every(cell => cell.cell?.props.shading?.fill !== "FFC000")).toBe(true)
		expect(firstTable[1]?.cell?.props.shading?.fill).toBe("D9F2D0")
		expect(firstTable[2]?.cell?.props.shading).toBeUndefined()
		expect(cells.slice(8).every(cell => cell.cell?.props.shading?.fill === "C1E4F5")).toBe(true)
		// The table style's paragraph properties reach paragraphs in the default style.
		expect(firstTable[0]?.props.spacing).toMatchObject({ after: 0, line: 240 })
	})

	it("reads w:tblLook attributes over w:val, missing attributes taking the §17.4.55 defaults", async () => {
		const opened = await openDocx(fixture("lo__tdf167843_tblLook_firstRow_tblHeader"))
		const table = opened.resolver.table(opened.store.table(opened.document, find(opened.document, "w:tblPr")))

		expect(table.props.look).toEqual({
			firstRow: false,
			lastRow: false,
			firstColumn: false,
			lastColumn: false,
			noHBand: false,
			noVBand: false
		})
		expect(opened.resolver.row(table, opened.store.row(opened.document, -1), { rowIndex: 0, rowCount: 8 }).header).toBe(false)

		const cells = walkAll(opened).filter(paragraph => paragraph.cell !== undefined)

		const edges = conditionBit("firstRow") | conditionBit("lastRow") | conditionBit("firstCol") | conditionBit("lastCol")
		const fill = (conditions: number): string | undefined =>
			conditions & conditionBit("band1Vert") ? "D9F2D0" : conditions & conditionBit("band1Horz") ? "F2CEED" : undefined

		expect(cells.length).toBeGreaterThan(8)
		expect(cells.every(cell => ((cell.cell?.conditions ?? 0) & edges) === 0)).toBe(true)
		expect(cells.every(cell => cell.cell?.props.shading?.fill === fill(cell.cell?.conditions ?? 0))).toBe(true)
		expect(cells[0]?.cell?.props.shading?.fill).toBe("D9F2D0")
	})

	it("numbers right-to-left multilevel lists and mirrors their alignment", async () => {
		const paragraphs = walkAll(await openDocx(fixture("oxpt__HC012-Test-07")), true)
		const numbered = paragraphs.filter(paragraph => paragraph.marker !== undefined)

		expect(numbered.slice(0, 4).map(paragraph => paragraph.marker?.text)).toEqual(["1.", "2.", "2.1.", "2.2."])
		expect(numbered.find(paragraph => paragraph.props.numbering?.ilvl === 2)?.marker?.text).toBe("7.1.1.")
		expect(numbered.at(-1)?.marker?.text).toBe("14.")
		expect(numbered[0]?.props).toMatchObject({ bidi: true, physicalAlignment: "right", indent: { start: 360, firstLine: -360 } })
		expect(paragraphs[0]?.runs[0]).toMatchObject({ rtl: true, bold: true, color: "FF0000", size: 48 })
	})

	it("continues a list across nums of one abstract definition", async () => {
		const paragraphs = walkAll(await openDocx(fixture("govuk__Ofsted_Audit_Committee_minutes_1_June_2016")), true)
		const marker = (text: string): string | undefined => byText(paragraphs, text).marker?.text

		expect(marker("John Roberts opened")).toBe("1.1.")
		expect(marker("The committee agreed that members")).toBe("2.1.")
		// numId 3 shares numId 1's abstract definition, so it continues at 3.1 and numId 1 carries on at 3.2.
		expect(byText(paragraphs, "The NAO introduced").props.numbering?.numId).toBe(3)
		expect(marker("The NAO introduced")).toBe("3.1.")
		expect(marker("The financial statements presented")).toBe("3.2.")
		expect(byText(paragraphs, "accruals and payables").marker?.bullet).toBe(true)
		expect(byText(paragraphs, "John Roberts opened").props.indent).toMatchObject({ start: 720, firstLine: -720 })
	})

	it("picks East Asian and complex-script theme fonts by w:themeFontLang, not the run's language", async () => {
		const opened = await openDocx(fixture("docx4j__themeFontLang"))
		const scratch = parseXml(
			encodeUtf8(
				`<w:document xmlns:w="${W}"><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia" w:cstheme="majorBidi"/><w:lang w:eastAsia="ko-KR" w:bidi="ar-SA"/></w:rPr>` +
					'<w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/><w:lang w:eastAsia="zh-TW"/></w:rPr></w:document>'
			)
		)
		const normal = opened.resolver.paragraph(opened.store.paragraph(opened.document, -1))

		expect(opened.settings.themeFontLanguage).toEqual({ latin: "en-US", eastAsia: "ko-KR" })
		expect(opened.resolver.run(normal, opened.store.run(scratch, find(scratch, "w:rPr"))).fonts).toMatchObject({
			eastAsia: "맑은 고딕",
			cs: "Times New Roman"
		})
		expect(
			opened.resolver.run(normal, opened.store.run(scratch, find(scratch, "w:rPr", find(scratch, "w:lang") + 1))).fonts.eastAsia
		).toBe("맑은 고딕")
		// Normal names Times New Roman for East Asian text itself, which wins over the defaults' theme font.
		expect(opened.resolver.run(normal, opened.store.run(opened.document, -1)).fonts).toMatchObject({
			ascii: "Calibri",
			eastAsia: "Times New Roman"
		})
	})

	it("parses tracked run property changes", async () => {
		const opened = await openDocx(shared("oxsdk__trackChange"))
		const changed = opened.store.run(opened.document, find(opened.document, "w:rPr", find(opened.document, "w:b") - 2))

		expect(changed).toMatchObject({ bold: true, color: { val: "FF0000" } })
		expect(changed.change).toEqual({ id: 0, author: "jazh", date: "2008-05-15T10:04:00Z", previous: { fonts: { hint: "eastAsia" } } })
	})
})

describe("interning at scale", () => {
	it("keeps tens of thousands of runs to a few dozen resolved objects", () => {
		const styles = Array.from(
			{ length: 12 },
			(_, k) =>
				`<w:style w:type="paragraph" w:styleId="S${k}"><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="${k * 20}"/></w:pPr><w:rPr><w:sz w:val="${20 + (k % 4) * 2}"/></w:rPr></w:style>`
		)
		const store = new PropsStore()
		const parsedStyles = parseStyles(
			wordPart(
				"styles",
				`<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${styles.join("")}`
			),
			store
		)
		const numbering = parseNumbering(
			wordPart(
				"numbering",
				'<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
			),
			store,
			parsedStyles
		)
		const resolver = new PropertyResolver({ styles: parsedStyles, numbering, theme: parseTheme(OFFICE_THEME) })
		const shapes = [
			"",
			"<w:b/>",
			"<w:i/>",
			'<w:color w:val="C00000"/>',
			'<w:b/><w:sz w:val="28"/>',
			'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
		]
		const body: string[] = []

		for (let p = 0; p < 20000; p++) {
			const list = p % 7 === 0 ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ""
			const runs = [0, 1, 2].map(
				k => `<w:r><w:rPr><w:lang w:val="en-GB"/>${shapes[(p + k) % shapes.length]}</w:rPr><w:t>x</w:t></w:r>`
			)

			body.push(`<w:p><w:pPr><w:pStyle w:val="S${p % 12}"/>${list}</w:pPr>${runs.join("")}</w:p>`)
		}

		const document = wordPart("document", `<w:body>${body.join("")}</w:body>`)
		const counters = new NumberingCounters(numbering)
		const rPr = document.nameIdOf("w:rPr")
		const pPr = document.nameIdOf("w:pPr")
		let runs = 0
		let last = ""
		let format = resolver.paragraph(store.paragraph(document, -1))

		for (let node = 0; node < document.count; node++) {
			const id = document.nameId(node)

			if (id === pPr) {
				format = resolver.paragraph(store.paragraph(document, node))

				if (format.props.numbering !== undefined)
					last = counters.next(format.props.numbering.numId, format.props.numbering.ilvl)?.text ?? ""
			} else if (id === rPr) {
				resolver.run(format, store.run(document, node))
				runs++
			}
		}

		const stats = resolver.stats()

		expect(runs).toBe(60000)
		expect(last).toBe(`${Math.ceil(20000 / 7)}.`)
		// The styles' own containers are in the store too: four run sizes, twelve paragraph spacings.
		expect(store.stats().runs).toBe(shapes.length + 4)
		expect(store.stats().paragraphs).toBe(24 + 12)
		expect(stats.runs).toBeLessThanOrEqual(4 * shapes.length)
		// 24 style and list combinations plus the empty paragraph the walk starts with.
		expect(stats.paragraphs).toBeLessThanOrEqual(25)
		expect(stats.runResolutions).toBeLessThanOrEqual(24 * shapes.length)
	})
})

describe("package entry", () => {
	it("exports the property, style, theme and numbering layer from @filen/office/core/docx", () => {
		expect(docx.PropertyResolver).toBe(PropertyResolver)
		expect(docx.PropsStore).toBe(PropsStore)
		expect(typeof docx.parseTheme).toBe("function")
		expect(typeof docx.NumberingCounters).toBe("function")
		expect(typeof docx.partReader).toBe("function")
	})
})
