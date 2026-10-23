import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { parseXml } from "@filen/office/core/xml"
import {
	OdfPackage,
	OdfStyleResolver,
	OdfStyleSheet,
	parseAngle,
	parseColor,
	parseLength,
	parseLengthOrPercent,
	parseRelativeLength,
	readContentPrologue,
	readStyleSheet
} from "../core/odf/index"
import { CORPUS_DIR } from "./corpusDir"
import { text } from "./zipBuilder"

const NS =
	'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"'

const STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NS} office:version="1.3">
<office:font-face-decls>
 <style:font-face style:name="Liberation Serif" svg:font-family="'Liberation Serif'" style:font-family-generic="roman" style:font-pitch="variable"/>
 <style:font-face style:name="Embedded" svg:font-family="Embedded"><svg:font-face-src><svg:font-face-uri xlink:href="Fonts/e.ttf"/></svg:font-face-src></style:font-face>
</office:font-face-decls>
<office:styles>
 <style:default-style style:family="paragraph">
  <style:paragraph-properties fo:margin="0.1in" style:tab-stop-distance="1.25cm"/>
  <style:text-properties style:font-name="Liberation Serif" fo:font-size="12pt" fo:language="en"/>
 </style:default-style>
 <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
 <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:list-style-name="">
  <style:paragraph-properties fo:margin-top="0.423cm" fo:margin-left="200%"><style:tab-stops><style:tab-stop style:position="1cm"/></style:tab-stops></style:paragraph-properties>
  <style:text-properties fo:font-size="130%" fo:font-weight="bold"/>
 </style:style>
 <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:default-outline-level="1">
  <style:paragraph-properties fo:border="0.5pt solid #000000" fo:border-left="none"/>
  <style:text-properties fo:font-size="115%"/>
 </style:style>
 <style:style style:name="Loop_A" style:family="paragraph" style:parent-style-name="Loop_B"/>
 <style:style style:name="Loop_B" style:family="paragraph" style:parent-style-name="Loop_A"/>
 <style:style style:name="Orphan" style:family="paragraph" style:parent-style-name="Nowhere"/>
 <style:style style:name="Standard" style:family="text"><style:text-properties fo:color="#ff0000"/></style:style>
 <text:outline-style style:name="Outline"><text:outline-level-style text:level="1" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.762cm"/></style:list-level-properties></text:outline-level-style></text:outline-style>
 <text:list-style style:name="Bullets"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:text-properties style:font-name="OpenSymbol"/></text:list-level-style-bullet><text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"/><text:list-level-style-number text:level="11" style:num-format="1"/></text:list-style>
 <number:currency-style style:name="N104P0" style:volatile="true"><number:currency-symbol number:language="de" number:country="DE">€</number:currency-symbol><number:text> </number:text><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>
 <number:currency-style style:name="N104"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text><number:number number:decimal-places="2" number:min-integer-digits="1"/><style:map style:condition="value()&gt;=0" style:apply-style-name="N104P0"/></number:currency-style>
 <text:notes-configuration text:note-class="footnote" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/>
 <draw:gradient draw:name="Gradient_1" draw:style="linear" draw:start-color="#000000" draw:end-color="#ffffff"/>
</office:styles>
<office:automatic-styles>
 <style:style style:name="MP1" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="9pt"/></style:style>
 <style:page-layout style:name="pm1" style:page-usage="mirrored">
  <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" fo:margin="2cm" fo:margin-bottom="1cm" loext:margin-gutter="0cm"><style:columns fo:column-count="2" fo:column-gap="0.5cm"/></style:page-layout-properties>
  <style:header-style><style:header-footer-properties fo:min-height="0cm" fo:margin-bottom="0.5cm"/></style:header-style>
  <style:footer-style/>
 </style:page-layout>
</office:automatic-styles>
<office:master-styles>
 <style:master-page style:name="Standard" style:page-layout-name="pm1" style:next-style-name="Standard"><style:header><text:p text:style-name="MP1">Head</text:p></style:header><style:footer style:display="false"><text:p/></style:footer><style:header-first><text:p/></style:header-first></style:master-page>
 <style:master-page style:name="Landscape" style:display-name="Landscape page" style:page-layout-name="pm1"/>
</office:master-styles>
</office:document-styles>`

const CONTENT = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NS} office:version="1.3">
<office:automatic-styles>
 <style:style style:name="MP1" style:family="paragraph" style:parent-style-name="Heading_20_1" style:master-page-name="Landscape"><style:text-properties fo:font-size="50%"/></style:style>
 <style:style style:name="T1" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>
 <text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="i"/></text:list-style>
 <number:date-style style:name="N37" number:automatic-order="true"><number:day number:style="long"/><number:text>.</number:text><number:month/></number:date-style>
</office:automatic-styles>
<office:body><office:text><text:p text:style-name="MP1">x</text:p></office:text></office:body>
</office:document-content>`

const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"

function sheet(): OdfStyleSheet {
	return new OdfStyleSheet(parseXml(text(STYLES)), readContentPrologue(text(CONTENT)))
}

describe("ODF value types", () => {
	it("reads lengths in every unit as points and refuses what is not one", () => {
		expect(parseLength("1in")).toBe(72)
		expect(parseLength("2.54cm")).toBeCloseTo(72, 9)
		expect(parseLength("25.4mm")).toBeCloseTo(72, 9)
		expect(parseLength("1pc")).toBe(12)
		expect(parseLength("4px")).toBe(3)
		expect(parseLength("-.5pt")).toBe(-0.5)
		expect(parseLength("0.5inch")).toBe(36)
		expect(parseLength("12")).toBeUndefined()
		expect(parseLength("1em")).toBeUndefined()
		expect(parseLength(undefined)).toBeUndefined()
		expect(parseLengthOrPercent("115%")).toEqual({ kind: "percent", percent: 115 })
		expect(parseRelativeLength("2716*")).toBe(2716)
		expect(parseAngle("100grad")).toBe(90)
		expect(parseAngle("45")).toBe(45)
		expect(parseColor("#C0ffee")).toBe(0xc0ffee)
		expect(parseColor("red")).toBeUndefined()
	})
})

describe("ODF style sheet", () => {
	it("reads font faces with their unquoted family and embedded sources", () => {
		const fonts = sheet().fontFaces

		expect(fonts.get("Liberation Serif")).toMatchObject({
			family: "Liberation Serif",
			generic: "roman",
			pitch: "variable",
			sources: []
		})
		expect(fonts.get("Embedded")?.sources).toEqual(["Fonts/e.ttf"])
	})

	it("resolves a style through its parents down from the family default", () => {
		const resolved = new OdfStyleResolver(sheet()).resolve("paragraph", "Heading_20_1", "styles")

		expect(resolved.chain.map(style => style.name)).toEqual(["", "Standard", "Heading", "Heading_20_1"])
		// 12pt × 130% × 115%.
		expect(parseLength(resolved.get("text", "fo:font-size"))).toBeCloseTo(17.94, 9)
		expect(resolved.get("text", "fo:font-weight")).toBe("bold")
		expect(resolved.get("text", "style:font-name")).toBe("Liberation Serif")
		expect(resolved.get("paragraph", "fo:margin-top")).toBe("0.423cm")
		// A percent margin is relative to the parent's (0.1in from the default style's shorthand).
		expect(resolved.get("paragraph", "fo:margin-left")).toBe("14.4pt")
		expect(resolved.get("paragraph", "fo:margin-bottom")).toBe("0.1in")
		// The border shorthand stands for each side not given in the same element.
		expect(resolved.get("paragraph", "fo:border-top")).toBe("0.5pt solid #000000")
		expect(resolved.get("paragraph", "fo:border-left")).toBe("none")
		expect(resolved.get("paragraph", "fo:border")).toBeUndefined()
		expect(resolved.properties("paragraph").children.map(child => child.name)).toEqual(["style:tab-stops"])
	})

	it("looks a name up among the referring stream's automatic styles before the common ones", () => {
		const resolver = new OdfStyleResolver(sheet())
		const body = resolver.resolve("paragraph", "MP1", "content")
		const header = resolver.resolve("paragraph", "MP1", "styles")

		expect(body.chain.map(style => [style.name, style.origin, style.automatic])).toEqual([
			["", "styles", false],
			["Standard", "styles", false],
			["Heading", "styles", false],
			["Heading_20_1", "styles", false],
			["MP1", "content", true]
		])
		expect(parseLength(body.get("text", "fo:font-size"))).toBeCloseTo(8.97, 9)
		expect(header.get("text", "fo:font-size")).toBe("9pt")
		expect(resolver.inherited("paragraph", "MP1", "content", "style:master-page-name")).toBe("Landscape")
		// An empty list-style-name on an ancestor is an explicit "no list", kept as such.
		expect(resolver.inherited("paragraph", "MP1", "content", "style:list-style-name")).toBe("")
		expect(resolver.inherited("paragraph", "MP1", "content", "style:default-outline-level")).toBe("1")
	})

	it("keeps families apart and resolves an unknown name to the family default", () => {
		const resolver = new OdfStyleResolver(sheet())

		expect(resolver.resolve("text", "Standard", "content").get("text", "fo:color")).toBe("#ff0000")
		expect(resolver.resolve("paragraph", "Standard", "content").get("text", "fo:color")).toBeUndefined()
		expect(resolver.resolve("paragraph", "Nope", "content").chain.map(style => style.name)).toEqual([""])
		expect(resolver.resolve("text", "T1", "content").get("text", "fo:font-style")).toBe("italic")
	})

	it("stops at a parent cycle and reports cycles and missing parents", () => {
		const styles = sheet()
		const resolver = new OdfStyleResolver(styles)
		const resolved = resolver.resolve("paragraph", "Loop_A", "content")

		expect(resolved.truncated).toBe(true)
		expect(resolved.chain.map(style => style.name)).toEqual(["", "Loop_B", "Loop_A"])
		expect(resolver.resolve("paragraph", "Orphan", "content").chain.map(style => style.name)).toEqual(["", "Orphan"])
		expect(styles.issues).toContainEqual({ kind: "missingParent", family: "paragraph", name: "Orphan", parent: "Nowhere" })
		expect(styles.issues).toContainEqual({ kind: "parentCycle", family: "paragraph", name: "Loop_A" })
	})

	it("reads page layouts and master pages with their header and footer regions", () => {
		const styles = sheet()
		const layout = styles.pageLayouts.get("pm1")

		expect(layout?.usage).toBe("mirrored")
		expect(parseLength(layout?.properties.attributes.get("fo:page-width"))).toBeCloseTo(595.3, 1)
		expect(layout?.properties.attributes.get("fo:margin-top")).toBe("2cm")
		expect(layout?.properties.attributes.get("fo:margin-bottom")).toBe("1cm")
		expect(layout?.properties.attributes.get("loext:margin-gutter")).toBe("0cm")
		expect(layout?.properties.children[0]?.attributes.get("fo:column-count")).toBe("2")
		expect(layout?.header?.attributes.get("fo:margin-bottom")).toBe("0.5cm")
		expect(layout?.footer?.attributes.size).toBe(0)
		expect(styles.masterOrder).toEqual(["Standard", "Landscape"])

		const standard = styles.masterPages.get("Standard")

		expect(standard?.regions.map(region => [region.kind, region.display])).toEqual([
			["header", true],
			["footer", false],
			["header-first", true]
		])
		expect(styles.stylesDocument?.textContent(standard?.regions[0]?.node ?? -1)).toBe("Head")
		expect(styles.masterPages.get("Landscape")?.displayName).toBe("Landscape page")
	})

	it("reads list styles, the outline style and their levels", () => {
		const styles = sheet()
		const bullets = styles.listStyle("Bullets", "content")

		expect([...(bullets?.levels.keys() ?? [])]).toEqual([1, 2])
		expect(bullets?.levels.get(1)).toMatchObject({ kind: "bullet" })
		expect(bullets?.levels.get(1)?.attributes.get("text:bullet-char")).toBe("•")
		expect(bullets?.levels.get(1)?.textProperties?.attributes.get("style:font-name")).toBe("OpenSymbol")
		expect(bullets?.levels.get(2)?.attributes.get("style:num-suffix")).toBe(")")
		expect(styles.listStyle("L1", "content")?.automatic).toBe(true)
		expect(styles.listStyle("L1", "styles")).toBeUndefined()
		expect(styles.outlineStyle?.outline).toBe(true)
		expect(styles.outlineStyle?.levels.get(1)?.labelAlignment?.get("fo:text-indent")).toBe("-0.762cm")
	})

	it("reads data styles with their parts in order, conditions and colors", () => {
		const styles = sheet()
		const negative = styles.dataStyle("N104", "content")

		expect(negative?.kind).toBe("currency")
		expect(negative?.parts.map(part => part.name)).toEqual(["number:text", "number:number"])
		expect(negative?.maps).toEqual([{ condition: "value()>=0", applyStyleName: "N104P0", baseCellAddress: undefined }])
		expect(negative?.textProperties?.attributes.get("fo:color")).toBe("#ff0000")
		expect(styles.dataStyle("N104P0", "content")?.parts[0]).toMatchObject({ name: "number:currency-symbol", text: "€" })
		expect(styles.dataStyle("N37", "content")?.parts.map(part => part.name)).toEqual(["number:day", "number:text", "number:month"])
		expect(styles.dataStyle("N37", "styles")).toBeUndefined()
	})

	it("reads notes configuration and named drawing resources", () => {
		const styles = sheet()

		expect(styles.notesConfiguration.get("footnote")?.get("text:start-numbering-at")).toBe("document")
		expect(styles.drawResources.get("draw:gradient")?.get("Gradient_1")?.attributes.get("draw:end-color")).toBe("#ffffff")
	})

	it("reads styles bound to other prefixes the same way", () => {
		const document = parseXml(
			text(
				STYLES.replaceAll("style:", "s:").replace(
					'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:s:1.0"',
					`xmlns:s="${STYLE_NS}"`
				)
			)
		)
		const styles = new OdfStyleSheet(document, undefined)

		expect(new OdfStyleResolver(styles).resolve("paragraph", "Heading_20_1", "styles").get("text", "fo:font-weight")).toBe("bold")
	})
})

const ODT = join(CORPUS_DIR, "odt")
const ODS = join(CORPUS_DIR, "ods")

describe.skipIf(!existsSync(ODT) || !existsSync(ODS))("ODF corpus styles", () => {
	it("reads every corpus style sheet and resolves every style the body names", async () => {
		const unresolved: string[] = []
		let references = 0

		for (const directory of [ODT, ODS]) {
			for (const name of readdirSync(directory).sort()) {
				const bytes = new Uint8Array(readFileSync(join(directory, name)))
				const opened = await OdfPackage.open(bytesSource(bytes))

				if (!opened.ok) continue

				const { sheet: styles, prologue } = await readStyleSheet(opened.package)
				const resolver = new OdfStyleResolver(styles)
				const content = (await opened.package.bytes("content.xml")) ?? new Uint8Array(0)
				const body = new TextDecoder().decode(content.subarray(Math.max(0, prologue?.bodyStart ?? 0)))

				expect(prologue?.sliced, name).toBe(true)

				for (const match of body.matchAll(/<text:(p|h) [^>]*?text:style-name="([^"]+)"/g)) {
					references++

					const styleName = (match[2] ?? "").replaceAll("&amp;", "&")

					if (styles.style("paragraph", styleName, "content") === undefined) unresolved.push(`${name}:${styleName}`)

					resolver.resolve("paragraph", styleName, "content").properties("text")
				}
			}
		}

		expect(references).toBeGreaterThan(10_000)
		expect(unresolved).toEqual([])
	}, 60_000)
})
