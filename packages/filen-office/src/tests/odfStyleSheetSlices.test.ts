import { describe, expect, it } from "vitest"
import { OdfStyleResolver, readContentPrologue } from "../core/odf/index"
import { readContentPrologueSliced } from "../core/odf/prologue"
import { OdfStyleSheet } from "../core/odf/styles"
import { encodeUtf8, parseXml } from "../core/xml/index"
import { ODF_NS } from "./odtFixtures"

// A style sheet read in slices, and one that keeps few of the automatic styles it read, give every style as one read
// at once and keeping them all does.

const COUNT = 300

function content(): Uint8Array {
	let automatic = ""

	for (let k = 0; k < COUNT; k++) {
		automatic +=
			`<style:style style:name="P${k}" style:family="paragraph" style:parent-style-name="${k % 7 === 0 ? "Missing" : "Standard"}"><style:paragraph-properties fo:margin-left="${k}mm"/><style:text-properties fo:color="#0000${(k % 256).toString(16).padStart(2, "0")}"/></style:style>` +
			`<style:style style:name="T${k}" style:family="text"><style:text-properties fo:font-size="${8 + (k % 9)}pt"/></style:style>`
	}

	// A name given twice: the first wins.
	automatic += '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:margin-left="99mm"/></style:style>'

	return encodeUtf8(
		`<office:document-content ${ODF_NS}><office:automatic-styles>${automatic}</office:automatic-styles><office:body><office:text/></office:body></office:document-content>`
	)
}

const STYLES = parseXml(
	encodeUtf8(
		`<office:document-styles ${ODF_NS}><office:styles><style:style style:name="Standard" style:family="paragraph"><style:text-properties fo:font-size="11pt"/></style:style></office:styles></office:document-styles>`
	)
)

function view(sheet: OdfStyleSheet, resolver: OdfStyleResolver): unknown[] {
	const out: unknown[] = []

	for (let k = 0; k < COUNT; k++) {
		for (const [family, name] of [
			["paragraph", `P${k}`],
			["text", `T${k}`]
		] as const) {
			const style = sheet.style(family, name, "content")
			const resolved = resolver.resolve(family, name, "content")

			out.push([style?.properties, style?.parent, resolved.properties("paragraph"), resolved.properties("text")])
		}
	}

	return out
}

describe("ODF style sheet", () => {
	it("reads in slices and lets read styles go as a sheet read at once that keeps them all", async () => {
		const bytes = content()
		const whole = new OdfStyleSheet(STYLES, readContentPrologue(bytes))
		const slicing = { sliceMs: 0, now: () => 0, pause: () => Promise.resolve() }
		const sliced = await OdfStyleSheet.build(STYLES, await readContentPrologueSliced(bytes, slicing), slicing)

		sliced.keepAtMost = 3

		const expected = view(whole, new OdfStyleResolver(whole))

		expect(view(sliced, new OdfStyleResolver(sliced, { keepAtMost: 5 }))).toEqual(expected)
		// Asked again, after the sheet let them go.
		expect(view(sliced, new OdfStyleResolver(sliced, { keepAtMost: 5 }))).toEqual(expected)
		expect(sliced.issues).toEqual(whole.issues)
		expect(whole.issues.some(issue => issue.kind === "duplicate")).toBe(true)
		expect(whole.issues.some(issue => issue.kind === "missingParent")).toBe(true)
	})
})
