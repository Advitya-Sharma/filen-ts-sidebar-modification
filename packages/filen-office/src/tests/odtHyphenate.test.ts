import { describe, expect, it } from "vitest"
import { bytesSource, ZipArchive } from "../core/opc/index"
import { translateOdt } from "../core/odt/index"
import { writeOdt } from "../core/odt/write/writer"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { odtBytes } from "./odtFixtures"
import { odtSource } from "./odtWriteFixtures"

// Hyphenation is a text property in ODF (fo:hyphenate on style:text-properties); WordprocessingML has it per
// paragraph (w:suppressAutoHyphens). The translation reads a paragraph style's text-properties fo:hyphenate as the
// paragraph's, and the writer writes it there, never on style:paragraph-properties (no such attribute there).

async function part(bytes: Uint8Array, name: string): Promise<string> {
	const archive = await ZipArchive.open(bytesSource(bytes))
	const entry = archive.get(name)

	return entry === undefined ? "" : decodeUtf8(await archive.read(entry))
}

const NO_HYPHENS =
	'<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:hyphenate="false"/></style:style>'
const STYLES =
	'<style:style style:name="Standard" style:family="paragraph" style:class="text"><style:text-properties fo:hyphenate="true"/></style:style>'

describe("ODT hyphenation", () => {
	it("reads a paragraph style's fo:hyphenate as w:suppressAutoHyphens", async () => {
		const translated = await translateOdt(
			bytesSource(
				odtBytes({ automatic: NO_HYPHENS, styles: STYLES, body: '<text:p text:style-name="P1">a</text:p><text:p>b</text:p>' })
			)
		)

		if (!translated.ok) throw new Error(translated.message)

		const body = decodeUtf8(translated.translation.parts.get("word/document.xml") as Uint8Array)
		const styles = decodeUtf8(translated.translation.parts.get("word/styles.xml") as Uint8Array)

		expect(body).toMatch(/<w:pPr>[^]*?<w:suppressAutoHyphens\/>[^]*?<\/w:pPr><w:r><w:t[^>]*>a</u)
		expect(styles).toMatch(/w:styleId="Standard"[^]*?<w:suppressAutoHyphens w:val="0"\/>/u)
	})

	it("writes w:suppressAutoHyphens as text-properties fo:hyphenate", async () => {
		const source = await odtSource(odtBytes({ styles: STYLES, body: "<text:p>a</text:p><text:p>b</text:p>" }))
		const document = source.translation.parts.get("word/document.xml") as Uint8Array
		const edited = decodeUtf8(document).replace(/<w:p>(<w:pPr>)?/u, match =>
			match === "<w:p><w:pPr>" ? "<w:p><w:pPr><w:suppressAutoHyphens/>" : "<w:p><w:pPr><w:suppressAutoHyphens/></w:pPr>"
		)

		expect(edited).toContain("suppressAutoHyphens")

		const written = await writeOdt(source, { parts: new Map([["word/document.xml", encodeUtf8(edited)]]) })
		const content = await part(written.bytes, "content.xml")

		expect(content).not.toMatch(/<style:paragraph-properties[^>]*fo:hyphenate/u)
		expect(content).toMatch(/<style:text-properties[^>]*fo:hyphenate="false"/u)
	})
})
