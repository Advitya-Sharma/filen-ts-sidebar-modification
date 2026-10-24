import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { ZipArchive, bytesSource } from "@filen/office/core/opc"
import { decodeUtf8 } from "@filen/office/core/xml"
import { proveOdt, writeOdt } from "../core/odt/write/writer"
import { CORPUS_DIR } from "./corpusDir"
import { odtBytes } from "./odtFixtures"
import { editOdt, odtEditsOf, odtSource } from "./odtWriteFixtures"

const STYLES = '<style:default-style style:family="paragraph"><style:text-properties fo:font-size="12pt"/></style:default-style>'
const BODY =
	'<text:p>Empty link: <text:a xlink:type="simple" xlink:href="https://example.com/"/></text:p>' +
	'<text:p>Hello <text:a xlink:type="simple" xlink:href="https://example.com/">a link</text:a> world</text:p>' +
	"<text:p>Plain paragraph</text:p>"
const STYLE_SELECTION = join(CORPUS_DIR, "odt", "odftk__TestStyleSelection.odt")

describe("an empty (self-closed) text:a", () => {
	it("leaves every block editable", async () => {
		const proof = await proveOdt(await odtSource(odtBytes({ styles: STYLES, body: BODY })))

		expect(proof.locked.length).toBe(0)
	})

	it("typing beside another link to the same target writes that link as a start tag, content and end tag", async () => {
		const edited = await editOdt(odtBytes({ styles: STYLES, body: BODY }), { prove: false })
		let at = -1
		let paragraphs = 0

		edited.session.state.doc.descendants((node, position) => {
			if (at >= 0) return false

			if (node.type.name === "paragraph") {
				if (paragraphs++ === 1) at = position

				return false
			}

			return true
		})
		edited.session.dispatch(edited.session.state.tr.insertText("Z", at + 1))

		const written = await writeOdt(edited.source, await odtEditsOf(edited))
		const archive = await ZipArchive.open(bytesSource(written.bytes))
		const entry = archive.get("content.xml")
		const xml = entry === undefined ? "" : decodeUtf8(await archive.read(entry))

		expect(xml).toContain('ZHello <text:a xlink:type="simple" xlink:href="https://example.com/">a link</text:a>')
	})

	it.skipIf(!existsSync(STYLE_SELECTION))("locks no block of a corpus document holding one", async () => {
		const proof = await proveOdt(await odtSource(new Uint8Array(readFileSync(STYLE_SELECTION))))

		expect(proof.locked.length).toBe(0)
	})
})
