import { readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { strToU8, unzipSync, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { OpcPackage, bytesSource, memorySink, writeZip } from "@filen/office/core/opc"
import { NodeKind, XmlEditor, assertWellFormed, decodeUtf8, parseXml } from "@filen/office/core/xml"
import { wordLikeDocument } from "./wordLike"
import { sameBytes } from "./zipBuilder"
import { ZIP_ENTRY_TIME } from "./zipTime"

const FIXTURES = join(import.meta.dirname, "fixtures")
const fixtures = readdirSync(FIXTURES).filter(name => name.endsWith(".docx"))

describe("corpus round trip", () => {
	it.each(fixtures)("parses every XML part of %s and re-emits it byte for byte", async name => {
		const opc = await OpcPackage.open(bytesSource(new Uint8Array(readFileSync(join(FIXTURES, name)))))
		let parts = 0

		for (const entry of opc.zip.entries) {
			if (!/\.(xml|rels)$/.test(entry.name)) continue

			const bytes = await opc.zip.read(entry)
			const document = parseXml(bytes)

			assertWellFormed(document)
			parts++

			// An editor with no edits hands back the input; one that rewrites every top-level node with its
			// own bytes must produce the same bytes again.
			expect(new XmlEditor(document).serialize()).toBe(bytes)

			const editor = new XmlEditor(document)

			for (let child = document.firstChild(0); child >= 0; child = document.nextSibling(child)) {
				editor.replace(child, document.raw(child).slice())
			}

			expect(sameBytes(editor.serialize(), bytes)).toBe(true)

			// Spans nest: every node lies inside its parent, and siblings follow each other.
			let broken = 0

			for (let node = 1; node < document.count; node++) {
				const parent = document.parent(node)
				const next = document.nextSibling(node)

				if (parent !== 0 && document.start(node) < document.openEnd(parent)) broken++
				if (document.end(node) > document.end(parent)) broken++
				if (next >= 0 && document.start(next) !== document.end(node)) broken++
			}

			expect(broken).toBe(0)
		}

		expect(parts).toBeGreaterThan(5)
	})
})

describe("large generated document", () => {
	it("edits one paragraph of a multi-megabyte document.xml and saves everything else untouched", async () => {
		const documentXml = wordLikeDocument(4000)
		const media = new Uint8Array(512 * 1024).map((_, k) => (k * 2654435761) >>> 24)
		const archive = zipSync(
			{
				"[Content_Types].xml": strToU8(
					'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>'
				),
				"word/document.xml": documentXml,
				"word/media/image1.png": [media, { level: 0 }]
			},
			{ level: 6, mtime: ZIP_ENTRY_TIME }
		)
		const opc = await OpcPackage.open(bytesSource(archive))
		const bytes = await opc.read("word/document.xml")
		const document = parseXml(bytes)

		expect(bytes.length).toBeGreaterThan(1_500_000)

		const paragraph = document.nameIdOf("w:p")
		const paragraphs: number[] = []

		for (let node = 0; node < document.count; node++) {
			if (document.kinds[node] === NodeKind.Element && document.nameIds[node] === paragraph) paragraphs.push(node)
		}

		expect(paragraphs.length).toBe(4000)

		const editor = new XmlEditor(document)

		editor.replace(paragraphs[2000] ?? -1, '<w:p><w:r><w:t xml:space="preserve">Edited by the test</w:t></w:r></w:p>')

		const edited = editor.serialize()
		const sink = memorySink()

		await writeZip(await opc.plan({ replace: [{ name: "word/document.xml", bytes: edited }] }), opc.zip, sink)

		const saved = await OpcPackage.open(bytesSource(sink.bytes()))
		const reread = await saved.read("word/document.xml")
		const files = unzipSync(sink.bytes())

		expect(sameBytes(reread, edited)).toBe(true)
		expect(sameBytes(files["word/media/image1.png"], media)).toBe(true)
		expect(decodeUtf8(reread)).toContain("Edited by the test")

		const originalImage = opc.entry("word/media/image1.png")
		const savedImage = saved.entry("word/media/image1.png")

		if (originalImage === undefined || savedImage === undefined) throw new Error("image missing")

		expect(sameBytes(await saved.zip.raw(savedImage), await opc.zip.raw(originalImage))).toBe(true)
	})
})
