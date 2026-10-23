import { describe, expect, it } from "vitest"
import { OpcError, OpcPackage } from "@filen/office/core/opc"
import { openDocxParts } from "../core/docx/package"
import { encodeUtf8 } from "../core/xml/index"

const CONTENT_TYPES =
	'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
const ROOT_RELS =
	'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
const DOCUMENT =
	'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>'

function parts(): Map<string, Uint8Array> {
	return new Map([
		["[Content_Types].xml", encodeUtf8(CONTENT_TYPES)],
		["_rels/.rels", encodeUtf8(ROOT_RELS)],
		["word/document.xml", encodeUtf8(DOCUMENT)]
	])
}

describe("an OPC package over parts held in memory", () => {
	it("reads parts by OPC name rules without copying them", async () => {
		const map = parts()
		const opc = OpcPackage.fromParts(map)

		expect(opc.has("/WORD/Document.xml")).toBe(true)
		expect(await opc.read("word/document.xml")).toBe(map.get("word/document.xml"))
		expect(opc.entry("word/document.xml")?.uncompressedSize).toBe(map.get("word/document.xml")?.length)
		expect((await opc.relationships("")).all[0]?.partName).toBe("word/document.xml")
	})

	it("refuses a part map without content types or with names OPC forbids", () => {
		const noTypes = parts()

		noTypes.delete("[Content_Types].xml")
		expect(() => OpcPackage.fromParts(noTypes)).toThrow(OpcError)
		expect(() => OpcPackage.fromParts(new Map([...parts(), ["word/a//b.xml", new Uint8Array(0)]]))).toThrow(OpcError)
	})

	it("has no archive to save from", () => {
		expect(() => OpcPackage.fromParts(parts()).zip).toThrow(OpcError)
	})

	it("opens as a word-processing package", async () => {
		const opened = await openDocxParts(OpcPackage.fromParts(parts()))

		expect(opened.ok).toBe(true)

		if (!opened.ok) return

		expect(opened.package.mainPart).toBe("word/document.xml")
		expect(opened.package.facts.variant).toBe("docx")
	})
})
