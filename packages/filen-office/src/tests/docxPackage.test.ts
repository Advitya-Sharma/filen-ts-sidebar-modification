import { strToU8, unzipSync, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { cfbBytes } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import { buildDocxModel, gateReasons, openDocxPackage, relationshipKind, type DocxFacts } from "@filen/office/core/docx"
import { buildDocx, openSynthetic, p } from "./docxSynthetic"

const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"

async function open(bytes: Uint8Array) {
	return openDocxPackage(bytesSource(bytes))
}

// A compound file holding a small stream of each name (the package's own CFB writer).
function compoundFile(names: readonly string[]): Promise<Uint8Array> {
	return cfbBytes(names.map(name => ({ path: name.split("/"), bytes: new Uint8Array(16).fill(7) })))
}

function facts(overrides: Partial<DocxFacts>): DocxFacts {
	return {
		variant: "docx",
		strict: false,
		macros: false,
		template: false,
		signatures: 0,
		documentProtection: undefined,
		writeProtection: undefined,
		glossary: false,
		attachedTemplate: false,
		subDocuments: 0,
		malformed: false,
		...overrides
	}
}

describe("openDocxPackage", () => {
	it("opens a plain document and finds its parts by relationship", async () => {
		const pkg = await openSynthetic({ body: p("x"), footnotes: "", settings: "" })

		expect(pkg.mainPart).toBe("word/document.xml")
		expect(pkg.facts).toEqual(facts({}))
		expect(pkg.part("footnotes")).toBe("word/footnotes.xml")
		expect(pkg.settingsPart).toBe("word/settings.xml")
		expect(pkg.settings.present).toBe(true)
		expect(gateReasons(pkg.facts)).toEqual([])

		const document = await pkg.xml("word/document.xml")

		expect(await pkg.xml("word/document.xml")).toBe(document)

		pkg.release("word/document.xml")

		expect(await pkg.xml("word/document.xml")).not.toBe(document)
	})

	it.each([
		[`${WML}.document.main+xml`, "docx", false, false],
		["application/vnd.ms-word.document.macroEnabled.main+xml", "docm", true, false],
		[`${WML}.template.main+xml`, "dotx", false, true],
		["application/vnd.ms-word.template.macroEnabledTemplate.main+xml", "dotm", true, true]
	])("tells the variant from the main content type %s", async (contentType, variant, macros, template) => {
		const pkg = await openSynthetic({ body: p("x"), mainContentType: contentType })

		expect(pkg.facts).toMatchObject({ variant, macros, template })
		// Templates and macro-enabled documents are edited like documents (their content type is kept).
		expect(gateReasons(pkg.facts)).toEqual([])
	})

	it("treats a macro project under a .docx content type as docm", async () => {
		const pkg = await openSynthetic({
			body: p("x"),
			relationships: [{ id: "rIdVba", type: "vbaProject", target: "vbaProject.bin" }],
			parts: [{ name: "word/vbaProject.bin", contentType: "application/vnd.ms-office.vbaProject", data: new Uint8Array([1, 2, 3]) }]
		})

		expect(pkg.facts).toMatchObject({ variant: "docm", macros: true })
	})

	it("recognises Strict packages", async () => {
		const pkg = await openSynthetic({ body: p("x"), strict: true, settings: "" })

		expect(pkg.facts.strict).toBe(true)
		expect(pkg.settings.strict).toBe(true)
		expect(gateReasons(pkg.facts)).toEqual(["strict"])
	})

	it("counts digital signatures", async () => {
		const origin = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
		const signature = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
		const pkg = await openSynthetic({
			body: p("x"),
			packageRelationships: [{ id: "rIdSig", type: origin, target: "_xmlsignatures/origin.sigs" }],
			parts: [
				{
					name: "_xmlsignatures/origin.sigs",
					contentType: "application/vnd.openxmlformats-package.digital-signature-origin",
					data: ""
				},
				{
					name: "_xmlsignatures/_rels/origin.sigs.rels",
					data: `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${signature}" Target="sig1.xml"/><Relationship Id="rId2" Type="${signature}" Target="sig2.xml"/></Relationships>`
				},
				{
					name: "_xmlsignatures/sig1.xml",
					contentType: "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
					data: "<Signature/>"
				},
				{
					name: "_xmlsignatures/sig2.xml",
					contentType: "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
					data: "<Signature/>"
				}
			]
		})

		expect(pkg.facts.signatures).toBe(2)
		expect(gateReasons(pkg.facts)).toEqual(["signed"])
	})

	it("reports protection, glossary, attached templates and sub-documents", async () => {
		const pkg = await openSynthetic({
			body: p("x"),
			settings:
				'<w:writeProtection w:hash="x" w:salt="y"/><w:documentProtection w:edit="readOnly" w:enforcement="on"/><w:attachedTemplate r:id="rIdT"/>',
			relationships: [
				{ id: "rIdG", type: "glossaryDocument", target: "glossary/document.xml" },
				{ id: "rIdS", type: "subDocument", target: "file:///C:/other.docx", external: true }
			],
			parts: [{ name: "word/glossary/document.xml", contentType: `${WML}.document.glossary+xml`, data: "<w:glossaryDocument/>" }]
		})

		expect(pkg.facts).toMatchObject({
			documentProtection: { edit: "readOnly", enforced: true },
			writeProtection: { recommended: false, hashed: true },
			glossary: true,
			attachedTemplate: true,
			subDocuments: 1
		})
		expect(gateReasons(pkg.facts)).toEqual(["protected", "writeProtected"])
		expect(gateReasons(facts({ documentProtection: { edit: "none", enforced: true, formatting: false, hashed: false } }))).toEqual([])
		expect(gateReasons(facts({ documentProtection: { edit: "comments", enforced: false, formatting: false, hashed: false } }))).toEqual(
			[]
		)
	})

	it("tells encrypted documents from legacy binaries and other compound files", async () => {
		expect(await open(await compoundFile(["EncryptionInfo", "EncryptedPackage"]))).toMatchObject({ ok: false, reason: "encrypted" })
		expect(await open(await compoundFile(["WordDocument", "1Table"]))).toMatchObject({ ok: false, reason: "legacyBinary" })
		expect(await open(await compoundFile(["Something"]))).toMatchObject({ ok: false, reason: "compoundFile" })
		// The package is checked first: Word 2007 and POI leave a WordDocument stream (and an ObjectPool) beside it.
		expect(await open(await compoundFile(["WordDocument", "ObjectPool/x", "EncryptionInfo", "EncryptedPackage"]))).toMatchObject({
			ok: false,
			reason: "encrypted"
		})
		// Rights management describes the package by its DataSpaces alone.
		expect(await open(await compoundFile(["\u0006DataSpaces/DataSpaceMap", "EncryptedPackage"]))).toMatchObject({
			ok: false,
			reason: "encrypted"
		})
		expect(await open(await compoundFile(["EncryptedPackage"]))).toMatchObject({ ok: false, reason: "compoundFile" })
		expect(await open(await compoundFile(["EncryptionInfo/EncryptedPackage", "EncryptedPackage/x"]))).toMatchObject({
			ok: false,
			reason: "compoundFile"
		})

		const truncated = (await compoundFile(["EncryptionInfo"])).subarray(0, 600)

		expect(await open(truncated)).toMatchObject({ ok: false, reason: "compoundFile" })
	})

	it("finds the package in a compound file whose FAT runs past the header's 109 DIFAT slots", async () => {
		// 8 MiB of 512-byte sectors need more FAT sectors than the header lists: the rest come from DIFAT sectors.
		const bytes = await cfbBytes([
			{ path: ["Filler"], bytes: new Uint8Array(8 * 1024 * 1024) },
			{ path: ["EncryptionInfo"], bytes: new Uint8Array(16) },
			{ path: ["EncryptedPackage"], bytes: new Uint8Array(16) }
		])

		expect(new DataView(bytes.buffer).getUint32(0x48, true)).toBeGreaterThan(0)
		expect(await open(bytes)).toMatchObject({ ok: false, reason: "encrypted" })
	})

	it("opens view-only, not failing, when settings do not parse", async () => {
		const bytes = buildDocx({ body: p("x"), settings: "" })
		const files = unzipSync(bytes)

		files["word/settings.xml"] = strToU8("<w:settings><unclosed></w:settings>")

		const opened = await open(zipSync(files))

		if (!opened.ok) throw new Error(opened.reason)

		expect(opened.package.facts.malformed).toBe(true)
		expect(opened.package.settings.present).toBe(false)
		expect(gateReasons(opened.package.facts)).toEqual(["malformed"])
	})

	it("refuses what is not a word-processing package", async () => {
		expect(await open(strToU8("plain text, not a zip at all"))).toMatchObject({ ok: false, reason: "notZip" })

		const docx = buildDocx({ body: p("x") })

		expect(await open(docx.subarray(0, docx.length - 30))).toMatchObject({ ok: false })

		const workbook = buildDocx({
			body: p("x"),
			mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
		})

		expect(await open(workbook)).toMatchObject({ ok: false, reason: "notWordprocessing" })

		const noMain = zipSync({
			"[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
			"_rels/.rels": strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
		})

		expect(await open(noMain)).toMatchObject({ ok: false, reason: "notWordprocessing" })

		const noTypes = zipSync({ "word/document.xml": strToU8("<x/>") })

		expect(await open(noTypes)).toMatchObject({ ok: false, reason: "invalidPackage" })
	})

	it("builds a model even without settings or a body", async () => {
		const pkg = await openSynthetic({
			document: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
		})
		const model = await buildDocxModel(pkg)

		expect(model.settings.present).toBe(false)
		expect(model.body).toBeUndefined()
		expect(model.sections).toEqual([])
	})
})

describe("relationshipKind", () => {
	it("compares transitional, Strict and Microsoft types alike", () => {
		expect(relationshipKind("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")).toBe(
			"officeDocument"
		)
		expect(relationshipKind("http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")).toBe("officeDocument")
		expect(relationshipKind("http://schemas.microsoft.com/office/2011/relationships/commentsExtended")).toBe("commentsExtended")
		expect(relationshipKind("urn:something/else")).toBeUndefined()
	})
})
