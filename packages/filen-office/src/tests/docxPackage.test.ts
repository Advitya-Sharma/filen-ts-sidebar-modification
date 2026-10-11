import { strToU8, unzipSync, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { buildDocxModel, gateReasons, openDocxPackage, relationshipKind, type DocxFacts } from "@filen/office/core/docx"
import { buildDocx, openSynthetic, p } from "./docxSynthetic"

const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"

async function open(bytes: Uint8Array) {
	return openDocxPackage(bytesSource(bytes))
}

// A minimal MS-CFB file (512-byte sectors): header, one directory sector, one FAT sector.
function compoundFile(names: readonly string[]): Uint8Array {
	const bytes = new Uint8Array(512 * 3)
	const view = new DataView(bytes.buffer)

	;[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].forEach((byte, k) => (bytes[k] = byte))
	view.setUint16(0x1a, 0x3e, true)
	view.setUint16(0x1c, 0xfffe, true)
	view.setUint16(0x1e, 9, true)
	view.setUint16(0x20, 6, true)
	view.setUint32(0x2c, 1, true)
	view.setUint32(0x30, 0, true)
	view.setUint32(0x44, 0xfffffffe, true)

	for (let k = 0; k < 109; k++) view.setUint32(0x4c + k * 4, k === 0 ? 1 : 0xffffffff, true)

	const directory = 512
	const entries = ["Root Entry", ...names]

	entries.slice(0, 4).forEach((name, index) => {
		const at = directory + index * 128

		for (let k = 0; k < name.length; k++) view.setUint16(at + k * 2, name.charCodeAt(k), true)

		view.setUint16(at + 0x40, (name.length + 1) * 2, true)
		view.setUint8(at + 0x42, index === 0 ? 5 : 2)
	})

	const fat = 1024

	view.setUint32(fat, 0xfffffffe, true)
	view.setUint32(fat + 4, 0xfffffffd, true)

	for (let k = 2; k < 128; k++) view.setUint32(fat + k * 4, 0xffffffff, true)

	return bytes
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
		expect(await open(compoundFile(["EncryptionInfo", "EncryptedPackage"]))).toMatchObject({ ok: false, reason: "encrypted" })
		expect(await open(compoundFile(["WordDocument", "1Table"]))).toMatchObject({ ok: false, reason: "legacyBinary" })
		expect(await open(compoundFile(["Something"]))).toMatchObject({ ok: false, reason: "compoundFile" })

		const truncated = compoundFile(["EncryptionInfo"]).subarray(0, 600)

		expect(await open(truncated)).toMatchObject({ ok: false, reason: "compoundFile" })
	})

	it("follows DIFAT sectors to find a directory far into a large compound file", async () => {
		const sectorSize = 512
		const directorySector = 14000
		const difatSector = 200
		const fatSector = 201
		const sectors = new Map<number, Uint8Array>()
		const header = new Uint8Array(sectorSize)
		const headerView = new DataView(header.buffer)

		;[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].forEach((byte, k) => (header[k] = byte))
		headerView.setUint16(0x1e, 9, true)
		headerView.setUint32(0x30, directorySector, true)
		headerView.setUint32(0x44, difatSector, true)
		headerView.setUint32(0x48, 1, true)

		for (let k = 0; k < 109; k++) headerView.setUint32(0x4c + k * 4, 0xffffffff, true)

		const difat = new Uint8Array(sectorSize)

		new DataView(difat.buffer).setUint32(0, fatSector, true)
		new DataView(difat.buffer).setUint32(sectorSize - 4, 0xfffffffe, true)
		sectors.set(difatSector, difat)

		const fat = new Uint8Array(sectorSize)

		// The directory spans two sectors; the stream names sit in the second, reached only through the FAT.
		new DataView(fat.buffer).setUint32((directorySector % 128) * 4, directorySector + 1, true)
		new DataView(fat.buffer).setUint32(((directorySector + 1) % 128) * 4, 0xfffffffe, true)
		sectors.set(fatSector, fat)
		sectors.set(directorySector, compoundFile(["Filler1", "Filler2", "Filler3"]).subarray(512, 1024))
		sectors.set(directorySector + 1, compoundFile(["EncryptionInfo", "EncryptedPackage", "Filler4"]).subarray(512, 1024))

		const size = (directorySector + 3) * sectorSize
		const source = {
			size,
			read(offset: number, length: number): Promise<Uint8Array> {
				if (offset === 0) return Promise.resolve(header.subarray(0, length))

				const sector = offset / sectorSize - 1

				return Promise.resolve(sectors.get(sector)?.subarray(0, length) ?? new Uint8Array(length))
			}
		}

		expect(await openDocxPackage(source)).toMatchObject({ ok: false, reason: "encrypted" })
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
