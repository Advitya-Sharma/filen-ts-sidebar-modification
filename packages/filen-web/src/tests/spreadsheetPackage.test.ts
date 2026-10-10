import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import JSZip from "jszip"
import { openXlsx, saveXlsx } from "hucre/xlsx"
import { spreadsheetFileKind, spreadsheetSaveFormat } from "@/features/spreadsheet/lib/fileKind"
import { XlsxDocument } from "@/features/spreadsheet/lib/xlsxDocument"
import { saveLosses } from "@/features/spreadsheet/lib/xlsxVerify"
import { rawEntries, xlsxSavePlan } from "@/features/spreadsheet/lib/xlsxWritable"
import { readZip } from "@/features/spreadsheet/lib/zipLimits"

const FIXTURES = new URL("./fixtures/spreadsheet/", import.meta.url)
const decoder = new TextDecoder()
const TEMPLATE = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
const MACRO_TEMPLATE = "application/vnd.ms-excel.template.macroEnabled.main+xml"
const SIGNATURE = "application/vnd.ms-office.vbaProjectSignature"
const VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
const SIGNATURE_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"

// An Excel file rebuilt with `change` made to its parts (by path), as another program might have written it.
async function rebuilt(name: string, change: (parts: Map<string, string | Uint8Array>) => void): Promise<Uint8Array> {
	const parts = new Map<string, string | Uint8Array>((await readZip(new Uint8Array(readFileSync(new URL(name, FIXTURES))))) ?? [])

	change(parts)

	const zip = new JSZip()

	for (const [path, data] of parts) zip.file(path, data)

	return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" })
}

function text(parts: Map<string, string | Uint8Array>, path: string): string {
	const data = parts.get(path)

	return typeof data === "string" ? data : decoder.decode(data)
}

// A template workbook, or with macros (and a signature over them) a macro-enabled template.
async function template(macros: boolean): Promise<Uint8Array> {
	return await rebuilt("x365_plain.xlsx", parts => {
		const types = text(parts, "[Content_Types].xml")
		const main = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

		if (!macros) {
			parts.set("[Content_Types].xml", types.replace(main, TEMPLATE))

			return
		}

		parts.set(
			"[Content_Types].xml",
			types
				.replace(main, MACRO_TEMPLATE)
				.replace(
					"</Types>",
					`<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/vbaProjectSignature.bin" ContentType="${SIGNATURE}"/></Types>`
				)
		)
		parts.set(
			"xl/_rels/workbook.xml.rels",
			text(parts, "xl/_rels/workbook.xml.rels").replace(
				"</Relationships>",
				`<Relationship Id="rIdVba" Type="${VBA_REL}" Target="vbaProject.bin"/></Relationships>`
			)
		)
		parts.set("xl/vbaProject.bin", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]))
		parts.set("xl/vbaProjectSignature.bin", new Uint8Array([4, 5, 6]))
		parts.set(
			"xl/_rels/vbaProject.bin.rels",
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${SIGNATURE_REL}" Target="vbaProjectSignature.bin"/></Relationships>`
		)
	})
}

async function editedAndSaved(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
	const document = new XlsxDocument(await openXlsx(bytes, { readStyles: true }))

	expect([await document.verifyWritable(), document.losses]).toEqual([true, []])
	document.apply({ type: "setCells", sheet: 0, cells: [{ row: 0, col: 0, input: "edited" }] })

	return (await readZip((await document.serialize()).bytes)) ?? new Map()
}

describe("spreadsheet templates", () => {
	it("opens .xltx and .xltm as workbooks, telling them apart from .xlsx and .xlsm", () => {
		expect(spreadsheetFileKind("xltx")).toBe("xlsx")
		expect(spreadsheetFileKind("xltm")).toBe("xlsx")
		expect(new Set(["xlsx", "xlsm", "xltx", "xltm"].map(spreadsheetSaveFormat)).size).toBe(4)
	})

	it("saves a template as a template", async () => {
		const saved = await editedAndSaved(await template(false))

		expect(decoder.decode(saved.get("[Content_Types].xml"))).toContain(`ContentType="${TEMPLATE}"`)
	})

	it("saves a macro-enabled template with its macros, their signature and the content types they had", async () => {
		const bytes = await template(true)
		const saved = await editedAndSaved(bytes)
		const types = decoder.decode(saved.get("[Content_Types].xml"))

		expect(types).toContain(`ContentType="${MACRO_TEMPLATE}"`)
		expect(types).toContain(`<Override PartName="/xl/vbaProjectSignature.bin" ContentType="${SIGNATURE}"/>`)
		expect(saved.get("xl/vbaProjectSignature.bin")).toEqual(new Uint8Array([4, 5, 6]))
		expect(saved.get("xl/vbaProject.bin")).toEqual(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]))
	})
})

describe("proof of content types", () => {
	it("finds a part saved with another content type than the file gave it, or none", async () => {
		const workbook = await openXlsx(await template(true), { readStyles: true })
		const plan = xlsxSavePlan(workbook)
		const original = new Map(rawEntries(workbook))
		const saved = (await readZip(await saveXlsx(workbook))) ?? new Map<string, Uint8Array>()
		const types = decoder.decode(saved.get("[Content_Types].xml"))
		const losses = async (edit: (xml: string) => string) =>
			(await saveLosses({
				original,
				saved: new Map([...saved, ["[Content_Types].xml", new TextEncoder().encode(edit(types))]]),
				sheetPaths: plan.sheets,
				dropped: plan.drop
			})) ?? []

		expect(await losses(xml => xml)).toEqual([])
		expect(await losses(xml => xml.replace(/<Override PartName="\/xl\/vbaProjectSignature\.bin"[^>]*\/>/, ""))).toEqual([
			`[Content_Types].xml: xl/vbaprojectsignature.bin: content type ${SIGNATURE} saved as application/vnd.ms-office.vbaProject`
		])
		expect(
			(await losses(xml => xml.replace(MACRO_TEMPLATE, "application/vnd.ms-excel.sheet.macroEnabled.main+xml"))).some(loss =>
				loss.includes("xl/workbook.xml: content type")
			)
		).toBe(true)
		expect(
			(await losses(xml => xml.replace(/<Default Extension="bin"[^>]*\/>/, ""))).some(loss =>
				loss.includes("xl/vbaproject.bin: saved with no content type")
			)
		).toBe(true)
	})
})

// Each entry's compression method (0 stored, 8 deflated), from the zip's central directory.
function compressionMethods(bytes: Uint8Array): Map<string, number> {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const end = bytes.length - 22
	const methods = new Map<string, number>()
	let offset = view.getUint32(end + 16, true)

	for (let entry = 0; entry < view.getUint16(end + 10, true); entry++) {
		const nameLength = view.getUint16(offset + 28, true)

		methods.set(decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength)), view.getUint16(offset + 10, true))
		offset += 46 + nameLength + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true)
	}

	return methods
}

describe("content types written once each", () => {
	it("replaces the type saving writes for a part it copies with the file's, rather than declare it twice", async () => {
		const own = "application/vnd.example.theme+xml"
		const bytes = await rebuilt("x365_plain.xlsx", parts => {
			parts.set(
				"[Content_Types].xml",
				text(parts, "[Content_Types].xml").replace(/(PartName="\/xl\/theme\/theme1\.xml" ContentType=")[^"]+"/, `$1${own}"`)
			)
		})
		const types = decoder.decode((await editedAndSaved(bytes)).get("[Content_Types].xml"))

		expect(types.match(/PartName="\/xl\/theme\/theme1\.xml"/g)).toHaveLength(1)
		expect(types).toContain(`ContentType="${own}"`)
	})

	it("finds a part whose content type is declared twice", async () => {
		const workbook = await openXlsx(await template(false), { readStyles: true })
		const plan = xlsxSavePlan(workbook)
		const saved = (await readZip(await saveXlsx(workbook))) ?? new Map<string, Uint8Array>()
		const types = decoder.decode(saved.get("[Content_Types].xml"))
		const twice = types.replace("</Types>", '<Override PartName="/xl/styles.xml" ContentType="application/xml"/></Types>')

		expect(
			await saveLosses({
				original: new Map(rawEntries(workbook)),
				saved: new Map([...saved, ["[Content_Types].xml", new TextEncoder().encode(twice)]]),
				sheetPaths: plan.sheets,
				dropped: plan.drop
			})
		).toContain("[Content_Types].xml: xl/styles.xml: content type declared more than once")
	})
})

describe("saved package size", () => {
	it("compresses the parts copied as the file had them, as small as it was", async () => {
		const original = new Uint8Array(readFileSync(new URL("x365_plain.xlsx", FIXTURES)))
		const document = new XlsxDocument(await openXlsx(original, { readStyles: true }))

		await document.verifyWritable()

		const { bytes } = await document.serialize()

		expect(compressionMethods(bytes).get("xl/theme/theme1.xml")).toBe(8)
		expect(bytes.length).toBeLessThan(original.length * 1.1)
	})
})
