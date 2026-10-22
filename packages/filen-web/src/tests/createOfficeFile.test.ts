import { describe, expect, it } from "vitest"
import { OpcPackage, bytesSource } from "@filen/office/core/opc"
import { OFFICE_FILE_MIME, freeOfficeFileName, normalizeOfficeFileName, officeFileBytes } from "@/features/drive/lib/createOfficeFile"

async function part(bytes: Uint8Array, name: string): Promise<string> {
	return new TextDecoder().decode(await (await OpcPackage.open(bytesSource(bytes))).read(name))
}

describe("normalizeOfficeFileName", () => {
	it("adds the extension when the name has none, or another one", () => {
		expect(normalizeOfficeFileName("Report", "docx")).toBe("Report.docx")
		expect(normalizeOfficeFileName("Report.txt", "docx")).toBe("Report.txt.docx")
		expect(normalizeOfficeFileName(".budget", "xlsx")).toBe(".budget.xlsx")
	})

	it("keeps the extension the name already ends with, in any case", () => {
		expect(normalizeOfficeFileName("Report.docx", "docx")).toBe("Report.docx")
		expect(normalizeOfficeFileName("Budget.XLSX", "xlsx")).toBe("Budget.XLSX")
	})
})

describe("freeOfficeFileName", () => {
	it("is the plain name while it is free", () => {
		expect(freeOfficeFileName("Untitled document", "docx", new Set())).toBe("Untitled document.docx")
	})

	it("numbers the name past every taken one, ignoring letter case", () => {
		const taken = new Set(["untitled document.docx", "untitled document (2).docx"])

		expect(freeOfficeFileName("Untitled document", "docx", taken)).toBe("Untitled document (3).docx")
	})
})

describe("officeFileBytes", () => {
	it("builds a blank document in the user's language and paper", async () => {
		const bytes = await officeFileBytes("docx", { locale: "de-DE", sheetName: "Sheet1" })

		expect(await part(bytes, "word/styles.xml")).toContain('<w:lang w:val="de-DE"/>')
		expect(await part(bytes, "word/document.xml")).toContain('<w:pgSz w:w="11906" w:h="16838"/>')
	})

	it("builds a blank workbook whose sheet has the localized name", async () => {
		const bytes = await officeFileBytes("xlsx", { locale: "en-US", sheetName: "Tabelle1" })

		expect(await part(bytes, "xl/workbook.xml")).toContain('<sheet name="Tabelle1" sheetId="1" r:id="rId1"/>')
	})

	it("uploads each kind under its Office Open XML type", () => {
		expect(OFFICE_FILE_MIME).toEqual({
			docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		})
	})
})
