import { describe, expect, it } from "vitest"
import { pdfDownloadName, pdfErrorKey, pdfProgressText, pdfReportKey, pdfTitle } from "@/features/preview/lib/docxPdf.logic"

describe("Download as PDF", () => {
	it("names the download after the document, with the PDF extension", () => {
		expect(pdfDownloadName("Quarterly report.docx")).toBe("Quarterly report.pdf")
		expect(pdfDownloadName("notes.v2.odt")).toBe("notes.v2.pdf")
		expect(pdfDownloadName("README")).toBe("README.pdf")
		expect(pdfTitle("Quarterly report.docx")).toBe("Quarterly report")
	})

	it("says preparing until pages are counted, then the page being written", () => {
		expect(pdfProgressText(null)).toEqual({ key: "previewDocxPdfPreparing" })
		expect(pdfProgressText({ phase: "layout", done: 3, total: 0 })).toEqual({ key: "previewDocxPdfPreparing" })
		expect(pdfProgressText({ phase: "pages", done: 0, total: 12 })).toEqual({ key: "previewDocxPdfProgress", done: 1, total: 12 })
		expect(pdfProgressText({ phase: "pages", done: 12, total: 12 })).toEqual({ key: "previewDocxPdfProgress", done: 12, total: 12 })
	})

	it("tells about characters no font has before simplified content, and nothing about refused links", () => {
		expect(pdfReportKey({ missing: [], refused: [] })).toBeNull()
		expect(pdfReportKey({ missing: [], refused: [["link", 2]] })).toBeNull()
		expect(pdfReportKey({ missing: [["Zzzz", 1]], refused: [["image", 1]] })).toBe("previewDocxPdfMissingGlyphs")
		expect(pdfReportKey({ missing: [], refused: [["image", 1]] })).toBe("previewDocxPdfSimplified")
	})

	it("tells about text drawn without its font's shaping", () => {
		expect(pdfReportKey({ missing: [], refused: [["shaping", 1]] })).toBe("previewDocxPdfUnshaped")
		expect(
			pdfReportKey({
				missing: [],
				refused: [
					["shaping", 1],
					["image", 1]
				]
			})
		).toBe("previewDocxPdfSimplified")
	})

	it("says why a password the PDF cannot take fails the export", () => {
		expect(pdfErrorKey("passwordUnsupported")).toBe("previewDocxPdfPasswordUnsupported")
	})

	it("says nothing after the user's own cancel", () => {
		expect(pdfErrorKey("aborted")).toBeNull()
		expect(pdfErrorKey("failed")).toBe("previewDocxPdfFailed")
		expect(pdfErrorKey("workerFailed")).toBe("previewDocxPdfFailed")
	})
})
