import { beforeEach, describe, expect, it, vi } from "vitest"
import type { DocxPdfExportRequest } from "@filen/office/dom/docx/pdfClient"
import { makeDocxPdf } from "@/features/preview/lib/docxPdf"

// What a file shows that the document exported leaves out (an RTF's equations) goes to the PDF worker with the export,
// which reports it, for the file as stored and for the editor's document alike.

const requests = vi.hoisted((): DocxPdfExportRequest[] => [])

vi.mock("@filen/office/dom/docx/pdfClient", async original => ({
	...(await original<typeof import("@filen/office/dom/docx/pdfClient")>()),
	docxPdfSnapshot: () => Promise.resolve({ pages: [] }),
	exportDocxPdfInWorker: (_worker: unknown, request: DocxPdfExportRequest) => {
		requests.push(request)

		return Promise.resolve({ blob: new Blob(), report: { missing: [], refused: [] } })
	}
}))

beforeEach(() => {
	requests.length = 0
})

const make = { signal: new AbortController().signal, onProgress: () => undefined }

describe("the PDF of a document whose file shows more than it holds", () => {
	it("hands the worker the count of what the file's package leaves out", async () => {
		await makeDocxPdf({ kind: "file", bytes: new Uint8Array(4), owned: true, untranslated: 2 }, "a.rtf", make)
		await makeDocxPdf({ kind: "editor", live: () => Promise.resolve({} as never), markup: false, untranslated: 3 }, "a.rtf", make)
		await makeDocxPdf({ kind: "file", bytes: new Uint8Array(4), owned: true }, "a.docx", make)

		expect(requests.map(request => request.options.untranslated)).toEqual([2, 3, undefined])
	})
})
