import type { SheetAbortSignal, SheetSessionApi } from "../../../core/sheet/protocol"
import { SheetSessionClosedError, sheetJobs } from "../../../dom/sheet/client"
import type { SheetPdfRefusal, SheetPdfRequest } from "./pdfExport"

// The page's side of print-to-PDF: the job runs in the sheet worker holding the session (pdfJob.ts), so neither the
// PDF writer nor its fonts load here. A session that does not run in a worker cannot make one ("unsupported"); one
// closed before its PDF was made answers "closed" (the user left; nothing to report).

export const SHEET_PDF_JOB = "print.pdf"

export type SheetPdfJobResult =
	| {
			readonly ok: true
			readonly blob: Blob
			readonly pages: number
			readonly missingImages: number
			readonly missingGlyphs: number
			readonly unshapedRuns: number
	  }
	| { readonly ok: false; readonly reason: SheetPdfRefusal; readonly pages?: number }

export type { SheetPdfRefusal, SheetPdfRequest }

// What the job is sent: the print, and the password that encrypts the PDF (a protected workbook's own).
export interface SheetPdfJobArgs {
	readonly request: SheetPdfRequest
	readonly password?: string
}

export async function sheetPdf(
	session: SheetSessionApi,
	request: SheetPdfRequest,
	options: {
		readonly signal?: SheetAbortSignal
		readonly onProgress?: (done: number, total: number) => void
		readonly password?: string
	} = {}
): Promise<SheetPdfJobResult | { readonly ok: false; readonly reason: "unsupported" | "closed" }> {
	const jobs = sheetJobs(session)

	if (jobs === null) return { ok: false, reason: "unsupported" }

	const onProgress = options.onProgress

	const args: SheetPdfJobArgs = options.password === undefined ? { request } : { request, password: options.password }

	try {
		return (await jobs.job(SHEET_PDF_JOB, args, {
			...(options.signal !== undefined && { signal: options.signal }),
			...(onProgress !== undefined && {
				onProgress: (value: unknown) => {
					const { done, total } = value as { done: number; total: number }

					onProgress(done, total)
				}
			})
		})) as SheetPdfJobResult
	} catch (error) {
		if (error instanceof SheetSessionClosedError) return { ok: false, reason: "closed" }

		throw error
	}
}
