import type { PdfFontAssets } from "@filen/office/core/pdfwrite/fonts"
import { sheetPasswordJob, type SheetWorkerJob } from "@filen/office/dom/sheet"
import { sheetPdfJob } from "@filen/office/react/sheet/print"

// The workbook worker's jobs: a print's PDF, and the password check a protected workbook's PDF export asks first.
export function sheetWorkbookJobs(fonts: PdfFontAssets): Readonly<Record<string, SheetWorkerJob>> {
	return { ...sheetPdfJob(fonts), ...sheetPasswordJob() }
}
