import { toast } from "sonner"
import type { SheetSessionApi } from "@filen/office/core/sheet"
import { MAX_PRINT_PAGES } from "@filen/office/core/sheet/print"
import { sheetPdf, type SheetPdfRequest } from "@filen/office/react/sheet/print"
import { fileStem } from "@/features/spreadsheet/save/sheetSave.logic"
import { downloadBlob } from "@/lib/downloadBlob"
import { i18n } from "@/lib/i18n"
import { log } from "@/lib/log"

// Download as PDF for a spreadsheet: the print made in the sheet worker beside the session (cells, objects and
// charts never cross to the page), with a toast that counts the pages and can cancel. A local download: the file in
// the drive is unchanged.

let exports = 0

export function pdfDownloadName(fileName: string): string {
	return `${fileStem(fileName)}.pdf`
}

// `password` (a protected workbook's, checked first) encrypts the PDF; it goes with this export only.
export async function downloadSheetPdf(
	session: SheetSessionApi,
	request: SheetPdfRequest,
	fileName: string,
	password?: string
): Promise<void> {
	const id = `sheet-pdf-${String(++exports)}`
	const controller = new AbortController()
	const action = {
		label: i18n.t("preview:previewSheetPdfCancel"),
		onClick: () => {
			controller.abort()
		}
	}
	const show = (done: number | null, total: number): void => {
		toast.loading(
			done === null ? i18n.t("preview:previewSheetPdfPreparing") : i18n.t("preview:previewSheetPdfProgress", { done, total }),
			{ id, action, duration: Number.POSITIVE_INFINITY }
		)
	}

	show(null, 0)

	try {
		const result = await sheetPdf(session, request, {
			signal: controller.signal,
			onProgress: show,
			...(password !== undefined && { password })
		})

		toast.dismiss(id)

		if (result.ok) {
			downloadBlob(pdfDownloadName(fileName), result.blob)

			if (result.missingGlyphs > 0 || result.missingImages > 0) {
				log.warn("preview", "spreadsheet pdf simplified", { glyphs: result.missingGlyphs, images: result.missingImages })
				toast.warning(i18n.t("preview:previewSheetPdfSimplified"))
			}

			if (result.unshapedRuns > 0) {
				log.warn("preview", "spreadsheet pdf unshaped", { runs: result.unshapedRuns })
				toast.warning(i18n.t("preview:previewSheetPdfUnshaped"))
			}

			return
		}

		switch (result.reason) {
			// The viewer was closed while the PDF was being made: the user left, nothing failed.
			case "closed":
				break
			case "nothing":
				toast.error(i18n.t("preview:previewSheetPdfNothing"))
				break
			case "tooManyPages":
				toast.error(i18n.t("preview:previewSheetPdfTooMany", { pages: MAX_PRINT_PAGES }))
				break
			case "passwordUnsupported":
				toast.error(i18n.t("preview:previewSheetPdfPasswordUnsupported"))
				break
			case "pageTooLarge":
				toast.error(i18n.t("preview:previewSheetPdfPageTooLarge", { page: result.pages ?? 1 }))
				break
			default:
				toast.error(i18n.t("preview:previewSheetPdfFailed"))
		}
	} catch (error) {
		toast.dismiss(id)

		// The user's own Cancel says nothing.
		if (controller.signal.aborted) {
			return
		}

		log.error("preview", "spreadsheet pdf failed", error)
		toast.error(i18n.t("preview:previewSheetPdfFailed"))
	}
}
