import { useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import type { SheetSessionApi } from "@filen/office/core/sheet"
import { checkSheetPassword } from "@filen/office/dom/sheet"
import type { SheetPdfRequest } from "@filen/office/react/sheet/print"
import { pdfPasswordAllowed } from "@/features/office/lib/pdfPassword"
import { DocxPasswordPrompt } from "@/features/preview/components/docxPasswordPrompt"
import { downloadSheetPdf } from "@/features/spreadsheet/save/sheetPdf"
import { log } from "@/lib/log"

// Download as ▸ PDF for a spreadsheet, and the password prompt a protected workbook asks first (render `prompt` beside
// the menu), as a Word document's: the password is checked in the worker against the file as stored, then goes with
// that one export (the PDF is encrypted with it) and is kept nowhere. A workbook Excel's default password opens is
// not asked about, and its PDF is not protected. `start` takes the print to export (null when nothing prints);
// undefined when nothing may be downloaded.
export type SheetPrintRequest = () => Promise<SheetPdfRequest | null>

export function useSheetPdfExport({
	session,
	fileName,
	encryptedFile,
	offered
}: {
	session: SheetSessionApi
	fileName: string
	// A password-protected workbook as stored.
	encryptedFile: Uint8Array | undefined
	offered: boolean
}): { readonly start: ((request: SheetPrintRequest) => void) | undefined; readonly prompt: ReactNode } {
	const { t } = useTranslation("preview")
	// The print the prompt is asked for.
	const [asking, setAsking] = useState<SheetPrintRequest | null>(null)

	const exportWith = (request: SheetPrintRequest, password: string | undefined): void => {
		void request().then(
			print => {
				if (print === null) {
					toast.error(t("previewSheetPdfNothing"))

					return undefined
				}

				return downloadSheetPdf(session, print, fileName, password)
			},
			(error: unknown) => {
				log.warn("preview", "spreadsheet pdf settings failed", error)
				toast.error(t("previewSheetPdfFailed"))
			}
		)
	}

	const start = !offered
		? undefined
		: (request: SheetPrintRequest) => {
				if (encryptedFile === undefined) {
					exportWith(request, undefined)

					return
				}

				// Opened with Excel's default password: nothing to ask, nothing to protect.
				void checkSheetPassword(session, encryptedFile, "").then(
					verdict => {
						if (verdict === "default") exportWith(request, undefined)
						else setAsking(() => request)
					},
					() => {
						setAsking(() => request)
					}
				)
			}

	const prompt =
		encryptedFile === undefined ? null : (
			<DocxPasswordPrompt
				open={asking !== null}
				title={t("previewOfficePdfPasswordTitle")}
				body={t("previewOfficePdfPasswordBody")}
				wrongMessage={t("previewOfficePdfPasswordWrong")}
				label={t("previewOfficePdfPasswordLabel")}
				submitLabel={t("previewOfficePdfPasswordSubmit")}
				progressLabel={t("previewOfficePdfPasswordChecking")}
				onOpenChange={open => {
					if (!open) setAsking(null)
				}}
				verify={async (password, signal) => {
					// The PDF is encrypted with this password: one it cannot take is refused before the worker checks it.
					if (!pdfPasswordAllowed(password)) {
						setAsking(null)
						toast.error(t("previewSheetPdfPasswordUnsupported"))

						return "aborted"
					}

					const verdict = await checkSheetPassword(session, encryptedFile, password, signal).catch(() => "failed" as const)

					if (verdict === "wrong") return "wrong"

					if (verdict === "failed") {
						toast.error(t("previewSheetPdfFailed"))
					}

					if (verdict !== "accepted" || asking === null || signal.aborted) return "aborted"

					// The prompt closes as the export starts (its toast shows the progress); one export per password.
					setAsking(null)
					exportWith(asking, password)

					return "accepted"
				}}
			/>
		)

	return { start, prompt }
}
