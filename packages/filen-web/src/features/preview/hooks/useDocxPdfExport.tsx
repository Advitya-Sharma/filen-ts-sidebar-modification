import { useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import { pdfPasswordAllowed } from "@/features/office/lib/pdfPassword"
import { DocxPasswordPrompt } from "@/features/preview/components/docxPasswordPrompt"
import { usePreviewDownload } from "@/features/preview/lib/accessMode"
import { downloadDocxPdf, unlockForPdf, type DocxPdfSource } from "@/features/preview/lib/docxPdf"

export interface DocxPdfExportTarget {
	readonly fileName: string
	// What is exported: the viewer's file, or the editor's layout.
	readonly source: () => DocxPdfSource
	// A password-protected file as it is stored: its password is asked and checked before the export, and the
	// PDF is encrypted with it. A viewer exports the document that check unlocked; an editor its own layout.
	readonly encryptedFile?: Uint8Array | undefined
}

// Download as ▸ PDF for a Word document: the menu's action, and the password prompt a protected document asks
// first (render `prompt` beside the menu). Offered only where the preview offers the file's own Download (the
// overlay's rule: never in the trash, never for a file that may not be downloaded); `target` null: not at all.
export function useDocxPdfExport(target: DocxPdfExportTarget | null): {
	readonly onDownloadPdf: (() => void) | undefined
	readonly prompt: ReactNode
} {
	const { t } = useTranslation("preview")
	const [asking, setAsking] = useState(false)
	const download = usePreviewDownload()
	const encrypted = target?.encryptedFile

	const onDownloadPdf =
		target !== null && download !== null
			? () => {
					if (encrypted === undefined) {
						void downloadDocxPdf(target.source(), target.fileName)
					} else {
						setAsking(true)
					}
				}
			: undefined

	const prompt =
		target === null || encrypted === undefined ? null : (
			<DocxPasswordPrompt
				open={asking}
				title={t("previewOfficePdfPasswordTitle")}
				body={t("previewOfficePdfPasswordBody")}
				wrongMessage={t("previewOfficePdfPasswordWrong")}
				label={t("previewOfficePdfPasswordLabel")}
				submitLabel={t("previewOfficePdfPasswordSubmit")}
				progressLabel={t("previewOfficePdfPasswordChecking")}
				onOpenChange={setAsking}
				verify={async (password, signal) => {
					// The PDF is encrypted with this password: one it cannot take is refused before any work starts.
					if (!pdfPasswordAllowed(password)) {
						setAsking(false)
						toast.error(t("previewDocxPdfPasswordUnsupported"))

						return "aborted"
					}

					const unlocked = await unlockForPdf(encrypted, password, signal).catch(() => ({ verdict: "failed" as const }))

					if (unlocked.verdict !== "accepted") {
						return unlocked.verdict === "wrong" ? "wrong" : "aborted"
					}

					// Cancelled while the check finished: nothing is exported.
					if (signal.aborted) {
						unlocked.document.release()

						return "aborted"
					}

					// The prompt closes as the export starts (its toast shows the progress); one export per password.
					setAsking(false)

					const source = target.source()

					if (source.kind === "editor") {
						unlocked.document.release()
						void downloadDocxPdf(source, target.fileName, password)
					} else {
						void downloadDocxPdf({ kind: "unlocked", document: unlocked.document }, target.fileName, password)
					}

					return "accepted"
				}}
			/>
		)

	return { onDownloadPdf, prompt }
}
