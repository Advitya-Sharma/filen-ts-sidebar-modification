import { useState } from "react"
import { useTranslation } from "react-i18next"
import { DocxPasswordPrompt } from "@/features/preview/components/docxPasswordPrompt"
import { usePasswordAsk } from "@/features/drive/lib/downloadAsPassword"

// The password prompt of a protected file's Download as (lib/downloadAs.ts), mounted once at the root beside the
// toasts: the menu that started the export is closed by then. The shared PDF password words, with the file named
// (a bulk export asks file by file).
export function DownloadAsPasswordDialog() {
	const { t } = useTranslation(["preview", "drive"])
	const ask = usePasswordAsk()
	// The last file asked about keeps its name through the exit animation.
	const [shown, setShown] = useState(ask)

	if (ask !== null && ask !== shown) {
		setShown(ask)
	}

	const name = (ask ?? shown)?.name ?? ""

	return (
		<DocxPasswordPrompt
			open={ask !== null}
			title={t("preview:previewOfficePdfPasswordTitle")}
			body={t("drive:driveDownloadAsPasswordBody", { name })}
			wrongMessage={t("preview:previewOfficePdfPasswordWrong")}
			label={t("preview:previewOfficePdfPasswordLabel")}
			submitLabel={t("preview:previewOfficePdfPasswordSubmit")}
			progressLabel={t("preview:previewOfficePdfPasswordChecking")}
			onOpenChange={open => {
				if (!open) ask?.settle(false)
			}}
			verify={(password, signal) => (ask === null ? Promise.resolve("aborted") : ask.verify(password, signal))}
		/>
	)
}
