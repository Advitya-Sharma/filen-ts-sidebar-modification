import { useState } from "react"
import { useTranslation } from "react-i18next"
import { LockIcon } from "lucide-react"
import type { DocxWriteProtection } from "@filen/office/core/docx/edit"
import { verifyProtectionPassword } from "@filen/office/dom"
import { DocxPasswordPrompt, type PasswordVerdict } from "@/features/preview/components/docxPasswordPrompt"
import { Button } from "@/components/ui/button"

// Word's "password to modify" and "read-only recommended" (w:writeProtection): the document opens view-only
// under a bar saying why, as Word asks on open. "Edit…" asks for the password and checks it here, in slices
// (the page stays responsive; Cancel stops it); "Edit anyway" takes a recommendation's way out. Either unlocks
// editing for this session only (onUnlock); the file keeps its protection. The password is never logged or
// kept.

export interface DocxWriteProtectionBarProps {
	protection: DocxWriteProtection
	// The file's name, for the prompt.
	name: string
	onUnlock: () => void
}

export function DocxWriteProtectionBar({ protection, name, onUnlock }: DocxWriteProtectionBarProps) {
	const { t } = useTranslation("preview")
	const [prompt, setPrompt] = useState(false)
	// The hash turned out not to be checkable here (not base64, the wrong size): only viewing is left.
	const [unchecked, setUnchecked] = useState(false)
	const password = protection.password
	const unsupported = password === "unsupported" || unchecked

	const verify = async (value: string, signal: AbortSignal): Promise<PasswordVerdict> => {
		if (password === undefined || password === "unsupported") {
			return "aborted"
		}

		try {
			const verdict = await verifyProtectionPassword(password, value, { signal })

			if (verdict === "correct") {
				onUnlock()

				return "accepted"
			}

			if (verdict === "unsupported") {
				setUnchecked(true)
				setPrompt(false)

				return "aborted"
			}

			return "wrong"
		} catch {
			// Cancelled.
			return "aborted"
		}
	}

	return (
		<div
			role="status"
			className="flex shrink-0 items-center gap-2 border-b bg-muted px-3 py-1.5 text-xs text-muted-foreground"
		>
			<LockIcon
				aria-hidden="true"
				className="size-3.5 shrink-0"
			/>
			<p className="min-w-0 flex-1">
				{t(
					password === undefined
						? "previewDocxWriteRecommended"
						: unsupported
							? "previewDocxWritePasswordUnsupported"
							: "previewDocxWritePassword"
				)}
			</p>
			{password === undefined ? (
				<Button
					variant="outline"
					size="xs"
					onClick={onUnlock}
				>
					{t("previewDocxWriteEditAnyway")}
				</Button>
			) : unsupported ? null : (
				<>
					<Button
						variant="outline"
						size="xs"
						onClick={() => {
							setPrompt(true)
						}}
					>
						{t("previewDocxWriteEnterPassword")}
					</Button>
					<DocxPasswordPrompt
						open={prompt}
						title={t("previewDocxWritePasswordTitle")}
						body={t("previewDocxWritePasswordBody", { name })}
						wrongMessage={t("previewDocxWritePasswordWrong")}
						label={t("previewDocxWritePasswordLabel")}
						submitLabel={t("previewDocxWritePasswordSubmit")}
						progressLabel={t("previewDocxWritePasswordChecking")}
						verify={verify}
						onOpenChange={setPrompt}
					/>
				</>
			)}
		</div>
	)
}
