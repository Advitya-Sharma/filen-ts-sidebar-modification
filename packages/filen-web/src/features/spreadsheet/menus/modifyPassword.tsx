import { useState } from "react"
import { useTranslation } from "react-i18next"
import type { SheetSessionApi } from "@filen/office/core/sheet"
import { sheetPasswordCheck } from "@filen/office/core/xlsx/edit"
import { verifyProtectionPassword, type ProtectionVerdict } from "@filen/office/dom"
import { Button } from "@/components/ui/button"
import { DocxPasswordPrompt, type PasswordVerdict } from "@/features/preview/components/docxPasswordPrompt"

// A workbook's password to modify (§18.2.12 fileSharing): it opens view-only, and "Edit…" asks for the password,
// checked here against the file's hash (the legacy 16-bit verifier at once, the modern one in slices on the
// platform). The right one unlocks editing for this session only (onUnlock); the file keeps its protection. The
// password is never logged or kept.

// The password checked against the file's hash.
async function checkPassword(session: SheetSessionApi, value: string, signal: AbortSignal): Promise<ProtectionVerdict> {
	const check = sheetPasswordCheck(await session.query("protection.hash", { target: "fileSharing" }), value)

	if (check.kind === "none") {
		return "correct"
	}

	if (check.kind === "legacy") {
		return check.correct ? "correct" : "wrong"
	}

	const algorithm =
		check.algorithm === "SHA-1" || check.algorithm === "SHA-256" || check.algorithm === "SHA-384" || check.algorithm === "SHA-512"
			? check.algorithm
			: null

	if (algorithm === null) {
		return "unsupported"
	}

	return await verifyProtectionPassword(
		{ algorithm, hash: check.hash, salt: check.salt, spinCount: check.spinCount, input: "plain" },
		value,
		{
			signal
		}
	)
}

// Cancelled, or the session went: nothing was decided.
function settled(session: SheetSessionApi, value: string, signal: AbortSignal): Promise<ProtectionVerdict | "aborted"> {
	return checkPassword(session, value, signal).catch(() => "aborted" as const)
}

export function SheetModifyPassword({ session, name, onUnlock }: { session: SheetSessionApi; name: string; onUnlock: () => void }) {
	const { t } = useTranslation("preview")
	const [prompt, setPrompt] = useState(false)
	const [unsupported, setUnsupported] = useState(false)

	const verify = async (value: string, signal: AbortSignal): Promise<PasswordVerdict> => {
		const verdict = await settled(session, value, signal)

		if (verdict === "correct") {
			onUnlock()

			return "accepted"
		}

		if (verdict === "unsupported") {
			setUnsupported(true)
			setPrompt(false)

			return "aborted"
		}

		return verdict
	}

	if (unsupported) {
		return (
			<span
				role="status"
				className="shrink truncate text-xs text-muted-foreground"
			>
				{t("previewSheetModifyPasswordUnsupported")}
			</span>
		)
	}

	return (
		<>
			<Button
				variant="outline"
				size="xs"
				className="shrink-0"
				onClick={() => {
					setPrompt(true)
				}}
			>
				{t("previewSheetModifyPasswordEdit")}
			</Button>
			<DocxPasswordPrompt
				open={prompt}
				title={t("previewSheetModifyPasswordTitle")}
				body={t("previewSheetModifyPasswordBody", { name })}
				wrongMessage={t("previewSheetModifyPasswordWrong")}
				label={t("previewSheetModifyPasswordLabel")}
				submitLabel={t("previewSheetModifyPasswordSubmit")}
				progressLabel={t("previewSheetModifyPasswordChecking")}
				verify={verify}
				onOpenChange={setPrompt}
			/>
		</>
	)
}
