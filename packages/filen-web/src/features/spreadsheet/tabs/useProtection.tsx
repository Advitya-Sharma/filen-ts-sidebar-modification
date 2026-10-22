import { useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import type { SheetCommand, SheetId } from "@filen/office/core/sheet"
import { sheetPasswordCheck, type SheetPasswordHash } from "@filen/office/core/xlsx/edit"
import { hashProtectionPassword, verifyProtectionPassword } from "@filen/office/dom"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { log } from "@/lib/log"
import { DocxPasswordPrompt, type PasswordVerdict } from "@/features/preview/components/docxPasswordPrompt"
import { ProtectDialog, type ProtectRequest } from "@/features/spreadsheet/tabs/protectDialog"

export type ProtectionTarget = { readonly sheet: SheetId } | "workbook"

const ALGORITHMS = new Set(["SHA-1", "SHA-256", "SHA-384", "SHA-512"] as const)

type Algorithm = "SHA-1" | "SHA-256" | "SHA-384" | "SHA-512"

function isAlgorithm(name: string): name is Algorithm {
	return (ALGORITHMS as ReadonlySet<string>).has(name)
}

// Whether `password` opens a protection's hash: the legacy verifier at once, the modern hash on the platform (which
// `signal` stops: it then rejects). Null for a hash this app cannot check.
export async function passwordOpens(hash: SheetPasswordHash | null, password: string, signal?: AbortSignal): Promise<boolean | null> {
	const check = sheetPasswordCheck(hash, password)

	switch (check.kind) {
		case "none":
			return true
		case "legacy":
			return check.correct
		case "modern": {
			if (!isAlgorithm(check.algorithm)) {
				return null
			}

			const verdict = await verifyProtectionPassword(
				{ algorithm: check.algorithm, hash: check.hash, salt: check.salt, spinCount: check.spinCount, input: "plain" },
				password,
				signal === undefined ? {} : { signal }
			)

			return verdict === "unsupported" ? null : verdict === "correct"
		}
	}
}

// Protecting and unprotecting a sheet or the workbook's structure: the dialogs (`dialogs`, rendered by the caller)
// and what starts them. A password is hashed here before the command is sent, and checked here before protection
// comes off (the session hashes nothing); closing the unprotect prompt stops a check that is still running.
export function useProtection(model: SheetGridModel): {
	protect: (target: ProtectionTarget) => void
	unprotect: (target: ProtectionTarget) => void
	dialogs: ReactNode
} {
	const { t } = useTranslation(["preview", "common"])
	const [protecting, setProtecting] = useState<ProtectionTarget | null>(null)
	const [unprotecting, setUnprotecting] = useState<{ target: ProtectionTarget; hash: SheetPasswordHash } | null>(null)
	const [pending, setPending] = useState(false)

	function send(command: SheetCommand): Promise<void> {
		void model.intent()

		return model.command(command).then(
			() => undefined,
			() => undefined
		)
	}

	function off(target: ProtectionTarget): Promise<void> {
		return send(target === "workbook" ? { type: "workbook.unprotect" } : { type: "sheet.unprotect", sheet: target.sheet })
	}

	function unprotect(target: ProtectionTarget): void {
		model.session.query("protection.hash", { target }).then(
			hash => {
				if (hash === null) {
					void off(target)
				} else {
					setUnprotecting({ target, hash })
				}
			},
			() => undefined
		)
	}

	async function protectWith(target: ProtectionTarget, request: ProtectRequest): Promise<void> {
		const hash = request.password === "" ? undefined : await hashProtectionPassword(request.password)

		await send(
			target === "workbook"
				? { type: "workbook.protect", structure: true, windows: false, ...(hash !== undefined && { hash }) }
				: { type: "sheet.protect", sheet: target.sheet, allows: request.allows, ...(hash !== undefined && { hash }) }
		)
	}

	async function unprotectWith(
		target: ProtectionTarget,
		hash: SheetPasswordHash,
		password: string,
		signal: AbortSignal
	): Promise<PasswordVerdict> {
		let opens: boolean | null

		try {
			opens = await passwordOpens(hash, password, signal)
		} catch (e) {
			if (!signal.aborted) {
				log.error("spreadsheet", "unprotecting failed", e)
			}

			return "aborted"
		}

		if (signal.aborted) {
			return "aborted"
		}

		if (opens === null) {
			toast.error(t("previewSheetPasswordUnsupported"))
			setUnprotecting(null)

			return "aborted"
		}

		if (!opens) {
			return "wrong"
		}

		await off(target)
		setUnprotecting(null)

		return "accepted"
	}

	const dialogs = (
		<>
			<ProtectDialog
				open={protecting !== null}
				pending={pending}
				workbook={protecting === "workbook"}
				onOpenChange={open => {
					if (!open) {
						setProtecting(null)
					}
				}}
				onProtect={request => {
					if (protecting === null) {
						return
					}

					setPending(true)
					protectWith(protecting, request)
						.catch((e: unknown) => {
							log.error("spreadsheet", "protecting failed", e)
						})
						.finally(() => {
							setPending(false)
							setProtecting(null)
						})
				}}
			/>
			<DocxPasswordPrompt
				open={unprotecting !== null}
				title={t(unprotecting?.target === "workbook" ? "previewSheetUnprotectWorkbookTitle" : "previewSheetUnprotectTitle")}
				body={t("previewSheetUnprotectBody")}
				wrongMessage={t("previewSheetWrongPassword")}
				label={t("previewSheetUnprotectPassword")}
				submitLabel={t("previewSheetUnprotectAction")}
				progressLabel={t("previewSheetUnprotectChecking")}
				verify={(password, signal) =>
					unprotecting === null
						? Promise.resolve("aborted")
						: unprotectWith(unprotecting.target, unprotecting.hash, password, signal)
				}
				onOpenChange={open => {
					if (!open) {
						setUnprotecting(null)
					}
				}}
			/>
		</>
	)

	return { protect: setProtecting, unprotect, dialogs }
}
