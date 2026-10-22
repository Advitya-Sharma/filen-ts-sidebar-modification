import { useState } from "react"
import { useTranslation } from "react-i18next"
import { DEFAULT_SHEET_ALLOWS } from "@filen/office/core/xlsx/edit"
import type { SheetProtectedAction } from "@filen/office/core/sheet"
import { FormDialog } from "@/components/dialogs/formDialog"
import { Field, FieldError, FieldGroup, FieldLabel } from "@/components/ui/field"
import { SecretInput } from "@/components/ui/secretInput"
import { PROTECT_ACTIONS } from "@/features/spreadsheet/tabs/sheetTabs.logic"

export interface ProtectRequest {
	// The optional password ("" for none).
	readonly password: string
	// A sheet's allowed actions (the workbook has none to choose).
	readonly allows: readonly SheetProtectedAction[]
}

// Protects a sheet (with the actions everyone may still do) or the workbook's structure, with an optional password
// typed twice. The caller hashes it and runs the command; `pending` while it does.
export function ProtectDialog({
	open,
	pending,
	workbook,
	onOpenChange,
	onProtect
}: {
	open: boolean
	pending: boolean
	workbook: boolean
	onOpenChange: (open: boolean) => void
	onProtect: (request: ProtectRequest) => void
}) {
	const { t } = useTranslation(["preview", "common"])
	const [password, setPassword] = useState("")
	const [confirm, setConfirm] = useState("")
	const [allows, setAllows] = useState<ReadonlySet<SheetProtectedAction>>(() => new Set(DEFAULT_SHEET_ALLOWS))
	const mismatch = password !== confirm

	function reset(): void {
		setPassword("")
		setConfirm("")
		setAllows(new Set(DEFAULT_SHEET_ALLOWS))
	}

	return (
		<FormDialog
			open={open}
			pending={pending}
			title={t(workbook ? "previewSheetProtectWorkbookTitle" : "previewSheetProtectTitle")}
			description={t(workbook ? "previewSheetProtectWorkbookBody" : "previewSheetProtectBody")}
			submitLabel={t("previewSheetProtectAction")}
			cancelLabel={t("common:cancel")}
			canSubmit={!mismatch}
			onOpenChange={next => {
				if (!next) {
					reset()
				}

				onOpenChange(next)
			}}
			onSubmit={event => {
				event.preventDefault()

				if (mismatch || pending) {
					return
				}

				onProtect({
					password,
					allows: workbook ? [] : PROTECT_ACTIONS.map(item => item.action).filter(action => allows.has(action))
				})
				reset()
			}}
		>
			<FieldGroup>
				<Field>
					<FieldLabel htmlFor="sheet-protect-password">{t("previewSheetProtectPassword")}</FieldLabel>
					<SecretInput
						id="sheet-protect-password"
						value={password}
						disabled={pending}
						onChange={event => {
							setPassword(event.target.value)

							// The confirmation is locked without a password: it goes with it.
							if (event.target.value === "") {
								setConfirm("")
							}
						}}
					/>
				</Field>
				<Field data-invalid={mismatch && confirm !== "" ? true : undefined}>
					<FieldLabel htmlFor="sheet-protect-confirm">{t("previewSheetProtectConfirm")}</FieldLabel>
					<SecretInput
						id="sheet-protect-confirm"
						value={confirm}
						disabled={pending || password === ""}
						aria-invalid={mismatch && confirm !== ""}
						onChange={event => {
							setConfirm(event.target.value)
						}}
					/>
					{mismatch && confirm !== "" ? <FieldError>{t("previewSheetProtectMismatch")}</FieldError> : null}
				</Field>
				{workbook ? null : (
					<fieldset className="flex flex-col gap-1.5">
						<legend className="mb-1.5 text-sm font-medium">{t("previewSheetProtectAllows")}</legend>
						<div className="grid grid-cols-1 gap-x-4 gap-y-1.5 sm:grid-cols-2">
							{PROTECT_ACTIONS.map(item => (
								<label
									key={item.action}
									className="flex items-center gap-2 text-sm"
								>
									<input
										type="checkbox"
										className="size-4 accent-primary"
										checked={allows.has(item.action)}
										disabled={pending}
										onChange={event => {
											const next = new Set(allows)

											if (event.target.checked) {
												next.add(item.action)
											} else {
												next.delete(item.action)
											}

											setAllows(next)
										}}
									/>
									{t(item.key)}
								</label>
							))}
						</div>
					</fieldset>
				)}
			</FieldGroup>
		</FormDialog>
	)
}
