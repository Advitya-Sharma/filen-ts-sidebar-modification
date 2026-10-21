import { useId, useState } from "react"
import { useTranslation } from "react-i18next"
import type { SheetSummary } from "@filen/office/core/sheet"
import { FormDialog } from "@/components/dialogs/formDialog"
import { Field, FieldError, FieldLabel } from "@/components/ui/field"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group"
import { Select, SelectContent, SelectGroup, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { linkTarget, placeName, type LinkForm, type SheetLink } from "@/features/spreadsheet/next/comments/annotations.logic"

// Excel's Insert Hyperlink (Ctrl/Cmd+K): a web, mail or phone address, or a place in this workbook (a sheet and a cell
// reference, or a defined name); what the cell shows and a ScreenTip. Opened on a linked cell, it edits that link.

export interface LinkRequest {
	readonly target: { readonly url: string } | { readonly location: string }
	readonly text: string | undefined
	readonly tooltip: string | undefined
}

export interface LinkDialogProps {
	open: boolean
	// The active cell's link (null: a new one) and what it shows.
	link: SheetLink | null
	shown: string
	sheets: readonly SheetSummary[]
	sheet: number
	onClose: () => void
	onApply: (request: LinkRequest) => void
}

function initialForm({ link, shown, sheets, sheet }: LinkDialogProps): LinkForm {
	const place = link?.place ?? null
	const own = sheets.find(candidate => candidate.id === sheet)?.name ?? ""

	if (place !== null || (link?.location ?? null) !== null) {
		const name = place === null ? "" : placeName(sheets, place)
		const bang = name.lastIndexOf("!")

		return {
			kind: "place",
			address: "",
			sheet: place === null ? own : (sheets.find(candidate => candidate.id === place.sheet)?.name ?? own),
			reference: place === null ? (link?.location ?? "") : name.slice(bang + 1),
			text: shown,
			tooltip: link?.tooltip ?? ""
		}
	}

	return { kind: "web", address: link?.url ?? "", sheet: own, reference: "A1", text: shown, tooltip: link?.tooltip ?? "" }
}

// Mounted while closed too, so the dialog opens and closes as one (giving focus back where it came from, the grid).
// Each opening starts from the cell, and keeps what it opened with while it closes.
export function LinkDialog(props: LinkDialogProps) {
	const { open, sheets, onClose, onApply } = props
	const { t } = useTranslation(["preview", "common"])
	const id = useId()
	const [opened, setOpened] = useState({ open, link: props.link, shown: props.shown })
	const [form, setForm] = useState<LinkForm>(() => initialForm(props))
	const [checked, setChecked] = useState(false)

	if (opened.open !== open) {
		setOpened(open ? { open, link: props.link, shown: props.shown } : { ...opened, open })

		if (open) {
			setForm(initialForm(props))
			setChecked(false)
		}
	}

	const { link, shown } = opened
	const result = linkTarget(form)
	const change = (next: Partial<LinkForm>): void => {
		setForm(current => ({ ...current, ...next }))
	}
	const sheetItems = sheets.filter(sheet => sheet.kind === "worksheet").map(sheet => ({ value: sheet.name, label: sheet.name }))
	const title = t(link === null ? "previewSheetLinkInsertTitle" : "previewSheetLinkEditTitle")

	return (
		<FormDialog
			open={open}
			pending={false}
			title={title}
			description={t("previewSheetLinkBody")}
			submitLabel={t("previewSheetLinkApply")}
			cancelLabel={t("common:cancel")}
			canSubmit
			onOpenChange={open => {
				if (!open) {
					onClose()
				}
			}}
			onSubmit={event => {
				event.preventDefault()
				setChecked(true)

				if (!result.ok) {
					return
				}

				onApply({
					target: result.target,
					// The cell's text is written only when it was changed here.
					text: form.text === shown ? undefined : form.text,
					tooltip: form.tooltip.trim() === "" ? undefined : form.tooltip
				})
				onClose()
			}}
		>
			<div className="flex flex-col gap-4">
				<RadioGroup
					aria-label={t("previewSheetLinkKind")}
					name={`${id}-kind`}
					value={form.kind}
					onValueChange={next => {
						change({ kind: next === "place" ? "place" : "web" })
					}}
				>
					{(["web", "place"] as const).map(kind => (
						<div
							key={kind}
							className="flex items-center gap-3"
						>
							<RadioGroupItem
								id={`${id}-${kind}`}
								value={kind}
							/>
							<Label htmlFor={`${id}-${kind}`}>
								{t(kind === "web" ? "previewSheetLinkKindWeb" : "previewSheetLinkKindPlace")}
							</Label>
						</div>
					))}
				</RadioGroup>
				{form.kind === "web" ? (
					<Field data-invalid={checked && !result.ok}>
						<FieldLabel htmlFor={`${id}-address`}>{t("previewSheetLinkAddress")}</FieldLabel>
						<Input
							id={`${id}-address`}
							value={form.address}
							inputMode="url"
							autoComplete="off"
							spellCheck={false}
							autoFocus
							aria-invalid={checked && !result.ok}
							placeholder="https://"
							onChange={event => {
								change({ address: event.target.value })
							}}
						/>
						{checked && !result.ok ? <FieldError>{t("previewSheetLinkAddressInvalid")}</FieldError> : null}
					</Field>
				) : (
					<div className="grid gap-4 sm:grid-cols-2">
						<Field>
							<FieldLabel htmlFor={`${id}-sheet`}>{t("previewSheetLinkSheet")}</FieldLabel>
							<Select
								items={sheetItems}
								value={form.sheet}
								onValueChange={(next: string | null) => {
									if (next !== null) {
										change({ sheet: next })
									}
								}}
							>
								<SelectTrigger id={`${id}-sheet`}>
									<SelectValue />
								</SelectTrigger>
								<SelectContent data-preview-surface>
									<SelectGroup>
										{sheetItems.map(item => (
											<SelectItem
												key={item.value}
												value={item.value}
											>
												{item.label}
											</SelectItem>
										))}
									</SelectGroup>
								</SelectContent>
							</Select>
						</Field>
						<Field data-invalid={checked && !result.ok}>
							<FieldLabel htmlFor={`${id}-reference`}>{t("previewSheetLinkReference")}</FieldLabel>
							<Input
								id={`${id}-reference`}
								value={form.reference}
								autoComplete="off"
								spellCheck={false}
								aria-invalid={checked && !result.ok}
								onChange={event => {
									change({ reference: event.target.value })
								}}
							/>
							{checked && !result.ok ? <FieldError>{t("previewSheetLinkReferenceInvalid")}</FieldError> : null}
						</Field>
					</div>
				)}
				<Field>
					<FieldLabel htmlFor={`${id}-text`}>{t("previewSheetLinkText")}</FieldLabel>
					<Input
						id={`${id}-text`}
						value={form.text}
						autoComplete="off"
						onChange={event => {
							change({ text: event.target.value })
						}}
					/>
				</Field>
				<Field>
					<FieldLabel htmlFor={`${id}-tooltip`}>{t("previewSheetLinkTooltip")}</FieldLabel>
					<Input
						id={`${id}-tooltip`}
						value={form.tooltip}
						maxLength={255}
						autoComplete="off"
						onChange={event => {
							change({ tooltip: event.target.value })
						}}
					/>
				</Field>
			</div>
		</FormDialog>
	)
}
