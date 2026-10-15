import { useState, type SubmitEvent } from "react"
import { useTranslation } from "react-i18next"
import { TABLE_LIMITS } from "@/features/preview/components/docxEditor.logic"
import { FormDialog } from "@/components/dialogs/formDialog"
import { Field, FieldError, FieldGroup, FieldLabel } from "@/components/ui/field"
import { Input } from "@/components/ui/input"
import { Select, SelectContent, SelectGroup, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"

// The editor's dialogs' shared fields: a labelled choice, and the rows × columns dialog (a new table, a cell
// split).

export function DocxChoiceField<T extends string>({
	label,
	value,
	items,
	onChange
}: {
	label: string
	value: T
	items: readonly { value: T; label: string }[]
	onChange: (value: T) => void
}) {
	return (
		<Field>
			<FieldLabel>{label}</FieldLabel>
			<Select
				items={items}
				value={value}
				onValueChange={next => {
					const chosen = items.find(item => item.value === next)

					if (chosen !== undefined) {
						onChange(chosen.value)
					}
				}}
			>
				<SelectTrigger aria-label={label}>
					<SelectValue />
				</SelectTrigger>
				<SelectContent>
					<SelectGroup>
						{items.map(item => (
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
	)
}

// A count from 1 to `most`, or undefined.
function countOf(text: string, most: number): number | undefined {
	const trimmed = text.trim()
	const value = Number(trimmed)

	return /^\d+$/.test(trimmed) && value >= 1 && value <= most ? value : undefined
}

export interface DocxRowsColumnsDialogProps {
	title: string
	description: string
	submitLabel: string
	initial: { readonly rows: number; readonly cols: number }
	// Whether the size is one to take (a split needs more than one cell).
	accepts?: (size: { rows: number; cols: number }) => boolean
	// True when done (the dialog closes); false keeps it open, saying `refusal`.
	onSubmit: (size: { rows: number; cols: number }) => boolean
	refusal?: string
	onClose: () => void
}

// Rows × columns within Word's table limits; mounted while asked for (its fields start from `initial`).
export function DocxRowsColumnsDialog({
	title,
	description,
	submitLabel,
	initial,
	accepts,
	onSubmit,
	refusal,
	onClose
}: DocxRowsColumnsDialogProps) {
	const { t } = useTranslation(["preview", "common"])
	const [rows, setRows] = useState(String(initial.rows))
	const [cols, setCols] = useState(String(initial.cols))
	// The last size was refused (shown until a field changes).
	const [refused, setRefused] = useState(false)
	const size = { rows: countOf(rows, TABLE_LIMITS.rows), cols: countOf(cols, TABLE_LIMITS.columns) }
	const valid =
		size.rows !== undefined && size.cols !== undefined && (accepts === undefined || accepts({ rows: size.rows, cols: size.cols }))

	return (
		<FormDialog
			open
			pending={false}
			title={title}
			description={description}
			submitLabel={submitLabel}
			cancelLabel={t("common:cancel")}
			canSubmit={valid}
			onOpenChange={open => {
				if (!open) {
					onClose()
				}
			}}
			onSubmit={(event: SubmitEvent) => {
				event.preventDefault()

				if (size.rows === undefined || size.cols === undefined || !valid) {
					return
				}

				if (onSubmit({ rows: size.rows, cols: size.cols })) {
					onClose()
				} else {
					setRefused(true)
				}
			}}
		>
			<FieldGroup className="grid grid-cols-2 gap-4">
				<Field>
					<FieldLabel htmlFor="docx-size-rows">{t("previewDocxTableRows")}</FieldLabel>
					<Input
						id="docx-size-rows"
						inputMode="numeric"
						value={rows}
						onChange={event => {
							setRows(event.target.value)
							setRefused(false)
						}}
					/>
				</Field>
				<Field>
					<FieldLabel htmlFor="docx-size-cols">{t("previewDocxTableColumns")}</FieldLabel>
					<Input
						id="docx-size-cols"
						inputMode="numeric"
						value={cols}
						onChange={event => {
							setCols(event.target.value)
							setRefused(false)
						}}
					/>
				</Field>
				{refused && refusal !== undefined ? (
					<FieldError
						role="alert"
						className="col-span-2"
					>
						{refusal}
					</FieldError>
				) : null}
			</FieldGroup>
		</FormDialog>
	)
}
