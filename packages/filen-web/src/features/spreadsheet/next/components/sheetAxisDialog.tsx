import { useId, useState } from "react"
import { useTranslation } from "react-i18next"
import type { SheetCommand } from "@filen/office/core/sheet"
import type { SheetGridState } from "@filen/office/react/sheet"
import { FormDialog } from "@/components/dialogs/formDialog"
import { Label } from "@/components/ui/label"
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group"

export type SheetAxisDialogKind = "insert" | "delete"

// Excel's Insert and Delete dialogs (Ctrl+Shift+Plus, Ctrl+Minus on cells): the selection's whole rows or columns
// inserted at it or deleted. Shifting part of a row or column is not offered: the engine moves whole ones.
export function SheetAxisDialog({
	kind,
	grid,
	run,
	onClose
}: {
	kind: SheetAxisDialogKind | null
	grid: SheetGridState
	run: (command: SheetCommand) => void
	onClose: () => void
}) {
	const { t } = useTranslation(["preview", "common"])
	const id = useId()
	const [axis, setAxis] = useState<"rows" | "cols">("rows")
	const range = grid.selection.ranges[grid.selection.ranges.length - 1]
	const title = t(kind === "delete" ? "previewSheetDeleteDialogTitle" : "previewSheetInsertDialogTitle")

	return (
		<FormDialog
			open={kind !== null}
			pending={false}
			title={title}
			description={t(kind === "delete" ? "previewSheetDeleteDialogBody" : "previewSheetInsertDialogBody")}
			submitLabel={title}
			cancelLabel={t("common:cancel")}
			canSubmit={range !== undefined}
			onOpenChange={open => {
				if (!open) {
					onClose()
				}
			}}
			onSubmit={event => {
				event.preventDefault()

				if (kind !== null && range !== undefined) {
					run({
						type: kind === "insert" ? "axis.insert" : "axis.delete",
						sheet: grid.sheet,
						axis,
						at: axis === "rows" ? range.startRow : range.startCol,
						count: axis === "rows" ? range.endRow - range.startRow + 1 : range.endCol - range.startCol + 1
					})
				}

				onClose()
			}}
		>
			<RadioGroup
				aria-label={title}
				name={`${id}-axis`}
				value={axis}
				onValueChange={next => {
					setAxis(next === "cols" ? "cols" : "rows")
				}}
			>
				{(["rows", "cols"] as const).map(value => (
					<div
						key={value}
						className="flex items-center gap-3"
					>
						<RadioGroupItem
							id={`${id}-${value}`}
							value={value}
						/>
						<Label htmlFor={`${id}-${value}`}>
							{t(value === "rows" ? "previewSheetDialogEntireRow" : "previewSheetDialogEntireColumn")}
						</Label>
					</div>
				))}
			</RadioGroup>
		</FormDialog>
	)
}
