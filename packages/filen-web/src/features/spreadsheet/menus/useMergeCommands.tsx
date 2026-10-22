import { useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import type { SheetCommand, SheetCommandOf, SheetRange } from "@filen/office/core/sheet"
import type { SheetGridModel, SheetGridState } from "@filen/office/react/sheet"
import { ConfirmDialog } from "@/components/dialogs/confirmDialog"
import { refusalKey } from "@/features/spreadsheet/lib/sheetMessages.logic"
import type { MergeMode } from "@/features/spreadsheet/menus/structureCommands"

// Merging and unmerging the selection. A merge that would drop values outside the upper-left cells asks first (as
// Excel does) and runs again once agreed; the confirmation is `dialog`, rendered by the caller outside its menu.
// `ask` puts that question for a merge refused elsewhere (F4 repeating one).
export function useMergeCommands(grid: SheetGridState): {
	merge: (mode: MergeMode) => void
	unmerge: () => void
	ask: (command: SheetCommandOf<"merge.cells">) => void
	dialog: ReactNode
} {
	const { t } = useTranslation(["preview", "common"])
	const model: SheetGridModel = grid.model
	const [asked, setAsked] = useState<SheetCommand | null>(null)

	function run(command: SheetCommand): void {
		void model.intent()
		void model.command(command, false).then(
			result => {
				if (result.ok) {
					return
				}

				if (result.refusal.reason === "mergeLosesValues" && command.type === "merge.cells") {
					setAsked({ ...command, discard: true })

					return
				}

				toast.error(t(refusalKey(result.refusal, "xlsx")), { id: result.refusal.reason })
			},
			() => undefined
		)
	}

	function merge(mode: MergeMode): void {
		const ranges: readonly SheetRange[] = grid.selection.ranges

		run({ type: "merge.cells", sheet: grid.sheet, ranges, mode })
	}

	function unmerge(): void {
		run({ type: "merge.unmerge", sheet: grid.sheet, ranges: grid.selection.ranges })
	}

	const dialog = (
		<ConfirmDialog
			open={asked !== null}
			pending={false}
			title={t("previewSheetMergeLosesTitle")}
			body={t("previewSheetMergeLosesBody")}
			confirmLabel={t("previewSheetMergeCells")}
			cancelLabel={t("common:cancel")}
			onOpenChange={open => {
				if (!open) {
					setAsked(null)
				}
			}}
			onConfirm={() => {
				if (asked !== null) {
					run(asked)
				}

				setAsked(null)
			}}
		/>
	)

	function ask(command: SheetCommandOf<"merge.cells">): void {
		setAsked({ ...command, discard: true })
	}

	return { merge, unmerge, ask, dialog }
}
