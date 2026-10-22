import { useState } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import type { SheetCommand, SheetFileKind } from "@filen/office/core/sheet"
import type { SheetGridModel, SheetGridState } from "@filen/office/react/sheet"
import type { SheetBodyExtension } from "@/features/spreadsheet/components/sheetExtension"
import type { PrintSheetSettings } from "@/features/spreadsheet/print/print.logic"
import { SheetPrintDialog } from "@/features/spreadsheet/print/printDialog"
import { sheetSettings } from "@/features/spreadsheet/print/printSettings"
import { log } from "@/lib/log"

export interface SheetPrint {
	readonly extension: SheetBodyExtension
	// Opens the print dialog (the file menu's Print…); undefined while the session cannot print.
	readonly open: (() => void) | undefined
}

// Printing: Mod+P and the file menu open the print dialog over the sheet on show, its settings read first (an open
// entry committed before), so the dialog opens with the pages as they print.
export function useSheetPrint({
	grid,
	model,
	kind,
	name,
	canEdit,
	run
}: {
	grid: SheetGridState
	model: SheetGridModel
	kind: SheetFileKind
	name: string
	canEdit: boolean
	run: (command: SheetCommand) => void
}): SheetPrint {
	const { t } = useTranslation("preview")
	const [opened, setOpened] = useState<{ readonly active: PrintSheetSettings; readonly count: number } | null>(null)
	const [count, setCount] = useState(0)
	const sheet = grid.sheets.find(entry => entry.id === grid.sheet)
	const open =
		grid.failed !== null || sheet === undefined
			? undefined
			: () => {
					void model
						.commitEditing(null)
						.then(done => (done ? sheetSettings(model, kind, sheet) : undefined))
						.then(
							active => {
								if (active === undefined) {
									return
								}

								if (active === null) {
									toast.error(t("previewSheetPrintNothing"))

									return
								}

								setCount(count + 1)
								setOpened({ active, count: count + 1 })
							},
							(error: unknown) => {
								log.warn("preview", "spreadsheet print settings failed", error)
								toast.error(t("previewSheetPrintFailed"))
							}
						)
				}

	return {
		open,
		extension: {
			id: "print",
			shortcut: shortcut => {
				if (shortcut !== "print") {
					return false
				}

				open?.()

				return true
			},
			dialogs:
				opened === null ? null : (
					<SheetPrintDialog
						key={opened.count}
						active={opened.active}
						grid={grid}
						model={model}
						kind={kind}
						name={name}
						canSave={canEdit}
						run={run}
						onClose={() => {
							setOpened(null)
						}}
					/>
				)
		}
	}
}
