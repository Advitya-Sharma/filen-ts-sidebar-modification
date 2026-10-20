import { useEffect, useState } from "react"
import type { CellStyleView, FormatPatch } from "@filen/office/core/sheet"
import type { SheetGridModel, SheetGridState } from "@filen/office/react/sheet"
import { FormatToolbar } from "@/features/spreadsheet/components/formatToolbar"
import { sheetFormatPatch, toolbarStyle } from "@/features/spreadsheet/next/lib/toolbarStyle.logic"

// The active cell's look and number format, asked of the session when the active cell or what it shows changes.
// Only while the formats are on show: a text file has none.
function useActiveStyle(grid: SheetGridState, wanted: boolean): CellStyleView | undefined {
	const { model, sheet, selection, state } = grid
	const { row, col } = selection.active
	const [style, setStyle] = useState<{ key: string; style: CellStyleView } | null>(null)
	const key = `${String(sheet)}:${String(row)}:${String(col)}:${String(state.version)}`

	useEffect(() => {
		if (!wanted) {
			return undefined
		}

		let live = true

		model.session.query("cell", { sheet, row, col }).then(
			detail => {
				if (live) {
					setStyle({ key, style: toolbarStyle(model.style(detail.style), detail.numFmt) })
				}
			},
			() => undefined
		)

		return () => {
			live = false
		}
	}, [model, sheet, row, col, key, wanted])

	// The last cell's look until this one's arrives, so the buttons do not flicker while the arrows move on.
	return wanted ? style?.style : undefined
}

// Undo and redo for every editable sheet; the formats too for a workbook (a text file holds none). Shown disabled
// while editing waits to become possible, so the grid below does not move when it does.
export function SheetToolbar({
	grid,
	disabled,
	formats,
	onFormat
}: {
	grid: SheetGridState
	disabled: boolean
	formats: boolean
	onFormat: (patch: ReturnType<typeof sheetFormatPatch>) => void
}) {
	const style = useActiveStyle(grid, formats && !disabled)
	const model: SheetGridModel = grid.model

	return (
		<FormatToolbar
			disabled={disabled}
			style={style}
			formats={formats}
			canUndo={grid.state.canUndo}
			canRedo={grid.state.canRedo}
			onUndo={() => {
				void model.undo()
			}}
			onRedo={() => {
				void model.redo()
			}}
			onFormat={(patch: FormatPatch) => {
				onFormat(sheetFormatPatch(patch))
			}}
		/>
	)
}
