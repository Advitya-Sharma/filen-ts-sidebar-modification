import type { SheetAxis } from "@filen/office/core/sheet"
import type { SheetGridState, SheetShortcut } from "@filen/office/react/sheet"
import { selectionSpans, unhideSpans } from "@/features/spreadsheet/next/menus/structure.logic"

const AXIS_SHORTCUTS: ReadonlyMap<SheetShortcut, { readonly axis: SheetAxis; readonly hidden: boolean }> = new Map([
	["hideRows", { axis: "rows", hidden: true }],
	["hideColumns", { axis: "cols", hidden: true }],
	["unhideRows", { axis: "rows", hidden: false }],
	["unhideColumns", { axis: "cols", hidden: false }]
] as const)

// Excel's hide and unhide shortcuts (Ctrl+9, Ctrl+0, Ctrl+Shift+9, Ctrl+Shift+0) on the selection, as the grid's menu
// does them: hide the selected rows (columns); show the hidden ones in or next to it. True for those shortcuts (taken
// even where nothing can change), false for any other.
export function structureShortcut(grid: SheetGridState, shortcut: SheetShortcut, canEdit: boolean): boolean {
	const action = AXIS_SHORTCUTS.get(shortcut)

	if (action === undefined) {
		return false
	}

	if (!canEdit) {
		return true
	}

	const selected = selectionSpans(grid.selection.ranges, action.axis)
	const spans = action.hidden ? selected : unhideSpans(selected, grid.layout, action.axis)

	if (spans === null) {
		return true
	}

	void grid.model.intent()
	void grid.model.command({ type: "axis.hide", sheet: grid.sheet, axis: action.axis, spans, hidden: action.hidden })

	return true
}
