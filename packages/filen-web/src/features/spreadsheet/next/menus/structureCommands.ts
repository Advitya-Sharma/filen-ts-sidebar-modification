import type { SheetGridState } from "@filen/office/react/sheet"
import { canFreezePanes, freezePane, type FreezeTarget } from "@/features/spreadsheet/next/menus/structure.logic"

// What the structure menus run: the merge and freeze items (the toolbar's menus and the grid's share them).

export type MergeMode = "center" | "across" | "merge"

export const MERGE_ITEMS = [
	{ mode: "center", key: "previewSheetMergeCenter" },
	{ mode: "across", key: "previewSheetMergeAcross" },
	{ mode: "merge", key: "previewSheetMergeCells" }
] as const

// Whether the selection holds a merge to undo.
export function selectionHasMerge(grid: SheetGridState): boolean {
	const layout = grid.layout

	if (layout === null) {
		return false
	}

	return grid.model
		.merges(layout)
		.some(merge =>
			grid.selection.ranges.some(
				range =>
					range.startRow <= merge.endRow &&
					range.endRow >= merge.startRow &&
					range.startCol <= merge.endCol &&
					range.endCol >= merge.startCol
			)
		)
}

export const FREEZE_ITEMS = [
	{ target: "panes", key: "previewSheetFreezePanes" },
	{ target: "topRow", key: "previewSheetFreezeTopRow" },
	{ target: "firstColumn", key: "previewSheetFreezeFirstColumn" },
	{ target: "none", key: "previewSheetUnfreeze" }
] as const

// Freezes the active sheet's panes. A workbook writes them into the sheet's view (an edit, undone as one); a text file
// keeps them beside the file (useLocalPane, held by the grid's menu).
export function freezeTo(grid: SheetGridState, target: FreezeTarget): void {
	const { rows, cols } = freezePane(target, grid.selection.active)

	void grid.model.intent()
	void grid.model.command({ type: "view.freeze", sheet: grid.sheet, rows, cols })
}

export function freezeEnabled(grid: SheetGridState, target: FreezeTarget): boolean {
	if (target === "none") {
		return (grid.layout?.pane ?? null) !== null
	}

	return target !== "panes" || canFreezePanes(grid.selection.active)
}
