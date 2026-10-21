import { validationEntry, type SheetGridState } from "@filen/office/react/sheet"

// Where a list validation's values open from: a cell, by sheet, row and column.
export function listCellKey(sheet: number, row: number, col: number): string {
	return `${String(sheet)}:${String(row)}:${String(col)}`
}

// Whether the active cell offers a list validation's values (Alt+Down opens them).
export function listShown(grid: SheetGridState, canEdit: boolean): boolean {
	const { row, col } = grid.selection.active

	return validationEntry(grid.layout, row, col)?.dropdown === true && canEdit && !grid.readOnly && grid.editing === null
}
