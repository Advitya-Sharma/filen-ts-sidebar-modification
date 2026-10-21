// "preview" namespace, spreadsheet viewer on the office engine — the Insert and Delete dialogs' cell options
// (Ctrl+Shift+Plus, Ctrl+Minus), as Excel's. Spread into preview.ts; same catalog rules.
export const previewSheetShift = {
	/** Insert dialog — inserts cells at the selection, moving the cells there and to their right further right */
	previewSheetShiftCellsRight: "Shift cells right",
	/** Insert dialog — inserts cells at the selection, moving the cells there and below them further down */
	previewSheetShiftCellsDown: "Shift cells down",
	/** Delete dialog — deletes the selected cells, moving the cells to their right left into their place */
	previewSheetShiftCellsLeft: "Shift cells left",
	/** Delete dialog — deletes the selected cells, moving the cells below them up into their place */
	previewSheetShiftCellsUp: "Shift cells up"
}
