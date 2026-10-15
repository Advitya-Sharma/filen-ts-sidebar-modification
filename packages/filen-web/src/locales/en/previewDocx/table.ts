// "preview" namespace, Word document editor — the editor's table commands. Spread into preview.ts; same catalog rules.
export const previewDocxTable = {
	/** Docx editor — toolbar's contextual Table group button (shown while the caret is in a table) */
	previewDocxTableGroup: "Table",
	/** Docx editor — table command: deletes the whole table */
	previewDocxTableDelete: "Delete table",
	/** Docx editor — table command */
	previewDocxTableRowAbove: "Insert row above",
	/** Docx editor — table command */
	previewDocxTableRowBelow: "Insert row below",
	/** Docx editor — table command */
	previewDocxTableRowDelete: "Delete rows",
	/** Docx editor — table command */
	previewDocxTableColumnLeft: "Insert column left",
	/** Docx editor — table command */
	previewDocxTableColumnRight: "Insert column right",
	/** Docx editor — table command */
	previewDocxTableColumnDelete: "Delete columns",
	/** Docx editor — table command */
	previewDocxTableMerge: "Merge cells",
	/** Docx editor — table command opening the split dialog */
	previewDocxTableSplit: "Split cells…",
	/** Docx editor — table toggle: the first selected row repeats as a header row */
	previewDocxTableHeaderRow: "Header row",
	/** Docx editor — Table group submenu: the document's table styles */
	previewDocxTableStyle: "Table style",
	/** Docx editor — Table group submenu: where the table sits on the page */
	previewDocxTableAlign: "Table alignment",
	/** Docx editor — table alignment option (left in left-to-right text) */
	previewDocxTableAlignStart: "Start",
	/** Docx editor — table alignment option */
	previewDocxTableAlignCenter: "Center",
	/** Docx editor — table alignment option (right in left-to-right text) */
	previewDocxTableAlignEnd: "End",
	/** Docx editor — label of dragging a column border */
	previewDocxTableColumnWidth: "Column width",
	/** Docx editor — Table group submenu: the selected cells' fill colour */
	previewDocxTableShading: "Shading",
	/** Docx editor — Table group submenu: vertical alignment of the selected cells' text */
	previewDocxTableCellAlign: "Cell alignment",
	/** Docx editor — cell vertical alignment option */
	previewDocxTableCellTop: "Top",
	/** Docx editor — cell vertical alignment option */
	previewDocxTableCellCenter: "Center",
	/** Docx editor — cell vertical alignment option */
	previewDocxTableCellBottom: "Bottom",
	/** Docx editor — table command: makes the selected columns equally wide */
	previewDocxTableDistribute: "Distribute columns",
	/** Docx editor — split cells dialog title */
	previewDocxTableSplitTitle: "Split cells",
	/** Docx editor — split cells dialog description */
	previewDocxTableSplitBody: "Split the selected cell into rows and columns.",
	/** Docx editor — rows × columns dialog field (a new table, a cell split) */
	previewDocxTableRows: "Rows",
	/** Docx editor — rows × columns dialog field (a new table, a cell split) */
	previewDocxTableColumns: "Columns",
	/** Docx editor — split cells dialog error: the table refused the split (the dialog stays open) */
	previewDocxTableSplitRefused: "This cell can't be split into that many rows and columns.",
	/** Docx editor — split cells dialog submit */
	previewDocxTableSplitAction: "Split",
	/** Docx editor — accessible label of the guide shown while dragging a table column border */
	previewDocxTableBorderGuide: "New column border"
} as const
