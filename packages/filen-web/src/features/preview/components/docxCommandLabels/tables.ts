import type { DocxCommandLabels } from "@/features/preview/components/docxEditor.logic"

// The Table group (docxTableMenu.tsx), the context menu and the column border drag; a new table is the Insert
// menu's.
export const TABLES_LABELS: Required<DocxCommandLabels<"tables">> = {
	"table.insert": "previewDocxInsertTable",
	"table.delete": "previewDocxTableDelete",
	"table.row.insertAbove": "previewDocxTableRowAbove",
	"table.row.insertBelow": "previewDocxTableRowBelow",
	"table.row.delete": "previewDocxTableRowDelete",
	"table.column.insertLeft": "previewDocxTableColumnLeft",
	"table.column.insertRight": "previewDocxTableColumnRight",
	"table.column.delete": "previewDocxTableColumnDelete",
	"table.cells.merge": "previewDocxTableMerge",
	"table.cell.split": "previewDocxTableSplit",
	"table.row.header": "previewDocxTableHeaderRow",
	"table.style": "previewDocxTableStyle",
	"table.align": "previewDocxTableAlign",
	"table.column.width": "previewDocxTableColumnWidth",
	"table.cell.shading": "previewDocxTableShading",
	"table.cell.vAlign": "previewDocxTableCellAlign",
	"table.distributeColumns": "previewDocxTableDistribute"
}
