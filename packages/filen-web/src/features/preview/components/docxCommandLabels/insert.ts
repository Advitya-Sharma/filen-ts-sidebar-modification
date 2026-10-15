import type { DocxCommandLabels } from "@/features/preview/components/docxEditor.logic"

// Links, breaks, symbols and the horizontal line: the Insert menu (docxInsertMenu.tsx) and its dialogs.
export const INSERT_LABELS: Required<DocxCommandLabels<"insert">> = {
	"link.set": "previewDocxInsertLink",
	"link.remove": "previewDocxRemoveLink",
	"insert.pageBreak": "previewDocxInsertPageBreak",
	"insert.lineBreak": "previewDocxInsertLineBreak",
	"insert.horizontalLine": "previewDocxInsertHorizontalLine",
	"insert.symbol": "previewDocxInsertSymbol"
}
