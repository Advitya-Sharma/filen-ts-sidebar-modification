import type { DocxCommandLabels } from "@/features/preview/components/docxEditor.logic"

// Footnotes and endnotes: inserted, converted and set up from the Insert menu; gone to and deleted from the
// context menu and the chip at a note reference (docxFieldChip.tsx).
export const NOTES_LABELS: Required<DocxCommandLabels<"notes">> = {
	"note.insertFootnote": "previewDocxInsertFootnote",
	"note.insertEndnote": "previewDocxInsertEndnote",
	"note.goto": "previewDocxNoteGoTo",
	"note.delete": "previewDocxNoteDelete",
	"note.convert": "previewDocxNoteConvert",
	"note.convertAll": "previewDocxNoteConvertAll",
	"note.options": "previewDocxNoteOptions"
}
