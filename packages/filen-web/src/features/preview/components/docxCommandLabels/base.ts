import type { DocxCommandLabels } from "@/features/preview/components/docxEditor.logic"

// Every base command is labelled: the toolbar row shows them all.
export const BASE_LABELS: Required<DocxCommandLabels<"base">> = {
	"format.bold": "previewDocxBold",
	"format.italic": "previewDocxItalic",
	"format.underline": "previewDocxUnderline",
	"format.strike": "previewDocxStrikethrough",
	"format.clear": "previewDocxClearFormatting",
	"style.set": "previewDocxParagraphStyle",
	"align.start": "previewDocxAlignStart",
	"align.center": "previewDocxAlignCenter",
	"align.end": "previewDocxAlignEnd",
	"align.justify": "previewDocxAlignJustify",
	"list.bullet": "previewDocxBulletList",
	"list.numbered": "previewDocxNumberedList",
	"list.indent": "previewDocxIndent",
	"list.outdent": "previewDocxOutdent",
	"history.undo": "previewDocxUndo",
	"history.redo": "previewDocxRedo"
}
