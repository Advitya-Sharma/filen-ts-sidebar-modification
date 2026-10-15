import type { DocxCommandLabels, DocxCommandVariant } from "@/features/preview/components/docxEditor.logic"

// Content controls: filled in where they are (docxFormControls.tsx: a list's entries added and removed in its
// popover), cleared and their repeating items added or removed from the context menu, inserted from Insert ▸
// Form control. A legacy text field is typed into.
export const FORMS_LABELS: DocxCommandLabels<"forms"> = {
	"control.check": "previewDocxControlCheck",
	"control.choose": "previewDocxControlChoose",
	"control.date": "previewDocxControlDate",
	"control.clear": "previewDocxControlClear",
	"control.insert": "previewDocxControlInsert",
	"control.items": "previewDocxControlItemsEdit",
	"control.item.add": "previewDocxControlItemAdd",
	"control.item.remove": "previewDocxControlItemRemove"
}

// Insert ▸ Form control ▸ each kind; a dropdown list asks for its items first (docxFormControls.tsx).
export const CONTROL_INSERT_VARIANTS: readonly DocxCommandVariant[] = [
	{ label: "previewDocxControlInsertCheckbox", command: { type: "control.insert", kind: "checkbox" } },
	{ label: "previewDocxControlInsertDropdown", input: "dropDownList" },
	{ label: "previewDocxControlInsertDate", command: { type: "control.insert", kind: "date" } },
	{ label: "previewDocxControlInsertText", command: { type: "control.insert", kind: "plainText" } }
]
