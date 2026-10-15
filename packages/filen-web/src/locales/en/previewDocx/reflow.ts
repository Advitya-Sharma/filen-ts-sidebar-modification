// "preview" namespace, Word document editor — the editor's reflowed (mobile) view. Spread into preview.ts; same catalog rules.
export const previewDocxReflow = {
	/** Docx editor — ⋯ menu switch: the document reflowed to the width of the screen for reading and editing on a phone, instead of its pages (Word's "Mobile view") */
	previewDocxMobileView: "Mobile view",
	/** Docx editor, Mobile view — accessible name of a footnote's text shown right after its reference mark in the text */
	previewDocxMobileFootnote: "Footnote",
	/** Docx editor, Mobile view — accessible name of an endnote's text shown right after its reference mark in the text */
	previewDocxMobileEndnote: "Endnote",
	/** Docx editor, Mobile view — accessible name of a footnote's text shown after its reference mark; {{mark}} is the note's number as shown (3, iv, c) */
	previewDocxMobileFootnoteNumbered: "Footnote {{mark}}",
	/** Docx editor, Mobile view — accessible name of an endnote's text shown after its reference mark; {{mark}} is the note's number as shown (3, iv, c) */
	previewDocxMobileEndnoteNumbered: "Endnote {{mark}}",
	/** Docx editor, Mobile view — accessible name of a text box's text shown where the text box is anchored in the text */
	previewDocxMobileTextBox: "Text box"
} as const
