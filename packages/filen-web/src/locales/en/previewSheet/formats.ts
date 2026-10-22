// "preview" namespace, spreadsheet viewer — the formats a spreadsheet opens as beyond .xlsx,
// CSV and TSV. Spread into preview.ts; same catalog rules.
export const previewSheetFormats = {
	/** Spreadsheet viewer — note beside the cell contents for a legacy Excel template (.xlt), which opens view-only; the "Save as .xlsx" button beside it makes an editable copy */
	previewSheetReadOnlyXlt: "Old .xlt template format — view only.",
	/** Spreadsheet viewer — confirm dialog body before a legacy Excel template (.xlt) is saved as an .xlsx copy: what the conversion keeps and loses */
	previewSheetSaveAsXlsxXltBody:
		"An editable .xlsx copy is saved in the same directory, and the .xlt stays as it is. Only the values come along: formulas become their results, and formatting, column widths, row heights and named ranges are not kept."
}
