// "preview" namespace, spreadsheet viewer — saving: the file menu (Save as .xlsx, Download as
// CSV, autosave). Spread into preview.ts; same catalog rules.
export const previewSheetSave = {
	/** Spreadsheet viewer — accessible label and tooltip of the button opening the file menu (Save as .xlsx, Download as CSV, Save automatically) */
	previewSheetFileMenu: "File",
	/** Spreadsheet viewer — file menu item for a CSV or TSV file: downloads its cells as a CSV file (UTF-8, comma-separated) */
	previewSheetDownloadCsv: "Download as CSV",
	/** Spreadsheet viewer — file menu item for a workbook: downloads the sheet on show as a CSV file (UTF-8, comma-separated), each cell as it is shown */
	previewSheetDownloadSheetCsv: "Download this sheet as CSV",
	/** Toast — the sheet could not be turned into a CSV file */
	previewSheetDownloadCsvFailed: "This sheet couldn't be downloaded as CSV.",
	/** Spreadsheet viewer — file menu checkbox: save edits automatically about 30 seconds after the last one (the same setting as Settings → Advanced → Documents) */
	previewSheetAutosave: "Save automatically",
	/** Spreadsheet viewer — confirm dialog body before a CSV or TSV file is saved as an .xlsx copy */
	previewSheetSaveAsXlsxTextBody:
		"An .xlsx copy with your changes is saved in the same directory and opens here; this file stays as it was last saved. Numbers become numbers, and everything else stays text, formulas and dates included.",
	/** Toast — the file holds more than an .xlsx workbook can: rows past 1,048,576, columns past 16,384, or a cell longer than 32,767 characters */
	previewSheetSaveAsXlsxTooLarge: "This file holds more than an .xlsx workbook can, so it can't be saved as one."
}
