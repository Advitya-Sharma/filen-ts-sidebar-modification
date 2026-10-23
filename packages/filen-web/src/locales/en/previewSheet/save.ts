// "preview" namespace, spreadsheet viewer — saving: the file menu (Save as .xlsx, Download as ▸
// PDF or CSV, autosave). Spread into preview.ts; same catalog rules.
export const previewSheetSave = {
	/** Spreadsheet viewer — accessible label and tooltip of the button opening the file menu (Save as .xlsx, Download as CSV, Save automatically) */
	previewSheetFileMenu: "File",
	/** Toast — the sheet could not be turned into a CSV file */
	previewSheetDownloadCsvFailed: "This sheet couldn't be downloaded as CSV.",
	/** Spreadsheet viewer — file menu submenu of the formats the file can be downloaded as (a local download; the file in the drive is unchanged) */
	previewSheetDownloadAs: "Download as",
	/** File menu → Download as — item downloading the sheet on show (its print area, as it prints) as a PDF file */
	previewSheetDownloadPdf: "PDF document (.pdf)",
	/** File menu → Download as — item for a CSV or TSV file: its cells as a CSV file (UTF-8, comma-separated) */
	previewSheetDownloadCsvFormat: "CSV (.csv)",
	/** File menu → Download as — item for a workbook: the sheet on show as a CSV file, each cell as it is shown */
	previewSheetDownloadSheetCsvFormat: "This sheet as CSV (.csv)",
	/** Toast while the PDF is being made, before the page count is known */
	previewSheetPdfPreparing: "Preparing PDF…",
	/** Toast while the PDF is being made: pages written so far of all; {{done}} and {{total}} are numbers */
	previewSheetPdfProgress: "Making PDF… page {{done}} of {{total}}",
	/** Toast action stopping the PDF export */
	previewSheetPdfCancel: "Cancel",
	/** Toast — the PDF could not be made */
	previewSheetPdfFailed: "The PDF could not be made",
	/** Toast — nothing on the sheet (or the chosen part) prints, so there is no PDF to make */
	previewSheetPdfNothing: "There's nothing to put in a PDF.",
	/** Toast — the print has more pages than a PDF is made of; {{pages}} is the most */
	previewSheetPdfTooMany: "This would be more than {{pages}} pages. Set a print area or scale to make fewer.",
	/** Toast — one page of the print holds more cells than a PDF page can be made with; {{page}} is that page's number */
	previewSheetPdfPageTooLarge: "Page {{page}} holds too many cells for a PDF. Set a print area, a larger scale or a smaller paper size.",
	/** Toast after the PDF was downloaded, when some text was too complex to shape and is drawn with its letters as they come */
	previewSheetPdfUnshaped: "Some text in the PDF is drawn without its font's letter shaping",
	/** Toast — the workbook's password uses characters a PDF password cannot hold, so no protected PDF can be made */
	previewSheetPdfPasswordUnsupported: "The PDF could not be made: this workbook's password cannot protect a PDF",
	/** Toast after the PDF was downloaded, when some characters had no font or some pictures could not be put in */
	previewSheetPdfSimplified: "Some characters or pictures could not be put in the PDF",
	/** Spreadsheet viewer — file menu checkbox: save edits automatically about 30 seconds after the last one (the same setting as Settings → Advanced → Documents) */
	previewSheetAutosave: "Save automatically",
	/** Spreadsheet viewer — confirm dialog body before a CSV or TSV file is saved as an .xlsx copy */
	previewSheetSaveAsXlsxTextBody:
		"An .xlsx copy with your changes is saved in the same directory and opens here; this file stays as it was last saved. Numbers become numbers, and everything else stays text, formulas and dates included.",
	/** Toast — the file holds more than an .xlsx workbook can: rows past 1,048,576, columns past 16,384, or a cell longer than 32,767 characters */
	previewSheetSaveAsXlsxTooLarge: "This file holds more than an .xlsx workbook can, so it can't be saved as one."
}
