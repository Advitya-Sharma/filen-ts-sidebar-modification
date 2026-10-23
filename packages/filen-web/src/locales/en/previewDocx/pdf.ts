// "preview" namespace, Word document viewer and editor — the File menu and Download as PDF. Spread into preview.ts;
// same catalog rules.
export const previewDocxPdf = {
	/** Docx viewer and editor — accessible label and tooltip of the toolbar's File menu button */
	previewDocxFileMenu: "File",
	/** Docx File menu — submenu of the formats the document can be downloaded as (a local download; the file in the drive is unchanged) */
	previewDocxDownloadAs: "Download as",
	/** Docx File menu → Download as — item downloading the document as a PDF file */
	previewDocxDownloadPdf: "PDF document (.pdf)",
	/** Toast while the PDF is being made, before the page count is known */
	previewDocxPdfPreparing: "Preparing PDF…",
	/** Toast while the PDF is being made: pages written so far of all */
	previewDocxPdfProgress: "Making PDF… page {{done}} of {{total}}",
	/** Toast action stopping the PDF export */
	previewDocxPdfCancel: "Cancel",
	/** Toast after the PDF was downloaded, when some characters had no font to draw them (they show as boxes) */
	previewDocxPdfMissingGlyphs: "Some characters could not be drawn in the PDF",
	/** Toast after the PDF was downloaded, when pictures the PDF cannot carry (unreadable or unsupported images, placeholders) were left out */
	previewDocxPdfSimplified: "Some pictures could not be included in the PDF",
	/** Toast after the PDF was downloaded, when some text was drawn without its font's shaping (ligatures, joined or reordered letters) because the font asked for more work than allowed */
	previewDocxPdfUnshaped: "Some text in the PDF is drawn without its font's letter shaping",
	/** Toast when the PDF export failed */
	previewDocxPdfFailed: "The PDF could not be made",
	/** Toast when the PDF export failed because a password-protected document's password cannot protect a PDF (PDF passwords may not mix right-to-left and left-to-right letters, nor hold certain control or private characters) */
	previewDocxPdfPasswordUnsupported: "The PDF could not be made: this document's password cannot protect a PDF",
	/** Progress toast description while the editor shows markup (tracked changes, comments): the PDF shows the final text without it */
	previewDocxPdfFinalOnly: "The PDF shows the final text, without markup"
} as const
