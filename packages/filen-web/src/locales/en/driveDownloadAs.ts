// "drive" namespace, Download as in the item menus (a document or spreadsheet exported to another format without
// opening it; a local download, the file in the drive is unchanged). Spread into drive.ts; same catalog rules.
export const driveDownloadAs = {
	/** Item menu and bulk menus — submenu of the formats a document or spreadsheet can be downloaded as (a local download made from the saved file; the file in the drive is unchanged) */
	driveActionDownloadAs: "Download as",
	/** Item menu → Download as — item downloading the document or spreadsheet as a PDF file */
	driveDownloadAsPdf: "PDF document (.pdf)",
	/** Download as progress toast — the file is being downloaded before it is converted */
	driveDownloadAsDownloading: "Downloading…",
	/** Download as progress toast description — the file being exported, when several are exported one after another: {{index}} of {{count}} */
	driveDownloadAsBulkItem: "{{name}} ({{index}} of {{count}})",
	/** Download as error — a ".dot" file that turned out to be a text file (a Graphviz graph), not a Word 97–2003 template, so there is no document to export */
	driveDownloadAsNotDocument: "This file is a text file, not a Word document",
	/** Download as of several files — report toast once the run ended: how many of the files gave no PDF (each listed below with why) */
	driveDownloadAsReportFailed_one: "{{count}} of {{total}} files could not be made into a PDF",
	driveDownloadAsReportFailed_other: "{{count}} of {{total}} files could not be made into a PDF",
	/** Download as of several files — report toast once the run ended: how many PDFs lack something of their file (each listed below with what) */
	driveDownloadAsReportWarned_one: "{{count}} PDF is not quite like its file",
	driveDownloadAsReportWarned_other: "{{count}} PDFs are not quite like their files",
	/** Download as of several files — last line of the report toast: how many more files it does not list */
	driveDownloadAsReportMore: "And {{count}} more",
	/** Download as — password prompt body for a password-protected file: the PDF made from it is protected with the same password */
	driveDownloadAsPasswordBody: "“{{name}}” is password-protected. The PDF is protected with the same password. Enter it to continue."
} as const
