// "preview" namespace, Word document editor — headers, footers, sections and page setup while editing. Spread into preview.ts; same catalog rules.
export const previewDocxSections = {
	/** Docx editor — Insert menu item: a section break (submenu of kinds) */
	previewDocxSectionBreak: "Section break",
	/** Docx editor — section break kind: the next section starts on a new page */
	previewDocxSectionBreakNextPage: "Next page",
	/** Docx editor — section break kind: the next section starts on the same page */
	previewDocxSectionBreakContinuous: "Continuous",
	/** Docx editor — section break kind: the next section starts on the next even page */
	previewDocxSectionBreakEvenPage: "Even page",
	/** Docx editor — section break kind: the next section starts on the next odd page */
	previewDocxSectionBreakOddPage: "Odd page",
	/** Docx editor — Insert menu item: edits the page's header */
	previewDocxHeaderEdit: "Header",
	/** Docx editor — Insert menu item: edits the page's footer */
	previewDocxFooterEdit: "Footer",
	/** Docx editor — header chip toggle: the section's first page has its own header and footer */
	previewDocxHeaderDifferentFirst: "Different first page",
	/** Docx editor — header chip toggle: even and odd pages have their own headers and footers */
	previewDocxHeaderDifferentOddEven: "Different odd and even pages",
	/** Docx editor — header chip toggle: the section shows the previous section's header ("Same as previous") */
	previewDocxHeaderLinkToPrevious: "Link to previous",
	/** Docx editor — header chip action: removes the header */
	previewDocxHeaderRemove: "Remove header",
	/** Docx editor — header chip action: removes the footer */
	previewDocxFooterRemove: "Remove footer",
	/** Docx editor — floating chip while a header is edited; {{section}} is the section's number */
	previewDocxHeaderChip: "Header — Section {{section}}",
	/** Docx editor — as previewDocxHeaderChip on a section's first page */
	previewDocxHeaderChipFirst: "First page header — Section {{section}}",
	/** Docx editor — as previewDocxHeaderChip on an even page */
	previewDocxHeaderChipEven: "Even page header — Section {{section}}",
	/** Docx editor — as previewDocxHeaderChip on an odd page of a document with different odd and even pages */
	previewDocxHeaderChipOdd: "Odd page header — Section {{section}}",
	/** Docx editor — floating chip while a footer is edited; {{section}} is the section's number */
	previewDocxFooterChip: "Footer — Section {{section}}",
	/** Docx editor — as previewDocxFooterChip on a section's first page */
	previewDocxFooterChipFirst: "First page footer — Section {{section}}",
	/** Docx editor — as previewDocxFooterChip on an even page */
	previewDocxFooterChipEven: "Even page footer — Section {{section}}",
	/** Docx editor — as previewDocxFooterChip on an odd page of a document with different odd and even pages */
	previewDocxFooterChipOdd: "Odd page footer — Section {{section}}",
	/** Docx editor — as previewDocxHeaderChip in a document with one section */
	previewDocxHeaderChipOnly: "Header",
	/** Docx editor — as previewDocxHeaderChipFirst in a document with one section */
	previewDocxHeaderChipFirstOnly: "First page header",
	/** Docx editor — as previewDocxHeaderChipEven in a document with one section */
	previewDocxHeaderChipEvenOnly: "Even page header",
	/** Docx editor — as previewDocxHeaderChipOdd in a document with one section */
	previewDocxHeaderChipOddOnly: "Odd page header",
	/** Docx editor — as previewDocxFooterChip in a document with one section */
	previewDocxFooterChipOnly: "Footer",
	/** Docx editor — as previewDocxFooterChipFirst in a document with one section */
	previewDocxFooterChipFirstOnly: "First page footer",
	/** Docx editor — as previewDocxFooterChipEven in a document with one section */
	previewDocxFooterChipEvenOnly: "Even page footer",
	/** Docx editor — as previewDocxFooterChipOdd in a document with one section */
	previewDocxFooterChipOddOnly: "Odd page footer",
	/** Docx editor — header chip: shown while the section shows the previous section's header */
	previewDocxHeaderSameAsPrevious: "Same as previous",
	/** Docx editor — header chip menu button holding its toggles */
	previewDocxHeaderOptions: "Header and footer options",
	/** Docx editor — header chip button going back to the body (Esc): its visible text */
	previewDocxHeaderClose: "Close",
	/** Docx editor — accessible name of the header chip's Close button while a header is edited */
	previewDocxHeaderCloseHeader: "Close header",
	/** Docx editor — accessible name of the header chip's Close button while a footer is edited */
	previewDocxHeaderCloseFooter: "Close footer",
	/** Docx editor — confirm dialog title before Link to previous deletes the section's own header and footer */
	previewDocxHeaderLinkTitle: "Link to the previous section?",
	/** Docx editor — confirm dialog body before Link to previous deletes the section's own header (Word asks the same) */
	previewDocxHeaderLinkBody: "This section's own header is deleted, and it shows the previous section's instead.",
	/** Docx editor — as previewDocxHeaderLinkBody while a footer is edited */
	previewDocxFooterLinkBody: "This section's own footer is deleted, and it shows the previous section's instead.",
	/** Docx editor — confirm dialog button: links to the previous section */
	previewDocxHeaderLinkConfirm: "Link",
	/** Docx editor — ⋯ menu item opening the Page setup dialog, and its title */
	previewDocxPageSetup: "Page setup",
	/** Docx editor — ⋯ menu item opening the Page setup dialog */
	previewDocxPageSetupAction: "Page setup…",
	/** Docx editor — Page setup dialog description */
	previewDocxPageSetupBody: "Paper, margins, columns and numbering of the pages.",
	/** Docx editor — Page setup field: paper size */
	previewDocxPaperSize: "Paper size",
	/** Docx editor — paper size */
	previewDocxPaperLetter: "Letter",
	/** Docx editor — paper size */
	previewDocxPaperLegal: "Legal",
	/** Docx editor — paper size */
	previewDocxPaperExecutive: "Executive",
	/** Docx editor — paper size */
	previewDocxPaperTabloid: "Tabloid",
	/** Docx editor — paper size */
	previewDocxPaperA3: "A3",
	/** Docx editor — paper size */
	previewDocxPaperA4: "A4",
	/** Docx editor — paper size */
	previewDocxPaperA5: "A5",
	/** Docx editor — paper size or margins typed in */
	previewDocxCustom: "Custom",
	/** Docx editor — Page setup field: paper width; {{unit}} is in or cm */
	previewDocxPaperWidth: "Width ({{unit}})",
	/** Docx editor — Page setup field: paper height; {{unit}} is in or cm */
	previewDocxPaperHeight: "Height ({{unit}})",
	/** Docx editor — Page setup field: orientation */
	previewDocxOrientation: "Orientation",
	/** Docx editor — orientation option */
	previewDocxPortrait: "Portrait",
	/** Docx editor — orientation option */
	previewDocxLandscape: "Landscape",
	/** Docx editor — Page setup field: margins preset */
	previewDocxMargins: "Margins",
	/** Docx editor — margins preset (1" all round) */
	previewDocxMarginsNormal: "Normal",
	/** Docx editor — margins preset (0.5" all round) */
	previewDocxMarginsNarrow: "Narrow",
	/** Docx editor — margins preset (1" top and bottom, 0.75" sides) */
	previewDocxMarginsModerate: "Moderate",
	/** Docx editor — margins preset (1" top and bottom, 2" sides) */
	previewDocxMarginsWide: "Wide",
	/** Docx editor — margins preset (1" top and bottom, 1.25" sides) */
	previewDocxMarginsOffice2003: "Office 2003 Default",
	/** Docx editor — margins preset for facing pages (1" top, bottom and outside, 1.25" inside) */
	previewDocxMarginsMirrored: "Mirrored",
	/** Docx editor — Page setup switch: inside and outside margins swap on facing pages (Word's "Mirror margins") */
	previewDocxMirrorMargins: "Mirror margins",
	/** Docx editor — Page setup margin field with mirror margins: the margin at the binding; {{unit}} is in or cm */
	previewDocxMarginInside: "Inside ({{unit}})",
	/** Docx editor — Page setup margin field with mirror margins: the margin away from the binding; {{unit}} is in or cm */
	previewDocxMarginOutside: "Outside ({{unit}})",
	/** Docx editor — Page setup margin field; {{unit}} is in or cm */
	previewDocxMarginTop: "Top ({{unit}})",
	/** Docx editor — Page setup margin field; {{unit}} is in or cm */
	previewDocxMarginBottom: "Bottom ({{unit}})",
	/** Docx editor — Page setup margin field (left in left-to-right text); {{unit}} is in or cm */
	previewDocxMarginLeft: "Left ({{unit}})",
	/** Docx editor — Page setup margin field; {{unit}} is in or cm */
	previewDocxMarginRight: "Right ({{unit}})",
	/** Docx editor — Page setup field: text columns */
	previewDocxColumns: "Columns",
	/** Docx editor — Page setup field: the first page number of the section */
	previewDocxPageNumberStart: "Start page numbers at",
	/** Docx editor — Page setup option: page numbers continue from the previous section */
	previewDocxPageNumberContinue: "Continue from previous section",
	/** Docx editor — Page setup field: page number format */
	previewDocxPageNumberFormat: "Page number format",
	/** Docx editor — page number format */
	previewDocxNumberDecimal: "1, 2, 3",
	/** Docx editor — page number format */
	previewDocxNumberUpperRoman: "I, II, III",
	/** Docx editor — page number format */
	previewDocxNumberLowerRoman: "i, ii, iii",
	/** Docx editor — page number format */
	previewDocxNumberUpperLetter: "A, B, C",
	/** Docx editor — page number format */
	previewDocxNumberLowerLetter: "a, b, c",
	/** Docx editor — page number format */
	previewDocxNumberDash: "- 1 -, - 2 -, - 3 -",
	/** Docx editor — page number format the document has that can't be chosen here */
	previewDocxNumberOther: "As it is",
	/** Docx editor — Page setup field: line numbers in the margin */
	previewDocxLineNumbers: "Line numbers",
	/** Docx editor — line numbers option */
	previewDocxLineNumbersNone: "None",
	/** Docx editor — line numbers option */
	previewDocxLineNumbersContinuous: "Continuous",
	/** Docx editor — line numbers option */
	previewDocxLineNumbersPage: "Restart each page",
	/** Docx editor — line numbers option */
	previewDocxLineNumbersSection: "Restart each section",
	/** Docx editor — Page setup field: which sections change */
	previewDocxApplyTo: "Apply to",
	/** Docx editor — apply to option */
	previewDocxApplySection: "This section",
	/** Docx editor — apply to option */
	previewDocxApplyDocument: "Whole document",
	/** Docx editor — Page setup error: a size or margin isn't a valid length */
	previewDocxPageSetupInvalid: "Some sizes aren't valid.",
	/** Docx editor — Page setup error when the document refused the settings (the dialog stays open) */
	previewDocxPageSetupRefused: "These settings can't be saved in this document. Try other values."
} as const
