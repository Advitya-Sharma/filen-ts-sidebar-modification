// English source catalog — "preview" namespace: the full-bleed drive file-preview overlay
// (features/preview/components/*) opened from a listing's handleOpen — header chrome (prev/next, download,
// close) and per-category viewer copy. Same typed-catalog rules as the other namespaces: flat `as
// const` object, camelCase keys, no literal '.' or ':' (real i18next namespaces, keySeparator/
// nsSeparator both ON).
export const preview = {
	// ── Header ───────────────────────────────────────────────────────────────
	/** Preview overlay header — accessible label for the button stepping to the previous previewable sibling; disabled at the first item */
	previewPreviousAction: "Previous file",
	/** Preview overlay header — accessible label for the button stepping to the next previewable sibling; disabled at the last item */
	previewNextAction: "Next file",
	/** Preview overlay header — accessible label for the button downloading the open item; hidden in trash */
	previewDownloadAction: "Download",
	/** Preview overlay header — accessible label for the button saving the editable text/code buffer; shown only while editable and dirty, also the Cmd/Ctrl+S keymap command's description */
	previewSaveAction: "Save",
	/** Shortcuts list — opens the search of the text editors (file editor and note editors) and the find field of the Word document viewer */
	previewEditorFind: "Find",
	/** Shortcuts list — opens the text editor's search with its replace field */
	previewEditorReplace: "Find and replace",
	/** Shortcuts list — wraps the selection in markdown bold (markdown editors only) */
	previewEditorBold: "Bold (markdown)",
	/** Shortcuts list — wraps the selection in markdown italics (markdown editors only) */
	previewEditorItalic: "Italic (markdown)",
	/** Shortcuts list — wraps the selection in markdown strikethrough (markdown editors only) */
	previewEditorStrikethrough: "Strikethrough (markdown)",
	/** Shortcuts list — wraps the selection in markdown inline code (markdown editors only) */
	previewEditorCode: "Inline code (markdown)",
	/** Shortcuts list — turns the selection into a markdown link (markdown editors only) */
	previewEditorLink: "Insert link (markdown)",
	/** Shortcuts list — switches a markdown file between its rendered and source views, or shows/hides a markdown note's preview pane */
	previewEditorTogglePreview: "Toggle markdown preview",

	// ── Body ─────────────────────────────────────────────────────────────────
	/** Preview overlay body — shown in place of a viewer for a previewable category with no renderer built yet */
	previewUnsupportedType: "Preview isn't available for this file type yet.",
	/** Preview overlay body — a file opened through "Open as text" turned out to be binary, so it is not shown */
	previewNotText: "This file isn't text, so it can't be shown as text.",
	/** Preview overlay body — a streamed image/video/audio failed mid-playback and the file is too large to safely retry buffered */
	previewStreamFailed: "This preview failed and the file is too large to retry.",
	/** Preview overlay body — a HEIC/HEIF image could not be converted for preview (corrupt or unsupported file) */
	previewTransformFailed: "This image couldn't be converted for preview.",
	/** Preview overlay body — a camera RAW file (.cr2, .nef, .arw, .dng…) embeds no preview image large enough to show, and no thumbnail could be made for it either */
	previewRawNoPreview: "This RAW image has no embedded preview to show.",
	/** Preview overlay body — shown by the scoped error boundary when a viewer throws while rendering (e.g. a parse failure) */
	previewRenderError: "This preview couldn't be displayed.",

	// ── Editable save ────────────────────────────────────────────────────────
	/** Unsaved-changes confirm dialog — title, shown on Escape/close/prev/next while the editable buffer is dirty */
	previewUnsavedChangesTitle: "Unsaved changes",
	/** Unsaved-changes confirm dialog — body */
	previewUnsavedChangesBody: "You have unsaved changes to this file. Discard them?",
	/** Unsaved-changes confirm dialog — the destructive confirm button, discards the buffer and proceeds */
	previewDiscardAction: "Discard",
	/** Toast shown after a save fails because the file's own parent directory no longer exists — the editor locks read-only for the rest of this session (mobile parity: retrying against the same broken parent would only fail again); every other save failure keeps the editor open for a retry instead */
	/** Toast — the editor's content could not be read for saving (the spreadsheet could not be written out) */
	previewSaveFailed: "This file couldn't be saved.",
	previewReadOnlyAfterSaveFailure: "This file is now read-only — saving failed and can't be retried until you reopen it.",

	// ── Changes made elsewhere ───────────────────────────────────────────────
	/** Remote-change dialog — title, shown while the open file has unsaved edits and a newer version of it is saved from another device, another tab, or restored from its version history */
	previewRemoteChangedTitle: "This file changed elsewhere",
	/** Remote-change dialog — body; {{name}} is the file name */
	previewRemoteChangedBody:
		"A newer version of {{name}} was saved while you were editing it. What should happen to your unsaved changes?",
	/** Remote-change dialog — opens a side-by-side comparison of the newer version and the unsaved changes (text, code and markdown files only) */
	previewRemoteCompare: "Compare",
	/** Remote-change dialog — leaves the comparison, back to the plain prompt */
	previewRemoteCompareBack: "Back",
	/** Remote-change dialog — writes the unsaved changes to a new file next to this one, then shows the newer version */
	previewRemoteSaveCopy: "Save mine as copy",
	/** Remote-change dialog — keeps editing the unsaved changes; saving them then makes them the newest version */
	previewRemoteKeepMine: "Keep mine",
	/** Remote-change dialog — discards the unsaved changes and shows the newer version */
	previewRemoteLoadTheirs: "Load theirs",
	/** Remote-change comparison — heading over the newer version saved elsewhere (left side) */
	previewRemoteTheirs: "Their version",
	/** Remote-change comparison — heading over the unsaved changes (right side) */
	previewRemoteMine: "Your version",
	/** Remote-change comparison on a narrow screen, where both versions are interleaved in one column */
	previewRemoteUnifiedLegend: "Red lines are only in their version, green lines only in yours.",
	/** Remote-change comparison — the newer version could not be downloaded */
	previewRemoteCompareFailed: "Their version couldn't be loaded for comparison.",
	/** Remote-deletion dialog — title, shown while the open file has unsaved edits and it is trashed or deleted elsewhere */
	previewRemoteDeletedTitle: "This file was deleted elsewhere",
	/** Remote-deletion dialog — body; {{name}} is the file name */
	previewRemoteDeletedBody: "{{name}} was moved to the trash or deleted while you were editing it. Your changes are still here.",
	/** Remote-deletion dialog — discards the unsaved changes and closes the file */
	previewRemoteDiscardMine: "Discard my changes",
	/** Remote-deletion dialog — writes the unsaved changes to a new file in the same directory */
	previewRemoteSaveAsNew: "Save as new file",
	/** Remote-change dialog, when the open viewer is read-only (the file was renamed to another type, or can't be saved): why no save is offered */
	previewRemoteSaveUnavailable: "Your changes can't be saved from this preview. Keep them on screen to copy them, or discard them.",
	/** Remote-deletion dialog, when the changes can't be saved: closes the dialog and leaves the changes on screen */
	previewRemoteKeepOpen: "Keep on screen",
	/** Name of the file the unsaved changes are written to when the original moved on; {{base}} is the original name without its extension, {{ext}} the extension with its dot (may be empty), {{date}} the current date and time */
	previewConflictCopyName: "{{base}} (conflicted copy {{date}}){{ext}}",
	/** Toast after the unsaved changes were written to a new file; {{name}} is its name */
	previewSavedAsNewFile: "Saved your changes as {{name}}.",
	/** Toast when writing the unsaved changes to a new file failed */
	previewSaveAsNewFileFailed: "Your changes couldn't be saved as a new file.",
	/** Toast when the open file, with no unsaved edits, was replaced by a newer version saved elsewhere and now shows it */
	previewUpdatedElsewhere: "Updated with changes saved elsewhere.",
	/** Toast when the open file, with unsaved edits, was moved to another directory elsewhere */
	previewMovedElsewhere: "This file was moved to another directory. Your changes will be saved there.",
	/** Toast when a save from this editor landed right after a version saved elsewhere, which it replaced */
	previewSaveReplacedNewer:
		"Your save replaced changes saved elsewhere moments before. With file versioning on, they're in the file's version history.",

	// ── Spreadsheet ──────────────────────────────────────────────────────────
	/** Spreadsheet viewer — a workbook with no sheet to show */
	previewSpreadsheetEmpty: "This spreadsheet has no sheets.",
	/** Spreadsheet viewer — the file could not be read as a spreadsheet (damaged, encrypted with a password, or too large to open) */
	previewSpreadsheetUnreadable: "This spreadsheet couldn't be opened. It may be damaged, password-protected or too large.",
	/** Spreadsheet viewer — accessible label for the box showing the selected cell or range, e.g. "B2:D7" */
	previewSpreadsheetSelectedCells: "Selected cells",
	/** Spreadsheet viewer — accessible label for the bar showing the active cell's contents or formula */
	previewSpreadsheetCellContents: "Cell contents",
	/** Spreadsheet viewer — accessible label for the row of sheet tabs */
	previewSpreadsheetSheets: "Sheets",
	/** Spreadsheet viewer — toast when the selection is too large to copy */
	previewSpreadsheetCopyTooLarge: "That selection is too large to copy.",
	/** Spreadsheet toolbar — undo the last edit (also Cmd/Ctrl+Z in the grid) */
	previewSpreadsheetUndo: "Undo",
	/** Spreadsheet toolbar — redo the last undone edit (also Cmd/Ctrl+Shift+Z or Cmd/Ctrl+Y) */
	previewSpreadsheetRedo: "Redo",
	/** Spreadsheet toolbar — toggle bold on the selected cells */
	previewSpreadsheetBold: "Bold",
	/** Spreadsheet toolbar — toggle italic on the selected cells */
	previewSpreadsheetItalic: "Italic",
	/** Spreadsheet toolbar — toggle underline on the selected cells */
	previewSpreadsheetUnderline: "Underline",
	/** Spreadsheet toolbar — toggle strikethrough on the selected cells */
	previewSpreadsheetStrikethrough: "Strikethrough",
	/** Spreadsheet toolbar — opens the palette for the selected cells' text colour */
	previewSpreadsheetTextColor: "Text colour",
	/** Spreadsheet toolbar — opens the palette for the selected cells' fill (background) colour */
	previewSpreadsheetFillColor: "Fill colour",
	/** Spreadsheet toolbar colour palette — removes the colour */
	previewSpreadsheetColorNone: "No colour",
	/** Spreadsheet toolbar — align the selected cells' content left */
	previewSpreadsheetAlignLeft: "Align left",
	/** Spreadsheet toolbar — centre the selected cells' content */
	previewSpreadsheetAlignCenter: "Centre",
	/** Spreadsheet toolbar — align the selected cells' content right */
	previewSpreadsheetAlignRight: "Align right",
	/** Spreadsheet toolbar — accessible label of the number format menu */
	previewSpreadsheetNumberFormat: "Number format",
	/** Number format menu — no particular format */
	previewSpreadsheetFormatGeneral: "General",
	/** Number format menu — 1,234.56 */
	previewSpreadsheetFormatNumber: "Number",
	/** Number format menu — $1,234.56 */
	previewSpreadsheetFormatDollar: "Currency ($)",
	/** Number format menu — 1,234.56 € */
	previewSpreadsheetFormatEuro: "Currency (€)",
	/** Number format menu — 12.34% */
	previewSpreadsheetFormatPercent: "Percent",
	/** Number format menu — 2026-09-27 */
	previewSpreadsheetFormatDate: "Date",
	/** Number format menu — the value kept as text */
	previewSpreadsheetFormatText: "Text",
	/** Number format menu — the cell carries a format not in this list */
	previewSpreadsheetFormatCustom: "Custom",
	/** Spreadsheet grid menu — insert as many rows as are selected, above them */
	previewSpreadsheetInsertRowsAbove_one: "Insert row above",
	previewSpreadsheetInsertRowsAbove_other: "Insert {{count}} rows above",
	/** Spreadsheet grid menu — insert as many rows as are selected, below them */
	previewSpreadsheetInsertRowsBelow_one: "Insert row below",
	previewSpreadsheetInsertRowsBelow_other: "Insert {{count}} rows below",
	/** Spreadsheet grid menu — delete the selected rows */
	previewSpreadsheetDeleteRows_one: "Delete row",
	previewSpreadsheetDeleteRows_other: "Delete {{count}} rows",
	/** Spreadsheet grid menu — insert as many columns as are selected, to their left */
	previewSpreadsheetInsertColumnsLeft_one: "Insert column left",
	previewSpreadsheetInsertColumnsLeft_other: "Insert {{count}} columns left",
	/** Spreadsheet grid menu — insert as many columns as are selected, to their right */
	previewSpreadsheetInsertColumnsRight_one: "Insert column right",
	previewSpreadsheetInsertColumnsRight_other: "Insert {{count}} columns right",
	/** Spreadsheet grid menu — delete the selected columns */
	previewSpreadsheetDeleteColumns_one: "Delete column",
	previewSpreadsheetDeleteColumns_other: "Delete {{count}} columns",
	/** Spreadsheet grid menu — empty the selected cells (formats stay) */
	previewSpreadsheetClearCells: "Clear contents",
	/** Spreadsheet grid menu — give the selected columns back their default (or file's own) width */
	previewSpreadsheetResetColumnWidth_one: "Reset column width",
	previewSpreadsheetResetColumnWidth_other: "Reset {{count}} column widths",
	/** Spreadsheet grid menu — give the selected rows back their default (or file's own) height */
	previewSpreadsheetResetRowHeight_one: "Reset row height",
	previewSpreadsheetResetRowHeight_other: "Reset {{count}} row heights",
	/** Spreadsheet viewer note — a legacy .xls opens read-only */
	previewSpreadsheetReadOnlyXls: "Old .xls format — view only.",
	/** Spreadsheet viewer — converts a read-only .xls into an editable .xlsx copy beside it */
	previewSpreadsheetSaveAsXlsx: "Save as .xlsx",
	/** Spreadsheet viewer — confirm dialog title before converting an .xls */
	previewSpreadsheetSaveAsXlsxTitle: "Save as .xlsx?",
	/** Spreadsheet viewer — confirm dialog body: what the .xls to .xlsx conversion keeps and loses */
	previewSpreadsheetSaveAsXlsxBody:
		"An editable .xlsx copy is saved in the same directory, and the .xls stays as it is. Only the values come along: formulas become their results, and formatting, column widths, row heights and named ranges are not kept.",
	/** Spreadsheet viewer — toast once the .xlsx copy is saved */
	previewSpreadsheetSavedAsXlsx: "Saved as {{name}}",
	/** Spreadsheet viewer — toast when the .xls could not be converted */
	previewSpreadsheetSaveAsXlsxFailed: "This file could not be saved as .xlsx",
	/** Spreadsheet grid — the size tip shown while dragging a column's edge */
	previewSpreadsheetWidthPx: "Width: {{size}} px",
	/** Spreadsheet grid — the size tip shown while dragging a row's edge */
	previewSpreadsheetHeightPx: "Height: {{size}} px",
	/** Spreadsheet grid — tooltip of a column header's resize edge */
	previewSpreadsheetResizeColumn: "Resize column",
	/** Spreadsheet grid — tooltip of a row header's resize edge */
	previewSpreadsheetResizeRow: "Resize row",
	/** Spreadsheet sheet tabs — adds a new empty sheet */
	previewSpreadsheetAddSheet: "Add sheet",
	/** Name given to a newly added sheet; {{number}} is its position */
	previewSpreadsheetNewSheetName: "Sheet {{number}}",
	/** Rename-sheet dialog — title and submit button */
	previewSpreadsheetRenameSheet: "Rename sheet",
	/** Rename-sheet dialog — body */
	previewSpreadsheetRenameSheetBody: "Formulas that refer to this sheet are updated to the new name.",
	/** Rename-sheet dialog — field label */
	previewSpreadsheetSheetName: "Sheet name",
	/** Toast — rows or columns can't be inserted or deleted, or a sheet renamed, in this workbook */
	previewSpreadsheetStructureLocked:
		"This can't be changed here: the workbook has charts, tables, conditional formats, validations or named ranges that depend on its rows, columns and sheet names.",
	/** Toast — a sheet name is taken or has characters Excel doesn't allow (\ / ? * [ ] :, at most 31 characters) */
	previewSpreadsheetSheetNameInvalid: "That sheet name is already used or isn't allowed.",
	/** Toast — an edit covers only part of an array formula's range */
	previewSpreadsheetArrayFormula: "Part of an array formula can't be changed. Select its whole range to replace it.",
	/** Toast — the file keeps a legacy text encoding that can't hold a character in the edit */
	previewSpreadsheetEncodingUnsupported:
		"This file uses an older text encoding that can't hold one of those characters, so the change wasn't made.",
	/** Toast — an edit would change the text of a table's header cell */
	previewSpreadsheetTableHeader: "A table's header can't be renamed here: formulas that refer to the column by name would break.",
	/** Toast — an edit reaches past what can be changed at once, or past a sheet's size */
	previewSpreadsheetTooLarge: "That's more than can be changed at once.",
	/** Toast — an edit failed unexpectedly */
	previewSpreadsheetEditFailed: "That change couldn't be made.",
	/** Spreadsheet viewer — note beside the cell contents when an editable file opens read-only because it holds parts (e.g. chart sheets) a save could damage */
	previewSpreadsheetReadOnlyUnsafe: "Read-only here: this file has parts that can't be saved safely.",
	/** Spreadsheet viewer — note beside the cell contents when a file opens read-only because its name has no spreadsheet extension */
	previewSpreadsheetReadOnlyUnnamed: "Read-only here: the file name doesn't say which spreadsheet format it is.",
	/** Spreadsheet viewer — note beside the cell contents when an open file was renamed to another format's extension, so saving it under that name would mislabel it */
	previewSpreadsheetReadOnlyRenamed: "Read-only here: the file name no longer matches its format.",
	/** Spreadsheet viewer — status beside the cell contents while an Excel file opens view-only until it is confirmed that saving it keeps everything in it */
	previewSpreadsheetCheckingWritable: "Checking this file can be saved…",
	/** Spreadsheet viewer — note beside the cell contents when an Excel file stays read-only because saving it here would lose some of its content */
	previewSpreadsheetReadOnlyLossy: "Read-only here: this file has content this app can't save without losing it.",
	/** Spreadsheet viewer — note beside the cell contents when an Excel file opens read-only because it is too large to check that saving it keeps everything in it */
	previewSpreadsheetReadOnlyTooLarge: "Read-only here: this file is too large to edit in the browser.",
	/** Spreadsheet viewer — shown beside the formula bar (and as a toast for a paste) when an entry starts with "=" but isn't a formula Excel can read, e.g. a parenthesis left open or ";" between arguments */
	previewSpreadsheetFormulaInvalid: "This formula can't be read. Check its parentheses and quotes, and separate arguments with commas.",
	/** Spreadsheet grid — below a sheet too tall to show whole; {{rows}} is the number of rows shown, already formatted */
	previewSpreadsheetRowsTruncated: "Showing the first {{rows}} rows. Rows further down can't be shown here.",

	// ── PDF ──────────────────────────────────────────────────────────────────
	/** PDF viewer — password dialog title, shown both on the first prompt and on a wrong-password retry */
	previewPdfPasswordTitle: "Password required",
	/** PDF viewer — password dialog body on the first prompt, before any attempt */
	previewPdfPasswordBody: "This PDF is password-protected. Enter the password to view it.",
	/** PDF viewer — password dialog body after a submitted password was rejected */
	previewPdfPasswordRetryBody: "That password was incorrect. Try again.",
	/** PDF viewer — password dialog field label */
	previewPdfPasswordLabel: "Password",
	/** PDF viewer — password dialog submit button */
	previewPdfPasswordSubmit: "Unlock",
	/** PDF viewer — button shown after the password prompt is dismissed, reopens it */
	previewPdfPasswordReopen: "Enter password",
	/** PDF viewer — shown in place of the document when it fails to load or render (corrupt file, decode error) */
	previewPdfLoadFailed: "This PDF couldn't be rendered.",
	/** PDF viewer — accessible label for the page-nav button stepping to the previous page; distinct from the overlay's own file-level previous/next */
	previewPdfPreviousPageAction: "Previous page",
	/** PDF viewer — accessible label for the page-nav button stepping to the next page */
	previewPdfNextPageAction: "Next page",
	/** PDF viewer — page-nav indicator and per-page canvas label; {{current}} = the page in view, {{total}} = page count */
	previewPdfPageIndicator: "Page {{current}} of {{total}}",
	/** PDF viewer — accessible label for the toolbar button stepping the zoom level down */
	previewPdfZoomOutAction: "Zoom out",
	/** PDF viewer — accessible label for the toolbar button stepping the zoom level up */
	previewPdfZoomInAction: "Zoom in",
	/** PDF viewer — zoom-level indicator between the zoom out/in buttons; {{percent}} = current zoom relative to the document's own default render scale */
	previewPdfZoomIndicator: "{{percent}}%",

	// ── Docx ─────────────────────────────────────────────────────────────────
	/** Docx viewer — shown in place of the document when it fails to load or render */
	previewDocxLoadFailed: "This document couldn't be rendered.",
	/** Docx viewer — shown in place of a password-protected (encrypted) Word document, which can't be opened here */
	previewDocxEncrypted: "This document is password-protected and can't be opened here.",
	/** Docx viewer — shown in place of an older binary Word document (.doc) saved with a .docx name */
	previewDocxLegacyFormat: "This document uses an older Word format that can't be opened here.",
	/** Docx viewer — shown when the file is not a Word document at all (another Office file, or not a document) */
	previewDocxNotDocument: "This file isn't a Word document.",
	/** Docx viewer — accessible label for the page-nav button stepping to the previous page; distinct from the overlay's own file-level previous/next */
	previewDocxPreviousPageAction: "Previous page",
	/** Docx viewer — accessible label for the page-nav button stepping to the next page */
	previewDocxNextPageAction: "Next page",
	/** Docx viewer — page-nav indicator and each page's accessible label; {{current}} = the page in view, {{total}} = page count */
	previewDocxPageIndicator: "Page {{current}} of {{total}}",
	/** Docx viewer — accessible label for the toolbar button stepping the zoom level down */
	previewDocxZoomOutAction: "Zoom out",
	/** Docx viewer — accessible label for the toolbar button stepping the zoom level up */
	previewDocxZoomInAction: "Zoom in",
	/** Docx viewer — zoom-level indicator between the zoom out/in buttons; {{percent}} = zoom relative to the document's actual size */
	previewDocxZoomIndicator: "{{percent}}%",
	/** Docx viewer — tooltip of the zoom-level indicator, which is a button: clicking it fits the page to the window's width again */
	previewDocxFitWidth: "Fit to width",
	/** Docx viewer — accessible label of the zoom-level button; {{percent}} = the current zoom */
	previewDocxFitWidthAction: "Zoom {{percent}}%, fit to width",
	/** Docx viewer — toolbar button (and its tooltip) opening the find field in the toolbar; also the Cmd/Ctrl+F shortcut */
	previewDocxFindAction: "Find in document",
	/** Docx viewer — placeholder and accessible label of the find field in the toolbar */
	previewDocxFindPlaceholder: "Find",
	/** Docx viewer — match counter beside the find field; {{current}} = the selected match, {{total}} = how many the document holds */
	previewDocxFindCount: "{{current, number}}/{{total, number}}",
	/** Docx viewer — as previewDocxFindCount when there are more matches than a find shows ({{total}}, the most it shows) */
	previewDocxFindCountCapped: "{{current, number}}/{{total, number}}+",
	/** Docx viewer — announced to screen readers as the find results change; {{current}} = the selected match, {{total}} = all matches */
	previewDocxFindStatus: "Match {{current, number}} of {{total, number}}",
	/** Docx viewer — as previewDocxFindStatus when there are more matches than a find shows ({{total}}, the most it shows) */
	previewDocxFindStatusCapped: "Match {{current, number}} of more than {{total, number}}",
	/** Docx viewer — announced to screen readers when the search text is nowhere in the document */
	previewDocxFindNoMatches: "No matches",
	/** Docx viewer — find bar button selecting the previous match (also Shift+Enter in the field) */
	previewDocxFindPrevious: "Previous match",
	/** Docx viewer — find bar button selecting the next match (also Enter in the field) */
	previewDocxFindNext: "Next match",
	/** Docx viewer — find bar toggle: only matches with the same upper and lower case */
	previewDocxFindMatchCase: "Match case",
	/** Docx viewer — find bar button closing the field (also Escape) */
	previewDocxFindClose: "Close find",
	/** Docx viewer — toolbar menu (and its tooltip) choosing how comments and tracked changes show, as Word's "Display for Review" */
	previewDocxMarkup: "Markup",
	/** Docx viewer — markup menu option: changes shown inline (insertions underlined, deletions struck through), comments and formatting changes in the margin */
	previewDocxMarkupAll: "All markup",
	/** Docx viewer — markup menu option: the document with every change accepted, changed lines marked by a bar, comments in the margin */
	previewDocxMarkupSimple: "Simple markup",
	/** Docx viewer — markup menu option: the document with every change accepted and nothing marked */
	previewDocxMarkupNone: "No markup",
	/** Docx viewer — markup menu option: the document as it was before any tracked change */
	previewDocxMarkupOriginal: "Original",
	/** Docx viewer — accessible name of the margin column holding the comment and change balloons */
	previewDocxComments: "Comments and changes",
	/** Docx viewer — a comment's author when the document records none */
	previewDocxUnknownAuthor: "Unknown author",
	/** Docx viewer — button in a comment or change balloon showing its whole text (the balloon is shortened otherwise) */
	previewDocxBalloonExpand: "Show all",
	/** Docx viewer — button in an expanded balloon shortening it again */
	previewDocxBalloonCollapse: "Show less",
	/** Docx viewer — control at the bottom of a page's balloon column standing for the balloons that do not fit beside the page; opens them in a popover; {{count}} = how many */
	previewDocxMoreBalloons_one: "+{{count}} more",
	/** Docx viewer — as previewDocxMoreBalloons_one, for several */
	previewDocxMoreBalloons_other: "+{{count}} more",
	/** Docx viewer — label on a resolved comment thread (marked done in Word) */
	previewDocxCommentResolved: "Resolved",
	/** Docx viewer — accessible label of the small comment marker shown at the page edge when there is no room for balloons; {{author}} = who wrote it */
	previewDocxCommentMarker: "Comment by {{author}}",
	/** Docx viewer — accessible label of a comment balloon; {{author}} = who wrote it */
	previewDocxCommentLabel: "Comment by {{author}}",
	/** Docx viewer — balloon of a formatting change whose changed properties are not named */
	previewDocxFormatted: "Formatted",
	/** Docx viewer — balloon of a formatting change; {{properties}} = the changed properties, comma-separated (e.g. "Bold, Font size") */
	previewDocxFormattedProperties: "Formatted: {{properties}}",
	/** Docx viewer — formatting change property: bold */
	previewDocxPropertyBold: "Bold",
	/** Docx viewer — formatting change property: italic */
	previewDocxPropertyItalic: "Italic",
	/** Docx viewer — formatting change property: underline */
	previewDocxPropertyUnderline: "Underline",
	/** Docx viewer — formatting change property: strikethrough (single or double) */
	previewDocxPropertyStrikethrough: "Strikethrough",
	/** Docx viewer — formatting change property: the font (typeface) */
	previewDocxPropertyFont: "Font",
	/** Docx viewer — formatting change property: font size */
	previewDocxPropertyFontSize: "Font size",
	/** Docx viewer — formatting change property: font colour */
	previewDocxPropertyFontColor: "Font color",
	/** Docx viewer — formatting change property: text highlight colour */
	previewDocxPropertyHighlight: "Highlight",
	/** Docx viewer — formatting change property: background shading */
	previewDocxPropertyShading: "Shading",
	/** Docx viewer — formatting change property: superscript or subscript */
	previewDocxPropertyPosition: "Superscript/subscript",
	/** Docx viewer — formatting change property: all caps or small caps */
	previewDocxPropertyCaps: "Caps",
	/** Docx viewer — formatting change property: the paragraph or character style */
	previewDocxPropertyStyle: "Style",
	/** Docx viewer — formatting change property: paragraph alignment */
	previewDocxPropertyAlignment: "Alignment",
	/** Docx viewer — formatting change property: paragraph indentation */
	previewDocxPropertyIndent: "Indent",
	/** Docx viewer — formatting change property: paragraph or character spacing */
	previewDocxPropertySpacing: "Spacing",
	/** Docx viewer — formatting change property: list numbering or bullets */
	previewDocxPropertyNumbering: "Numbering",
	/** Docx viewer — formatting change property: tab stops */
	previewDocxPropertyTabs: "Tabs",
	/** Docx viewer — formatting change property: borders */
	previewDocxPropertyBorders: "Borders",
	/** Docx viewer — formatting change property: keep with next / keep lines together / page break before */
	previewDocxPropertyPagination: "Line and page breaks",
	/** Docx viewer — formatting change property: anything about a table (width, borders, layout, rows, cells) */
	previewDocxPropertyTable: "Table",
	/** Docx viewer — toolbar badge shown when the document can only be viewed; opens a popover explaining why */
	previewDocxViewOnly: "View only",
	/** Docx viewer — heading of the popover opened from the "View only" badge */
	previewDocxViewOnlyTitle: "This document can only be viewed",
	/** Docx viewer — view-only reason: the document is restricted with Word's "Restrict Editing" */
	previewDocxViewOnlyProtected: "It's protected against editing in Word.",
	/** Docx viewer — view-only reason: the document asks for a password before it can be changed */
	previewDocxViewOnlyWriteProtected: "It needs a password to be changed.",
	/** Docx viewer — view-only reason: the document is digitally signed, and any change would break the signature */
	previewDocxViewOnlySigned: "It's digitally signed; a change would break the signature.",
	/** Docx viewer — view-only reason: the document is saved in the "Strict Open XML" variant of the format */
	previewDocxViewOnlyStrict: "It's saved as Strict Open XML.",
	/** Docx viewer — view-only reason: parts of the file are damaged, so it is only shown, to keep it intact */
	previewDocxViewOnlyMalformed: "Parts of it are damaged, so it's left as it is.",
	/** Docx viewer — view-only reason: the file was opened from a public link */
	previewDocxViewOnlyPublicLink: "It was opened from a public link.",
	/** Docx viewer — view-only reason: someone else's file, opened through a link they sent (e.g. a link card in a chat) */
	previewDocxViewOnlyLinkedFile: "It's someone else's file, opened through their link.",
	/** Docx viewer — view-only reason: the file is inside an archive (zip and similar) */
	previewDocxViewOnlyArchive: "It's inside an archive.",
	/** Docx viewer — view-only reason: the file is in the trash */
	previewDocxViewOnlyTrash: "It's in the trash.",
	/** Docx viewer — view-only reason: someone else shared the file with you */
	previewDocxViewOnlySharedIn: "It was shared with you.",
	/** Docx viewer — view-only reason: opened from Recents, Favorites, Links or Shared, not from its directory in the drive */
	previewDocxViewOnlyNotInDrive: "Open it from its directory in your drive to edit it.",
	/** Docx editor — view-only reason: a text box is being edited, which this editor can only show */
	previewDocxViewOnlyTextbox: "Text boxes can't be edited here yet.",
	/** Docx editor — view-only reason: the file was renamed to another format while open */
	previewDocxViewOnlyRenamed: "It was renamed to another format while open, so changes can't be saved to it.",
	/** Docx editor — view-only reason: changes can no longer be saved to this file (its directory is gone) */
	previewDocxViewOnlyNotWritable: "Changes can't be saved to this file any more.",
	/** Docx editor — accessible name of the formatting toolbar row */
	previewDocxEditToolbar: "Formatting",
	/** Docx editor — toolbar button: undo the last change */
	previewDocxUndo: "Undo",
	/** Docx editor — toolbar button: redo the change last undone */
	previewDocxRedo: "Redo",
	/** Docx editor — toolbar button: bold text */
	previewDocxBold: "Bold",
	/** Docx editor — toolbar button: italic text */
	previewDocxItalic: "Italic",
	/** Docx editor — toolbar button: underlined text */
	previewDocxUnderline: "Underline",
	/** Docx editor — toolbar button: struck-through text */
	previewDocxStrikethrough: "Strikethrough",
	/** Docx editor — toolbar button: remove direct character formatting */
	previewDocxClearFormatting: "Clear formatting",
	/** Docx editor — accessible name of the paragraph style picker */
	previewDocxParagraphStyle: "Paragraph style",
	/** Docx editor — paragraph style picker shows this when the selected paragraphs have different styles */
	previewDocxStyleMixed: "Several styles",
	/** Docx editor — paragraph style: Word's Normal style */
	previewDocxStyleNormal: "Normal",
	/** Docx editor — paragraph style: Word's Title style */
	previewDocxStyleTitle: "Title",
	/** Docx editor — paragraph style: Word's Subtitle style */
	previewDocxStyleSubtitle: "Subtitle",
	/** Docx editor — paragraph style: a heading level */
	previewDocxStyleHeading: "Heading {{level}}",
	/** Docx editor — paragraph style: Word's Quote style */
	previewDocxStyleQuote: "Quote",
	/** Docx editor — toolbar button: align paragraph to the start (left in left-to-right text) */
	previewDocxAlignStart: "Align left",
	/** Docx editor — toolbar button: center paragraph */
	previewDocxAlignCenter: "Center",
	/** Docx editor — toolbar button: align paragraph to the end (right in left-to-right text) */
	previewDocxAlignEnd: "Align right",
	/** Docx editor — toolbar button: justify paragraph */
	previewDocxAlignJustify: "Justify",
	/** Docx editor — toolbar button: bulleted list */
	previewDocxBulletList: "Bulleted list",
	/** Docx editor — toolbar button: numbered list */
	previewDocxNumberedList: "Numbered list",
	/** Docx editor — toolbar button: move list item one level in */
	previewDocxIndent: "Increase indent",
	/** Docx editor — toolbar button: move list item one level out */
	previewDocxOutdent: "Decrease indent",
	/** Docx editor — overflow menu trigger holding tools that don't fit and editor options */
	previewDocxMoreTools: "More tools",
	/** Docx editor — overflow menu: zoom submenu label */
	previewDocxZoom: "Zoom",
	/** Docx editor — overflow menu checkbox: save changes automatically after a pause in editing */
	previewDocxAutosave: "Save automatically",
	/** Docx editor — toolbar button opening find with the replace field */
	previewDocxFindReplaceAction: "Find and replace",
	/** Docx editor — find bar toggle showing or hiding the replace field */
	previewDocxShowReplace: "Toggle replace",
	/** Docx editor — find bar: placeholder and accessible name of the replacement field */
	previewDocxReplacePlaceholder: "Replace with",
	/** Docx editor — find bar button: replace the current match */
	previewDocxReplace: "Replace",
	/** Docx editor — find bar button: replace every match */
	previewDocxReplaceAll: "Replace all",
	/** Docx editor — toast after Replace all, for one replacement */
	previewDocxReplacedCount_one: "Replaced {{count}} match",
	/** Docx editor — as previewDocxReplacedCount_one, for several */
	previewDocxReplacedCount_other: "Replaced {{count}} matches",
	/** Docx editor — context and touch menu: cut the selection */
	previewDocxCut: "Cut",
	/** Docx editor — context and touch menu: copy the selection */
	previewDocxCopy: "Copy",
	/** Docx editor — context and touch menu: paste the clipboard */
	previewDocxPaste: "Paste",
	/** Docx editor — context and touch menu: select the whole document */
	previewDocxSelectAll: "Select all",
	/** Docx editor — toast when the browser refuses access to the clipboard; names the keyboard shortcut that still works */
	previewDocxClipboardBlocked: "Your browser blocked access to the clipboard. Use {{shortcut}} instead.",
	/** Docx editor — context menu while spelling suggestions load */
	previewDocxSpellingLoading: "Looking for suggestions…",
	/** Docx editor — context menu when a misspelled word has no suggestions */
	previewDocxSpellingNone: "No suggestions",
	/** Docx editor — context menu: stop marking this word in this document */
	previewDocxSpellingIgnoreAll: "Ignore all",
	/** Docx editor — context menu: add the word to the user's own dictionary (synced across devices) */
	previewDocxSpellingAdd: "Add to dictionary",
	/** Docx editor — context menu on a word repeated right after itself: delete the repetition */
	previewDocxSpellingDeleteRepeated: "Delete repeated word",
	/** Docx editor — context menu on a repeated word: stop flagging this repetition */
	previewDocxSpellingIgnoreRepeated: "Ignore",
	/** Docx editor — toast after opening: blocks the save check could not prove safe stay as they are, for one */
	previewDocxLockedBlocks_one: "{{count}} paragraph or table can't be edited without changing how it is stored, so it stays as it is.",
	/** Docx editor — as previewDocxLockedBlocks_one, for several */
	previewDocxLockedBlocks_other:
		"{{count}} paragraphs or tables can't be edited without changing how they are stored, so they stay as they are.",
	/** Docx editor — toolbar badge while the cursor is in a part that can't be edited */
	previewDocxLocked: "Can't edit here",
	/** Docx editor — locked part popover title */
	previewDocxLockedTitle: "This part stays as it is",
	/** Docx editor — locked part reason: Word marks the content as locked */
	previewDocxLockedContent: "It's locked in Word.",
	/** Docx editor — locked part reason: editing it here could change how Word stores it */
	previewDocxLockedProof: "Editing it here could change how it's stored, so it's left untouched.",
	/** Docx editor — locked part reason: a structure the editor can't change yet (fields across paragraphs, unusual tables...) */
	previewDocxLockedStructure: "It uses a structure this editor can't change yet.",
	/** Docx editor — toast when saving refuses because the written file did not pass the check; the edits stay open and unsaved, with an action copying the document's text */
	previewDocxSaveRefused:
		"The document couldn't be saved without risking damage to it. It stays open with your changes, unsaved; copy its text to keep them.",
	/** Docx editor — action on the save-refused toast: copies the document's text as edited */
	previewDocxSaveRefusedCopy: "Copy text",
	/** Docx editor — toast once a save is stored: edited parts the save check refused were saved as they were last saved, and went back to that in the editor (undo brings the edits back), for one */
	previewDocxSaveLocked_one:
		"{{count}} paragraph or table couldn't be saved with your latest changes, so it was saved as it was before them and is back to that here. Undo brings your changes back.",
	/** Docx editor — as previewDocxSaveLocked_one, for several */
	previewDocxSaveLocked_other:
		"{{count}} paragraphs or tables couldn't be saved with your latest changes, so they were saved as they were before them and are back to that here. Undo brings your changes back.",
	/** Docx editor — as previewDocxSaveLocked when the editor could not take all of them back (some were edited while the file was stored), for one */
	previewDocxSaveLockedKept_one:
		"{{count}} paragraph or table couldn't be saved with your latest changes, so the file has it as it was before those changes. If you changed it while it saved, your changes are still here, unsaved; otherwise undo brings them back.",
	/** Docx editor — as previewDocxSaveLockedKept_one, for several */
	previewDocxSaveLockedKept_other:
		"{{count}} paragraphs or tables couldn't be saved with your latest changes, so the file has them as they were before those changes. Any you changed while it saved keep your changes here, unsaved; undo brings back the others.",
	/** Docx editor — toast while a save looks for the changes the save check refuses (the editor waits meanwhile) */
	previewDocxSaveSearching: "Checking which changes can be saved…",
	/** Docx editor — notice under the toolbar while autosave is held because the save check refused the document; {{shortcut}} is the save shortcut, e.g. ⌘S */
	previewDocxAutosaveHeld: "Changes aren't being saved automatically: this document needs a save by you. Press {{shortcut}} to save.",
	/** Docx editor — action on the save-locked toast: copies the text the user had typed into those parts */
	previewDocxSaveLockedCopy: "Copy my changes",
	/** Docx editor — dialog title asking for the name comments and tracked changes are signed with */
	previewDocxAuthorNameTitle: "Your name in documents",
	/** Docx editor — dialog body asking for the author name; it is synced and can be changed in settings */
	previewDocxAuthorNameBody: "Comments and tracked changes show this name. You can change it later in settings.",
	/** Docx editor — author name field label */
	previewDocxAuthorNameLabel: "Name",
	/** Docx editor — author name dialog submit button */
	previewDocxAuthorNameSave: "Save",

	// ── Markdown ─────────────────────────────────────────────────────────────
	/** Markdown viewer — toolbar button shown while viewing the rendered output; switches to the raw-text source view */
	previewMarkdownViewSourceAction: "View source",
	/** Markdown viewer — toolbar button shown while viewing the raw-text source; switches back to the rendered output */
	previewMarkdownViewRenderedAction: "View rendered",
	/** Markdown viewer — native tooltip on the disabled rendered/source toggle: switching views unmounts the editor buffer, so the toggle is locked until the edit is saved or discarded */
	previewMarkdownToggleDirtyHint: "Save or discard your changes before switching views.",

	// ── Video player ─────────────────────────────────────────────────────────
	/** Video/audio preview — the browser can't decode the file; {{codecs}} lists the formats it reported it can't play, e.g. "HEVC, AC-3" (format names, not translated). Shown with a Download button */
	previewMediaFormatUnsupportedCodecs: "Your browser can't play this file's format ({{codecs}}).",
	/** Video player — notice over a video that plays without sound because the browser can't decode its audio; {{codec}} is the audio format's name, e.g. AC-3 */
	previewMediaNoSound: "No sound: your browser can't play this video's {{codec}} audio.",
	/** Video player — notice over a video that plays its sound without a picture because the browser can't decode its video; {{codec}} is the video format's name, e.g. HEVC */
	previewMediaNoPicture: "No picture: your browser can't play this video's {{codec}} video.",
	/** Video player — accessible label of the button closing the no-sound/no-picture notice */
	previewMediaNoticeDismiss: "Dismiss",
	/** Video player — accessible label of the floating control bar */
	previewMediaControls: "Playback controls",
	/** Video player — heading of the playback-speed menu */
	previewMediaPlaybackSpeed: "Playback speed",
	/** Video player — accessible label of the playback-speed button; {{rate}} is the current speed as shown on the button, e.g. 1.5× */
	previewMediaPlaybackSpeedLabel: "Playback speed {{rate}}",
	/** Video player — one playback speed, e.g. 1.5× */
	previewMediaSpeedRate: "{{rate}}×",
	/** Video player — the 1× entry of the playback-speed menu */
	previewMediaNormalSpeed: "Normal",
	/** Video player — enters picture-in-picture */
	previewMediaPictureInPicture: "Picture in picture",
	/** Video player — leaves picture-in-picture */
	previewMediaExitPictureInPicture: "Exit picture in picture",
	/** Video player — fullscreens the player */
	previewMediaFullscreen: "Full screen",
	/** Video player — leaves full screen */
	previewMediaExitFullscreen: "Exit full screen",
	/** Preview overlay body — a `.ts` file that turned out to be TypeScript, too large to load as text */
	previewTextTooLarge: "This file is too large to show as text.",
	/** Preview overlay body — the browser has no way to stream this video or audio file (no service worker, no MediaSource); the Download button sits below */
	previewMediaCannotStream: "This browser can't stream this file.",
	/** Video player — subtitles menu button and its heading */
	previewMediaSubtitles: "Subtitles",
	/** Video player — subtitles menu: no subtitles */
	previewMediaSubtitlesOff: "Off",
	/** Video player — a subtitle track whose language is not known */
	previewMediaSubtitleUnknownLanguage: "Unknown language",
	/** Video player — marks a forced subtitle track (only foreign-language lines) */
	previewMediaSubtitleForced: "Forced",
	/** Video player — marks a subtitle track for the deaf and hard of hearing */
	previewMediaSubtitleSdh: "SDH",
	/** Video player — marks a subtitle track stored inside the video file, as opposed to a file beside it */
	previewMediaSubtitleEmbedded: "Embedded",
	/** Video player — a chosen subtitle track that could not be read; {{name}} is the track's label */
	previewMediaSubtitleUnreadable: "{{name}} (couldn't be loaded)",
	/** Video player — announced to screen readers when the chosen subtitle file could not be read */
	previewMediaSubtitleUnreadableNotice: "The subtitles couldn't be loaded.",

	// ── Archive browser ──────────────────────────────────────────────────────
	/** Archive browser — accessible label of the entry list; {{name}} is the directory shown (the archive's name at its root) */
	previewArchiveListLabel: "Contents of {{name}}",
	/** Archive browser gate (a tarball or single compressed file over 8 MiB) — listing it downloads all of it; {{size}} is the archive's size */
	previewArchiveGateBody: "Listing its contents reads the whole {{size}} archive.",
	/** Archive browser gate — a listing that would start by itself (a zip, a 7z, a small tarball) was cancelled before it read anything */
	previewArchiveNotListed: "Its contents weren't listed.",
	/** Archive browser gate — starts the listing */
	previewArchiveBrowse: "Browse contents",
	/** Archive browser — extracts the whole archive; the footer's button and the gate's */
	previewArchiveExtractAll: "Extract all",
	/** Archive browser footer — opens the menu of where the selected entries go */
	previewArchiveExtractSelected: "Extract selected",
	/** Archive browser — a public link's owner allows no downloads, so the archive can be browsed but not extracted */
	previewArchiveExtractNotAllowed: "The link's owner doesn't allow downloads, so nothing can be extracted from it.",
	/** Archive browser footer — accessible label of the arrow beside Extract all that opens its other destinations */
	previewArchiveExtractAllMore: "More places to extract to",
	/** Archive browser extract menu — into a new directory next to the archive; {{name}} is that directory's name */
	previewArchiveExtractBesideNewFolder: "Extract to “{{name}}/” next to the archive",
	/** Archive browser extract menu — straight into the directory the archive is in */
	previewArchiveExtractBeside: "Extract next to the archive",
	/** Archive browser — another compress, extract or listing holds the page's one archive slot, so this listing waits */
	previewArchiveWaiting: "Waiting for another archive job",
	/** Archive browser waiting state — the job holding the slot is paused, so it never frees it by itself */
	previewArchiveWaitingPaused: "A paused job is holding it — resume or cancel it in Transfers.",
	/** Archive browser — shown while the archive is being read */
	previewArchiveReading: "Reading the archive…",
	/** Archive browser reading progress — {{read}} and {{total}} are formatted sizes */
	previewArchiveReadingBytes: "{{read}} of {{total}}",
	/** Archive browser reading progress — entries listed so far */
	previewArchiveEntryCount_one: "{{count}} entry",
	/** Archive browser reading progress — entries listed so far */
	previewArchiveEntryCount_other: "{{count}} entries",
	/** Archive browser reading progress — time left; {{eta}} is a clock like 1:05 */
	previewArchiveEta: "{{eta}} left",
	/** Archive browser — stops a running listing; what was read stays browsable */
	previewArchiveStop: "Stop",
	/** Archive browser — the user stopped the listing */
	previewArchiveStopped_one: "Listing stopped after {{count}} entry.",
	/** Archive browser — the user stopped the listing */
	previewArchiveStopped_other: "Listing stopped after {{count}} entries.",
	/** Archive browser — lists the archive again from the start */
	previewArchiveListAgain: "List again",
	/** Archive browser — nothing can be listed without the password (a 7-Zip archive whose file names are encrypted) */
	previewArchiveNeedsPassword: "This archive's contents are encrypted.",
	/** Archive password prompt — title */
	previewArchivePasswordTitle: "Archive password",
	/** Archive password prompt — body; {{name}} is the archive's name */
	previewArchivePasswordBody: "Enter the password for “{{name}}”.",
	/** Archive password prompt — body after a wrong password; {{name}} is the archive's name */
	previewArchivePasswordWrongBody: "That password didn't open “{{name}}”. Try again.",
	/** Archive password prompt — the field's label */
	previewArchivePasswordLabel: "Password",
	/** Archive password prompt — submit button */
	previewArchivePasswordSubmit: "Unlock",
	/** Archive browser — reopens a dismissed password prompt, or opens it from the encrypted-entries banner */
	previewArchiveEnterPassword: "Enter password",
	/** Archive browser banner — some listed entries are encrypted and no password has been given */
	previewArchiveEncryptedBanner: "Some entries are encrypted. Extracting them needs the password.",
	/** Archive browser banner — the password given didn't open the encrypted entries */
	previewArchiveWrongPasswordBanner: "That password didn't open the encrypted entries.",
	/** Archive browser banner — the password is being checked against the encrypted entries */
	previewArchiveCheckingPassword: "Checking the password…",
	/** Archive browser banner — the password check waits until another archive job frees the page's one archive slot */
	previewArchiveCheckingPasswordWaiting: "Waiting for another archive job to check the password",
	/** Archive browser — the listing failed: the format isn't supported */
	previewArchiveFailedUnsupported: "This archive's format isn't supported.",
	/** Archive browser — the listing failed: the archive is damaged */
	previewArchiveFailedCorrupt: "This archive is damaged.",
	/** Archive browser — the listing failed: too many entries to list in the browser, or it unpacks to an implausible size */
	previewArchiveFailedTooLarge: "This archive is too large to list in the browser.",
	/** Archive browser — a damaged archive whose first entries could be read */
	previewArchivePartial_one: "Showing the {{count}} entry read before the damage.",
	/** Archive browser — a damaged archive whose first entries could be read */
	previewArchivePartial_other: "Showing the {{count}} entries read before the damage.",
	/** Archive browser note — entries the listing found but could not hand to the page */
	previewArchiveUndelivered_one: "{{count}} entry couldn't be shown; Extract all still extracts everything.",
	/** Archive browser note — entries the listing found but could not hand to the page */
	previewArchiveUndelivered_other: "{{count}} entries couldn't be shown; Extract all still extracts everything.",
	/** Archive browser note — a zip names some entries more than once; only the last of each name is listed and extracted. {{count}} is how many were left out; the names follow in a list */
	previewArchiveDuplicates_one: "{{count}} entry is left out because a later entry has the same name.",
	/** Archive browser note — a zip names some entries more than once; only the last of each name is listed and extracted. {{count}} is how many were left out; the names follow in a list */
	previewArchiveDuplicates_other: "{{count}} entries are left out because a later entry has the same name.",
	/** Archive browser — the whole archive holds nothing */
	previewArchiveEmpty: "This archive is empty.",
	/** Archive browser — the directory shown holds nothing */
	previewArchiveDirectoryEmpty: "This directory is empty.",
	/** Archive browser search — nothing below the directory shown matches; {{query}} is the search text */
	previewArchiveNoResults: "Nothing here matches “{{query}}”.",
	/** Archive browser search — only the first {{count}} matches are shown */
	previewArchiveSearchTruncated: "Showing the first {{count}} matches — refine your search.",
	/** Archive browser search field — placeholder and accessible label */
	previewArchiveSearchPlaceholder: "Search this directory",
	/** Archive browser — accessible label of the breadcrumbs */
	previewArchiveBreadcrumbs: "Location in the archive",
	/** Archive browser breadcrumbs — accessible label of the "…" menu holding the levels not shown */
	previewArchiveHiddenLevels: "Show the directories above",
	/** Archive browser column header — sorts by name */
	previewArchiveColumnName: "Name",
	/** Archive browser column header — sorts by size */
	previewArchiveColumnSize: "Size",
	/** Archive browser column header — sorts by date modified */
	previewArchiveColumnModified: "Modified",
	/** Archive browser — the header checkbox selecting everything in the directory shown, or every search match */
	previewArchiveSelectAll: "Select all",
	/** Archive browser row — a directory's size column: how many entries it directly holds */
	previewArchiveDirItems_one: "{{count}} item",
	/** Archive browser row — a directory's size column: how many entries it directly holds */
	previewArchiveDirItems_other: "{{count}} items",
	/** Archive browser footer — files selected and their size; {{size}} is formatted */
	previewArchiveSelectedFiles_one: "{{count}} file selected · {{size}}",
	/** Archive browser footer — files selected and their size; {{size}} is formatted */
	previewArchiveSelectedFiles_other: "{{count}} files selected · {{size}}",
	/** Archive browser footer — only directories selected (no files in them) */
	previewArchiveSelectedItems_one: "{{count}} item selected",
	/** Archive browser footer — only directories selected (no files in them) */
	previewArchiveSelectedItems_other: "{{count}} items selected",
	/** Archive browser row — an encrypted entry */
	previewArchiveEncrypted: "Encrypted",
	/** Archive browser row — the name holds invisible or direction-changing characters */
	previewArchiveMisleading: "This name has invisible or direction-changing characters; it may not be what it looks like.",
	/** Archive browser row — the name was changed to be usable in the drive; {{path}} is the name as the archive stores it */
	previewArchiveStoredAs: "Stored as {{path}}",
	/** Archive browser row — a link's target; {{target}} is the path it points to */
	previewArchiveLinkTarget: "Points to {{target}}",
	/** Archive browser row — a symbolic link, which the drive cannot hold */
	previewArchiveSkip_symlink: "Symbolic link · not extracted",
	/** Archive browser row — a tar hard link whose file is not in the archive */
	previewArchiveSkip_hardlink: "Hard link to a missing file · not extracted",
	/** Archive browser row — a device node or pipe */
	previewArchiveSkip_device: "Device or pipe · not extracted",
	/** Archive browser row — a sparse file */
	previewArchiveSkip_sparse: "Sparse file · not extracted",
	/** Archive browser row — an entry of a kind that is not extracted */
	previewArchiveSkip_unsupportedType: "Unsupported entry · not extracted",
	/** Archive browser row — a path longer than the drive allows */
	previewArchiveSkip_pathTooLong: "Path too long · not extracted",
	/** Archive browser row — a path nested too deeply */
	previewArchiveSkip_pathTooDeep: "Nested too deeply · not extracted",
	/** Archive browser row — a path leading out of the extract's directory, or not usable as a drive name */
	previewArchiveSkip_unsafePath: "Unsafe path · not extracted",
	/** Archive browser row — a ZIP entry whose data overlaps another's (a damaged or crafted archive) */
	previewArchiveSkip_overlappingData: "Damaged entry · not extracted",
	/** Archive browser row — compressed or encrypted in a way that cannot be read */
	previewArchiveSkip_unsupportedMethod: "Unsupported compression or encryption · not extracted",
	/** Archive browser row — a 7-Zip deletion marker */
	previewArchiveSkip_antiItem: "Deletion marker · not extracted",
	/** Archive browser row — macOS metadata, left out like an extract leaves it out */
	previewArchiveSkip_macMetadata: "macOS metadata · not extracted",
	/** Hard-link dialog — title: selected hard links point to files outside the directory shown */
	previewArchiveLinkTargetsTitle: "Linked files outside this directory",
	/** Hard-link dialog — body; {{base}} is the directory shown */
	previewArchiveLinkTargetsBody_one:
		"{{count}} selected hard link points to a file outside “{{base}}”, which an extract from here can't include.",
	/** Hard-link dialog — body; {{base}} is the directory shown */
	previewArchiveLinkTargetsBody_other:
		"{{count}} selected hard links point to files outside “{{base}}”, which an extract from here can't include.",
	/** Hard-link dialog — extracts the selection without those links */
	previewArchiveLinkTargetsLeaveOut: "Leave those links out",
	/** Hard-link dialog — extracts with paths taken from a directory higher up that holds the link targets too; {{name}} is that directory */
	previewArchiveLinkTargetsFromParent: "Extract from “{{name}}” instead",
	/** Archive entry viewer — leaves the file shown for the archive's contents again (Escape does too) */
	previewArchiveEntryBack: "Back to the archive",
	/** Solid 7z cost dialog title — opening or saving one file means reading the files stored before it in the same block first */
	previewArchiveEntryCostTitle: "Read part of the archive first?",
	/** Solid 7z cost dialog body, opening; {{name}} is the file, {{download}} roughly how much of the archive is downloaded, {{skipped}} how much of the files stored before it is decoded and thrown away */
	previewArchiveEntryCostOpenBody:
		"This archive stores “{{name}}” together with other files. To open it, about {{download}} of the archive is downloaded, and {{skipped}} of the files stored before it is unpacked and discarded first.",
	/** Solid 7z cost dialog body, saving; same placeholders as the opening one */
	previewArchiveEntryCostDownloadBody:
		"This archive stores “{{name}}” together with other files. To save it, about {{download}} of the archive is downloaded, and {{skipped}} of the files stored before it is unpacked and discarded first.",
	/** Solid 7z cost dialog — opens the file anyway */
	previewArchiveEntryCostOpen: "Open",
	/** Solid 7z cost dialog — saves the file anyway */
	previewArchiveEntryCostDownload: "Download",
	/** Toast when a browser can't stream one file of an archive to disk (no File System Access and no transferable streams) */
	previewArchiveEntryDownloadUnsupported: "This browser can't save a single file from an archive. Extract it to your drive instead."
} as const
