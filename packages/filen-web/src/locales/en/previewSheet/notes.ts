// "preview" namespace, spreadsheet viewer — notes, threaded comments and hyperlinks. Spread into
// preview.ts; same catalog rules.
export const previewSheetNotes = {
	/** Spreadsheet grid menu — adds a note (a plain annotation) to the active cell */
	previewSheetNoteInsert: "Insert note",
	/** Spreadsheet grid menu — opens the active cell's note to change its text */
	previewSheetNoteEdit: "Edit note",
	/** Spreadsheet grid menu — removes the notes of the selected cells */
	previewSheetNoteDelete: "Delete note",
	/** Spreadsheet grid menu — keeps the active cell's note shown all the time */
	previewSheetNoteShow: "Show note",
	/** Spreadsheet grid menu — shows the active cell's note only while the pointer is on the cell again */
	previewSheetNoteHide: "Hide note",
	/** Note card — the field holding the note's text */
	previewSheetNoteText: "Note",
	/** Note card — placeholder of an empty note's field */
	previewSheetNotePlaceholder: "Type a note",
	/** Note card — saves the note's text */
	previewSheetNoteSave: "Save",
	/** Note and comment cards — closes the field without saving */
	previewSheetNoteCancel: "Cancel",
	/** Note card (read-only) — accessible name, with the note's author */
	previewSheetNoteBy: "Note by {{author}}",
	/** Spreadsheet grid menu — starts a comment thread on the active cell */
	previewSheetCommentNew: "New comment",
	/** Spreadsheet grid menu — opens the active cell's comment thread */
	previewSheetCommentShow: "Show comment",
	/** Spreadsheet grid menu — deletes the active cell's comment thread with its replies */
	previewSheetCommentDeleteThread: "Delete comment",
	/** Comment card — accessible name, with the first comment's author */
	previewSheetCommentThread: "Comment thread by {{author}}",
	/** Comment card — placeholder of the field starting a thread */
	previewSheetCommentPlaceholder: "Start a conversation",
	/** Comment card — posts a new comment */
	previewSheetCommentPost: "Post",
	/** Comment card — placeholder of the reply field, and the button opening it */
	previewSheetCommentReplyPlaceholder: "Reply",
	/** Comment card — posts a reply */
	previewSheetCommentReply: "Reply",
	/** Comment card — saves an edited comment */
	previewSheetCommentSave: "Save",
	/** Comment card — the ⋯ menu of one comment */
	previewSheetCommentActions: "Comment actions",
	/** Comment card ⋯ menu — changes the text of one's own comment */
	previewSheetCommentEdit: "Edit",
	/** Comment card ⋯ menu — removes one reply */
	previewSheetCommentDelete: "Delete",
	/** Comment card — beside the first comment of a thread marked resolved in Excel */
	previewSheetCommentResolved: "Resolved",
	/** Note and comment cards — the author of a comment whose author the file does not name */
	previewSheetUnknownAuthor: "Unknown author",
	/** Spreadsheet grid menu and Ctrl+K — opens the dialog that links the selected cells */
	previewSheetLinkInsert: "Link…",
	/** Spreadsheet grid menu — opens the dialog that changes the active cell's link */
	previewSheetLinkEdit: "Edit link…",
	/** Spreadsheet grid menu and link card — opens the active cell's link */
	previewSheetLinkOpen: "Open link",
	/** Spreadsheet grid menu and link card — removes the links of the selected cells */
	previewSheetLinkRemove: "Remove link",
	/** Link card — how to open a link from the grid (Windows and Linux) */
	previewSheetLinkHint: "Ctrl+click to open",
	/** Link card — how to open a link from the grid (macOS) */
	previewSheetLinkHintMac: "⌘+click to open",
	/** Link card — a link to cells of this workbook, with the place */
	previewSheetLinkPlace: "Go to {{place}}",
	/** Link card and toast — a link whose target is not a web, mail or phone address Filen opens */
	previewSheetLinkBlocked: "This link can't be opened here",
	/** Link card — a HYPERLINK formula whose target is only known once it is calculated */
	previewSheetLinkCalculated: "The formula works out where this link goes",
	/** Link card — a link to a place this workbook does not have */
	previewSheetLinkNowhere: "This link points to a place the workbook doesn't have",
	/** Link dialog — title when the cells have no link yet */
	previewSheetLinkInsertTitle: "Insert link",
	/** Link dialog — title when the active cell has a link */
	previewSheetLinkEditTitle: "Edit link",
	/** Link dialog — description */
	previewSheetLinkBody: "A link opens with a click while Ctrl (⌘ on a Mac) is held.",
	/** Link dialog — the choice of what the link goes to */
	previewSheetLinkKind: "Link to",
	/** Link dialog — a web, email or phone address */
	previewSheetLinkKindWeb: "Web address",
	/** Link dialog — cells or a name in this workbook */
	previewSheetLinkKindPlace: "Place in this workbook",
	/** Link dialog — the address field */
	previewSheetLinkAddress: "Address",
	/** Link dialog — the address field's error */
	previewSheetLinkAddressInvalid: "Enter a web address (https://…), an email address (mailto:…) or a phone number (tel:…).",
	/** Link dialog — the sheet the link goes to */
	previewSheetLinkSheet: "Sheet",
	/** Link dialog — the cell reference or defined name the link goes to */
	previewSheetLinkReference: "Cell or name",
	/** Link dialog — the reference field's error */
	previewSheetLinkReferenceInvalid: "Enter a cell reference such as A1 or B2:C5, or a defined name.",
	/** Link dialog — what the cell shows (Excel's "Text to display") */
	previewSheetLinkText: "Text to display",
	/** Link dialog — the tooltip shown with the link (Excel's "ScreenTip") */
	previewSheetLinkTooltip: "ScreenTip",
	/** Link dialog — applies the link */
	previewSheetLinkApply: "Apply",
	/** Toast — the sheet's notes or comments could not be read, so they cannot be changed */
	previewSheetNotesUnavailable: "This sheet's notes and comments can't be changed: part of the file holding them is damaged",
	/** Toast — the author name is needed to write a note or comment */
	previewSheetAuthorNeeded: "Set an author name to add notes and comments",
	/** Toast — someone else's comment cannot be edited */
	previewSheetNotOwnComment: "Only its author can edit a comment",
	/** Toast — a link target that cannot be used */
	previewSheetLinkTargetInvalid: "This link target can't be used"
}
