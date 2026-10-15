import {
	DOCX_EDITOR_COMMANDS,
	isWriteCommand,
	type DocxEditorCommand,
	type DocxEditorCommandType,
	type DocxReviewSnapshot
} from "@filen/office/core/docx/edit"

// Commands that can record an author: the host has the user's name before running one that will.
const AUTHOR_COMMANDS: ReadonlySet<string> = new Set(DOCX_EDITOR_COMMANDS.filter(info => info.author === true).map(info => info.type))

// Writes that are never tracked: undo and redo, resolving changes, turning tracking on or off, and what changes
// a comment's state (as in Word, no change is recorded for them).
const UNTRACKED: ReadonlySet<DocxEditorCommandType> = new Set<DocxEditorCommandType>([
	"history.undo",
	"history.redo",
	"review.track",
	"review.accept",
	"review.reject",
	"review.acceptAll",
	"review.rejectAll",
	"comment.resolve",
	"comment.reopen",
	"comment.delete"
])

// Whether running `command` records the author: turning tracking off never does (it must stay possible without
// a name).
export function commandRecordsAuthor(command: DocxEditorCommand, tracking: boolean): boolean {
	if (!AUTHOR_COMMANDS.has(command.type)) {
		return false
	}

	return command.type === "review.track" ? (command.on ?? !tracking) : true
}

// Whether the host asks for the author's name before running `command` in the document's review state: a
// command that records the author, and while tracking without a name (the session would refuse it) every write
// that would be tracked.
export function commandNeedsAuthor(command: DocxEditorCommand, review: DocxReviewSnapshot): boolean {
	if (commandRecordsAuthor(command, review.tracking)) {
		return true
	}

	return review.authorNeeded && isWriteCommand(command.type) && !UNTRACKED.has(command.type)
}

// Runs `run` with the user's author name, asking for it first when none is set; never when the user declines.
export type DocxAuthorGate = (run: (author: string) => void) => void
