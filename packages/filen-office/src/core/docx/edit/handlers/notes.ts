import { docxNotes } from "../../stories/notes"
import { bindEditor } from "../../stories/headers"
import { noteOptions } from "../../stories/sections"
import type { DocumentCommandContext, DocumentHandlers, FeatureHandlers } from "../registry"

// Footnotes and endnotes: insert, step between a note and its reference, delete, convert, numbering options
// (document commands: a note is a story of its own and a reference in the body). Enablement is decided per
// snapshot, so it looks only at the caret (and the editor's model), never through the document.

function notes(context: DocumentCommandContext) {
	bindEditor(context.context, context.editor)

	return docxNotes(context.editor)
}

const writable = (context: DocumentCommandContext): boolean => !context.editor.body.session.readOnly

// The caret is in a note, or next to a reference (where the store finds a note).
const onNote = (context: DocumentCommandContext): boolean => {
	const kind = context.editor.session.story.story.kind

	return kind === "footnote" || kind === "endnote" || notes(context).at() !== undefined
}

// Whether the document has notes of `kind` (normal ones: Word's separators are not notes).
function hasNotes(context: DocumentCommandContext, kind: "footnote" | "endnote"): boolean {
	const map = kind === "footnote" ? context.editor.edits.model.footnotes : context.editor.edits.model.endnotes

	for (const story of map.values()) if (story.noteType === "normal") return true

	return false
}

const document: DocumentHandlers = {
	"note.insertFootnote": {
		enabled: context => notes(context).canInsert(),
		run: (context, command) => notes(context).insert("footnote", command.mark)
	},
	"note.insertEndnote": {
		enabled: context => notes(context).canInsert(),
		run: (context, command) => notes(context).insert("endnote", command.mark)
	},
	"note.goto": {
		enabled: onNote,
		run: context => notes(context).goto()
	},
	"note.delete": {
		enabled: context => writable(context) && onNote(context),
		run: context => notes(context).remove()
	},
	"note.convert": {
		enabled: context => writable(context) && onNote(context),
		run: context => notes(context).convert()
	},
	"note.convertAll": {
		enabled: context => writable(context) && (hasNotes(context, "footnote") || hasNotes(context, "endnote")),
		run: (context, command) => notes(context).convertAll(command.to)
	},
	"note.options": {
		enabled: writable,
		run: (context, command) => {
			bindEditor(context.context, context.editor)

			return noteOptions(context.editor, command)
		}
	}
}

export const NOTE_HANDLERS: FeatureHandlers = { document }
