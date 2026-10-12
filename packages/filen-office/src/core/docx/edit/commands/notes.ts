import type { DocxCommandInfo, DocxCommandParser } from "./base"

// Footnotes and endnotes.

export type NoteCommand =
	| { readonly type: "note.insertFootnote" }
	| { readonly type: "note.insertEndnote" }
	// Between a note reference and its note.
	| { readonly type: "note.goto" }

const document = { scope: "document", group: "note", context: "note" } as const

export const NOTE_COMMANDS: readonly DocxCommandInfo<NoteCommand["type"]>[] = [
	{ type: "note.insertFootnote", labelKey: "docx.note.insertFootnote", icon: "footnote", write: true, ...document },
	{ type: "note.insertEndnote", labelKey: "docx.note.insertEndnote", icon: "endnote", write: true, ...document },
	{ type: "note.goto", labelKey: "docx.note.goto", icon: "note-goto", write: false, ...document }
]

export const parseNoteCommand: DocxCommandParser<NoteCommand> = type => ({ type })
