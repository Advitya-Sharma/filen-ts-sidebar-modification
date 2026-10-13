import { TextSelection } from "prosemirror-state"
import { freshSlice, markInternal } from "../../../core/docx/edit/invariants/index"
import { sourceCopies } from "../../../core/docx/edit/session"
import { docxNotes, noteReferences } from "../../../core/docx/stories/notes"
import { docxSchema, type DocxEditorCommand, type DocxEditorSession, type PagedEditor } from "../harness"
import { textPositions } from "./base"
import type { EditOp, OpContext, Rand } from "./index"

// Footnotes and endnotes under the paged editor (whose pages must stay those of the saved file): inserting
// them (numbered, with a custom mark, while tracking), typing in them, stepping between note and reference,
// deleting, converting one or all, numbering options; references made raw for the writer and Gate 2; and the
// deletion of a source note's reference in the scripted suite (its note leaves its part).

const nodes = docxSchema.nodes
const AUTHOR = "Fuzz Author"
const MARKS = ["*", "†", "a)", "§§"]
const WORDS = ["note", "Filen", "see p. 4", "Kraków", "x"]

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

// A caret somewhere in the body's text (the body entered first).
function bodyCaret(editor: PagedEditor, rand: Rand): number | undefined {
	if (editor.session !== editor.body.session) editor.exitStory()

	const session = editor.session
	const positions = textPositions(session.state.doc)

	if (positions.length === 0) return undefined

	const at = pick(rand, positions)

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at)))

	return at
}

// A caret next to a random note reference of the body.
function referenceCaret(editor: PagedEditor, rand: Rand): number | undefined {
	if (editor.session !== editor.body.session) editor.exitStory()

	const session = editor.session
	const references = noteReferences(session.state.doc)

	if (references.length === 0) return undefined

	const at = pick(rand, references).pos

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, at)))

	return at
}

function noteOp(name: string, weight: number, run: (context: OpContext, editor: PagedEditor) => string): EditOp {
	return {
		name,
		harness: "paged",
		weight,
		run(context) {
			const editor = context.editor

			return editor === undefined ? "noop" : run(context, editor)
		}
	}
}

function execute(context: OpContext, command: DocxEditorCommand): string {
	return context.execute(command.type, command)
}

// Runs `run` with changes tracked by the fuzz's author, as before after it.
function tracked(session: DocxEditorSession, run: () => string): string {
	const review = session.context.review
	const identity = session.context.identity
	const was = review.tracking
	const author = identity.author

	review.tracking = true
	identity.author ??= AUTHOR

	try {
		return run()
	} finally {
		review.tracking = was
		identity.author = author
	}
}

function insert(context: OpContext, editor: PagedEditor): string {
	const rand = context.rand
	const at = bodyCaret(editor, rand)

	if (at === undefined) return "noop"

	const kind = rand() < 0.7 ? "note.insertFootnote" : "note.insertEndnote"
	const command: DocxEditorCommand = rand() < 0.2 ? { type: kind, mark: pick(rand, MARKS) } : { type: kind }
	const run = (): string => {
		const done = execute(context, command)

		// The caret is in the new note: what is typed there.
		if (done === "true" && rand() < 0.8) editor.session.dispatch(editor.session.state.tr.insertText(pick(rand, WORDS)))

		return done
	}

	return `${kind}${"mark" in command ? ` ${command.mark}` : ""}@${at} -> ${rand() < 0.2 ? tracked(editor.body.session, run) : run()}`
}

function atReference(name: string, command: (rand: Rand) => DocxEditorCommand): EditOp {
	return noteOp(name, 1, (context, editor) => {
		const at = referenceCaret(editor, context.rand)

		if (at === undefined) return "noop"

		const chosen = command(context.rand)

		return `${chosen.type}@${at} -> ${execute(context, chosen)}`
	})
}

// ---- damage: references made raw -----

function referenceDamage(context: OpContext): string {
	const rand = context.rand
	const session = context.session
	const positions = textPositions(session.state.doc)

	if (positions.length === 0) return "noop"

	const at = pick(rand, positions)
	const kind = rand() < 0.5 ? "footnote" : "endnote"
	const known = [...(kind === "footnote" ? session.context.model.footnotes.keys() : session.context.model.endnotes.keys())]
	const id = pick(rand, [...known, "0", "-1", "9999", String(Math.floor(rand() * 5))])
	const node = nodes.note_ref.create({ kind, id })

	return `noteRef ${kind}:${id}@${at} -> ${context.transact("noteRef", tr => markInternal(tr.insert(at, node)))}`
}

function deleteReference(session: DocxEditorSession, rand: Rand): boolean {
	const references = noteReferences(session.state.doc).filter(reference => (reference.node.attrs["src"] as number) >= 0)

	if (references.length === 0) return false

	const reference = pick(rand, references)

	return session.dispatch(session.state.tr.delete(reference.pos, reference.pos + 1))
}

export const noteOps: readonly EditOp[] = [
	noteOp("noteInsert", 4, insert),
	atReference("noteGoto", () => ({ type: "note.goto" })),
	atReference("noteDelete", () => ({ type: "note.delete" })),
	atReference("noteConvert", () => ({ type: "note.convert" })),
	noteOp("noteConvertAll", 1, (context, editor) => {
		const to = context.rand() < 0.5 ? "footnote" : "endnote"

		if (editor.session !== editor.body.session) editor.exitStory()

		return `note.convertAll ${to} -> ${execute(context, { type: "note.convertAll", to })}`
	}),
	noteOp("noteOptions", 1, (context, editor) => {
		const rand = context.rand

		if (bodyCaret(editor, rand) === undefined) return "noop"

		const kind = rand() < 0.6 ? "footnote" : "endnote"
		const command: DocxEditorCommand = {
			type: "note.options",
			kind,
			scope: rand() < 0.5 ? "section" : "document",
			format: pick(rand, ["decimal", "lowerRoman", "upperLetter", "chicago"] as const),
			...(rand() < 0.5 ? { start: 1 + Math.floor(rand() * 9) } : {}),
			...(kind === "footnote" && rand() < 0.3 ? { restart: pick(rand, ["continuous", "eachSect", "eachPage"] as const) } : {}),
			...(rand() < 0.3
				? {
						position:
							kind === "footnote"
								? pick(rand, ["pageBottom", "beneathText"] as const)
								: pick(rand, ["sectEnd", "docEnd"] as const)
					}
				: {})
		}

		return `note.options ${JSON.stringify(command)} -> ${execute(context, command)}`
	}),
	noteOp("noteReferenceDelete", 1, (context, editor) => {
		if (editor.session !== editor.body.session) editor.exitStory()

		return `noteReferenceDelete -> ${String(deleteReference(editor.session, context.rand))}`
	}),
	// A paragraph's text with a note reference copied (or cut) and pasted into the body as the clipboard does:
	// the pasted reference refers to a copy of its note (a cut keeps it), in the paste's one step.
	noteOp("notePaste", 1, (context, editor) => {
		const rand = context.rand

		if (referenceCaret(editor, rand) === undefined) return "noop"

		const session = editor.session
		const $at = session.state.selection.$from
		const from = $at.start()
		const to = $at.end()
		const slice = freshSlice(
			session.state.doc.slice(from, to),
			sourceCopies(session.context),
			session.story.part.index,
			session.story.part.index,
			true
		)
		const cut = rand() < 0.3

		if (cut && !session.dispatch(session.state.tr.delete(from, to).setMeta("uiEvent", "cut"))) return "cut refused"
		if (bodyCaret(editor, rand) === undefined) return "noop"

		const tr = session.state.tr.replaceSelection(slice).setMeta("uiEvent", "paste")

		// Refused, as the clipboard refuses it, where a note cannot be copied whole.
		if (!docxNotes(editor).pasted(tr)) return `notePaste ${cut ? "cut" : "copy"} ${from}-${to} -> refused`

		return `notePaste ${cut ? "cut" : "copy"} ${from}-${to} -> ${String(session.dispatch(tr))}`
	}),
	{ name: "noteRefDamage", harness: "damage", run: referenceDamage },
	{
		name: "noteDeleteScripted",
		harness: "damage",
		weight: 0,
		run: () => "noop",
		writes: ["footnotes", "endnotes"],
		script: {
			order: 1000,
			// The first source note reference of the body goes, and its note with it.
			run: target => deleteReference(target.session, () => 0)
		}
	}
]
