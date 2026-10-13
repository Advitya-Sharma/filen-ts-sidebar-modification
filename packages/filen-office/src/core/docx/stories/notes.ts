import { Fragment, Slice, type Mark, type Node } from "prosemirror-model"
import { NodeSelection, TextSelection, type EditorState, type Transaction } from "prosemirror-state"
import { MAX_NOTE_MARK, type DocxNoteKind } from "../edit/commands/notes"
import type { DocxEditContext } from "../edit/context"
import { builtInStyleId } from "../edit/formats"
import { freshSlice, markCommand, markInternal } from "../edit/invariants/index"
import { NO_SOURCE, WRAPPER_MARKS, docxSchema } from "../edit/schema"
import { sourceCopies, type DocxEditorSession } from "../edit/session"
import type { Story } from "../model/index"
import type { PagedEditor } from "../paged/editor"
import { revisionStamp } from "../review/track"
import { hasStyleDefinition } from "../write/styles"

// Footnotes and endnotes: a note is a story of its own (Word's w:footnote / w:endnote) and a reference to it in
// the body (a note_ref). Inserting one adds the story through the paged editor (it is saved only while a
// reference to it exists) and the reference at the caret; deleting the reference deletes the note, as Word does
// (the save leaves out a note nothing refers to). Converting moves a note's content into a new note of the
// other kind. Notes live in the main document's text only, as in Word (not in headers, footers, comments or
// other notes).

const nodes = docxSchema.nodes
const marks = docxSchema.marks
const WRAPPERS: ReadonlySet<unknown> = new Set(WRAPPER_MARKS.map(name => marks[name]))
const MARK_KINDS: Readonly<Record<DocxNoteKind, string>> = { footnote: "FootnoteMark", endnote: "EndnoteMark" }

// A note as the caret finds it: its kind and id, its story, and where its reference is in the body.
export interface DocxNoteAt {
	readonly kind: DocxNoteKind
	readonly id: string
	readonly story: number
	// The reference's position in the body, -1 when there is none.
	readonly reference: number
}

function otherKind(kind: DocxNoteKind): DocxNoteKind {
	return kind === "footnote" ? "endnote" : "footnote"
}

// A built-in style's id where the document has it or the writer can define it, else undefined.
function styleOf(context: DocxEditContext, name: string, type: "paragraph" | "character"): string | undefined {
	const existing = context.styles.byName(name, type)

	if (existing !== undefined) return existing.id

	return hasStyleDefinition(name) ? builtInStyleId(context, name) : undefined
}

// The marks of a note's number (its reference, and its mark in the note): the reference style, else (a
// document that cannot get it) superscript as the style makes it.
function referenceMarks(context: DocxEditContext, kind: DocxNoteKind): readonly Mark[] {
	const style = styleOf(context, `${kind} reference`, "character")

	return style === undefined ? [marks.vert_align.create({ value: "superscript" })] : [marks.char_style.create({ styleId: style })]
}

function textStyle(context: DocxEditContext, kind: DocxNoteKind): string | null {
	return styleOf(context, `${kind} text`, "paragraph") ?? null
}

// The note references of a story document, with their positions, in order.
export function noteReferences(doc: Node): { readonly node: Node; readonly pos: number }[] {
	const out: { node: Node; pos: number }[] = []

	doc.descendants((node, pos) => {
		if (node.type === nodes.note_ref) out.push({ node, pos })

		return !node.isInline
	})

	return out
}

function referenceOf(doc: Node, kind: DocxNoteKind, id: string): number {
	for (const reference of noteReferences(doc))
		if (reference.node.attrs["kind"] === kind && reference.node.attrs["id"] === id) return reference.pos

	return -1
}

// The reference at a selection: the one selected, or the one right after or before its start (enablement asks
// with every snapshot: nothing walks the selection).
function referenceAt(state: EditorState): { readonly node: Node; readonly pos: number } | undefined {
	const selection = state.selection

	if (selection instanceof NodeSelection && selection.node.type === nodes.note_ref) return { node: selection.node, pos: selection.from }

	const { $from } = selection
	const after = $from.nodeAfter
	const before = $from.nodeBefore

	if (after?.type === nodes.note_ref) return { node: after, pos: $from.pos }
	if (before?.type === nodes.note_ref && selection.empty) return { node: before, pos: $from.pos - before.nodeSize }

	return undefined
}

// The size of the custom mark following a reference (the text after it in the reference's style), or 0.
function customMarkSize(doc: Node, pos: number): number {
	const reference = doc.nodeAt(pos)

	if (reference?.type !== nodes.note_ref || reference.attrs["custom"] !== true) return 0

	const after = doc.resolve(pos + 1).nodeAfter
	const style = reference.marks.find(mark => mark.type === marks.char_style || mark.type === marks.vert_align)

	return after?.isText === true && style !== undefined && style.isInSet(after.marks) !== undefined
		? Math.min(after.nodeSize, MAX_NOTE_MARK)
		: 0
}

const STORES = new WeakMap<PagedEditor, DocxNotes>()

// The notes of the document `editor` edits.
export function docxNotes(editor: PagedEditor): DocxNotes {
	let store = STORES.get(editor)

	if (store === undefined) {
		store = new DocxNotes(editor)
		STORES.set(editor, store)
	}

	return store
}

export class DocxNotes {
	private readonly editor: PagedEditor

	constructor(editor: PagedEditor) {
		this.editor = editor
	}

	private get body(): DocxEditorSession {
		return this.editor.body.session
	}

	private get context(): DocxEditContext {
		return this.body.context
	}

	// The story of note `id` of `kind` (the document's or one the editor added), or undefined.
	storyOf(kind: DocxNoteKind, id: string): Story | undefined {
		const map = kind === "footnote" ? this.editor.edits.model.footnotes : this.editor.edits.model.endnotes

		return map.get(id)
	}

	// The note at the caret: the one being edited, or the one whose reference the caret is at.
	at(): DocxNoteAt | undefined {
		const session = this.editor.session
		const story = session.story.story

		if ((story.kind === "footnote" || story.kind === "endnote") && story.id !== undefined && story.noteType === "normal")
			return { kind: story.kind, id: story.id, story: story.index, reference: referenceOf(this.body.state.doc, story.kind, story.id) }

		if (session !== this.body) return undefined

		const found = referenceAt(session.state)
		const kind = found?.node.attrs["kind"] === "endnote" ? "endnote" : "footnote"
		const id = found?.node.attrs["id"]
		const note = typeof id === "string" ? this.storyOf(kind, id) : undefined

		return found === undefined || note === undefined || typeof id !== "string"
			? undefined
			: { kind, id, story: note.index, reference: found.pos }
	}

	// Whether a note can be inserted at the caret: in the body's text, where it can be edited.
	canInsert(): boolean {
		const session = this.editor.session

		if (session !== this.body || session.readOnly) return false

		const $to = session.state.selection.$to

		return $to.parent.type === nodes.paragraph && (!this.context.review.tracking || revisionStamp(this.context.identity) !== undefined)
	}

	// A new note of `kind` at the caret (after the selection), numbered or with `mark`; the caret goes into it.
	insert(kind: DocxNoteKind, mark?: string): boolean {
		if (!this.canInsert()) return false
		// A custom mark needs w:customMarkFollows on the reference, which the schema must hold.
		if (mark !== undefined && !Object.hasOwn(nodes.note_ref.spec.attrs ?? {}, "custom")) return false

		const session = this.body
		const context = this.context
		const at = session.state.selection.to
		const made = this.newNote(kind, mark, undefined)

		if (made === undefined) return false

		const tr = session.state.tr
		const reference = this.referenceNode(kind, made.id, mark !== undefined, tr.doc, at)
		// A custom mark is the text right after the reference (w:customMarkFollows), in its style.
		const content =
			mark === undefined
				? Fragment.from(reference)
				: Fragment.from(reference).append(Fragment.from(docxSchema.text(mark, reference.marks)))

		tr.insert(at, content)
		tr.setSelection(TextSelection.create(tr.doc, at + content.size))

		// A command transaction: tracked changes record the reference as inserted.
		if (!session.dispatch(markCommand(this.tracked(tr, at, at + content.size, context)))) return false

		return this.enter(made.story)
	}

	// Moves between a note and its reference: into the note of the reference at the caret, or back to the
	// reference of the note being edited (the caret after it).
	goto(): boolean {
		const note = this.at()

		if (note === undefined) return false

		const editor = this.editor

		if (editor.session !== this.body) {
			if (note.reference < 0) return false

			editor.exitStory()

			const doc = this.body.state.doc

			return this.body.dispatch(this.body.state.tr.setSelection(TextSelection.create(doc, note.reference + 1)).scrollIntoView())
		}

		return this.enter(note.story)
	}

	// Deletes the note at the caret by deleting its reference (tracked as a deletion while tracking is on).
	remove(): boolean {
		const note = this.at()

		if (note === undefined || note.reference < 0 || this.body.readOnly) return false
		if (this.editor.session !== this.body) this.editor.exitStory()

		const session = this.body
		const doc = session.state.doc

		return session.dispatch(session.state.tr.delete(note.reference, note.reference + 1 + customMarkSize(doc, note.reference)))
	}

	// Whether the note at the caret can become one of the other kind.
	canConvert(): boolean {
		const note = this.at()

		return note !== undefined && note.reference >= 0 && !this.body.readOnly
	}

	// The note at the caret becomes one of the other kind, its content moved (as a paste moves it).
	convert(): boolean {
		const note = this.at()

		if (note === undefined || note.reference < 0 || this.body.readOnly) return false
		if (this.editor.session !== this.body) this.editor.exitStory()

		return this.convertNotes([{ kind: note.kind, id: note.id, pos: note.reference }])
	}

	// Every note of the other kind becomes one of kind `to`.
	convertAll(to: DocxNoteKind): boolean {
		if (this.body.readOnly) return false

		const from = otherKind(to)
		const found = noteReferences(this.body.state.doc)
			.filter(reference => reference.node.attrs["kind"] === from && typeof reference.node.attrs["id"] === "string")
			.map(reference => ({ kind: from, id: reference.node.attrs["id"] as string, pos: reference.pos }))

		if (found.length === 0) return false
		if (this.editor.session !== this.body) this.editor.exitStory()

		return this.convertNotes(found)
	}

	// Whether the document has notes of `kind` (references in the body).
	any(kind?: DocxNoteKind): boolean {
		return noteReferences(this.body.state.doc).some(reference => kind === undefined || reference.node.attrs["kind"] === kind)
	}

	// ---- internals -----------------------------------------------------------------------------------------

	private enter(story: number): boolean {
		const editor = this.editor

		if (!editor.enterStory(story, -1)) return false

		const session = editor.session
		const doc = session.state.doc
		const first = doc.firstChild
		// After the note's mark and the space after it: where Word puts the caret in a new note.
		const end = first === null ? 1 : first.nodeSize - 1

		return session.dispatch(session.state.tr.setSelection(TextSelection.create(doc, Math.max(1, end))).scrollIntoView())
	}

	// Insertion marks for `from`..`to` of `tr` while tracking.
	private tracked(tr: Transaction, from: number, to: number, context: DocxEditContext): Transaction {
		const stamp = context.review.tracking ? revisionStamp(context.identity) : undefined

		if (stamp === undefined) return tr

		return tr.addMark(from, to, marks.insertion.create({ src: NO_SOURCE, id: context.ids.next("revision"), ...stamp }))
	}

	private referenceNode(kind: DocxNoteKind, id: string, custom: boolean, doc: Node, at: number): Node {
		const $at = doc.resolve(at)
		const around = [...($at.nodeBefore?.marks ?? $at.nodeAfter?.marks ?? [])].filter(mark => WRAPPERS.has(mark.type))
		let nodeMarks: readonly Mark[] = referenceMarks(this.context, kind)

		for (const mark of around) nodeMarks = mark.addToSet(nodeMarks)

		return nodes.note_ref.create({ src: NO_SOURCE, kind, id, custom }, null, nodeMarks)
	}

	// A new note's story holding `content` (blocks, or made for the part the note is on; by default one paragraph
	// of the note's text style with its mark and a space, as Word starts a note), outside the undo history: the
	// reference that refers to it is what undoes.
	private newNote(
		kind: DocxNoteKind,
		mark: string | undefined,
		content: Fragment | ((part: number) => Fragment | undefined) | undefined
	): { id: string; story: number } | undefined {
		const context = this.context
		const session = this.editor.addStory(kind)
		const id = session === undefined ? undefined : context.newStories.get(session.story.story)?.id

		if (session === undefined || id === undefined) return undefined

		const made = typeof content === "function" ? content(session.story.part.index) : content

		// Nothing refers to the story: it is left out of the save.
		if (typeof content === "function" && made === undefined) return undefined

		const blocks = made ?? Fragment.from(this.firstParagraph(kind, mark, Fragment.from(docxSchema.text(" "))))
		const tr = session.state.tr.replaceWith(0, session.state.doc.content.size, blocks)
		const tracked = this.tracked(tr, 0, tr.doc.content.size, context)

		if (!session.dispatch(markInternal(tracked).setMeta("addToHistory", false))) return undefined

		return { id, story: session.story.story.index }
	}

	// A note's first paragraph: its mark, then `rest`.
	private firstParagraph(kind: DocxNoteKind, mark: string | undefined, rest: Fragment, attrs?: Readonly<Record<string, unknown>>): Node {
		const context = this.context
		const style = referenceMarks(context, kind)
		const lead =
			mark === undefined
				? nodes.run_object.create({ src: NO_SOURCE, kind: MARK_KINDS[kind], text: "" }, null, style)
				: docxSchema.text(mark, style)

		return nodes.paragraph.create(
			{ ...attrs, styleId: attrs?.["styleId"] ?? textStyle(context, kind) },
			Fragment.from(lead).append(rest)
		)
	}

	// The blocks of note story `story` as content of a new note of `kind` on part `target`: made new (as a paste
	// into another part makes them), the note's mark of the new kind, the text style mapped from the old kind's.
	// A note with a custom mark keeps its content (the mark is its text) and its reference stays custom.
	private convertedContent(story: Story, kind: DocxNoteKind, custom: boolean, target: number): Fragment | undefined {
		const session = this.editor.storySession(story.index)

		if (session === undefined) return undefined

		const context = this.context
		// What only the note's part holds (fills of its cells, form fields' data) stays with the copy.
		const copies = sourceCopies(context)
		const from = otherKind(kind)
		const fromStyle = textStyle(context, from)
		const toStyle = textStyle(context, kind)
		const blocks: Node[] = []
		let marked = false

		session.state.doc.forEach(block => {
			const fresh = freshSlice(new Slice(Fragment.from(block), 0, 0), copies, session.story.part.index, target).content.firstChild

			if (fresh === null) return

			const style = fresh.attrs["styleId"]
			const attrs = style !== null && style === fromStyle ? { ...fresh.attrs, styleId: toStyle } : fresh.attrs

			if (!marked && !custom && fresh.type === nodes.paragraph) {
				marked = true

				// Its own mark (and the space Word puts after it) gives way to the new kind's.
				const rest: Node[] = []

				block.forEach((child, _offset, index) => {
					if (
						index === 0 &&
						child.type === nodes.run_object &&
						(child.attrs["kind"] === "FootnoteMark" || child.attrs["kind"] === "EndnoteMark")
					)
						return

					rest.push(child)
				})

				const content = freshSlice(new Slice(Fragment.fromArray(rest), 0, 0), copies, session.story.part.index, target).content

				blocks.push(this.firstParagraph(kind, undefined, content, attrs))

				return
			}

			blocks.push(fresh.type.create(attrs, fresh.content, fresh.marks))
		})

		return blocks.length === 0 ? undefined : Fragment.fromArray(blocks)
	}

	private convertNotes(found: readonly { readonly kind: DocxNoteKind; readonly id: string; readonly pos: number }[]): boolean {
		const session = this.body
		const tr = session.state.tr
		let converted = 0

		for (const note of found) {
			const story = this.storyOf(note.kind, note.id)
			const kind = otherKind(note.kind)
			const pos = tr.mapping.map(note.pos)
			const old = tr.doc.nodeAt(pos)
			const custom = old?.attrs["custom"] === true
			const made =
				story === undefined || old?.type !== nodes.note_ref
					? undefined
					: this.newNote(kind, undefined, target => this.convertedContent(story, kind, custom, target))

			if (made === undefined || old === null) continue

			const nodeMarks = old.marks.map(mark =>
				mark.type === marks.char_style ? (referenceMarks(this.context, kind)[0] ?? mark) : mark
			)

			tr.replaceWith(pos, pos + 1, nodes.note_ref.create({ src: NO_SOURCE, kind, id: made.id, custom }, null, nodeMarks))
			converted++
		}

		return converted > 0 && session.dispatch(markInternal(tr))
	}
}
