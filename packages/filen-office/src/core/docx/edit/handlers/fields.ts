import { closeHistory } from "prosemirror-history"
import type { Node } from "prosemirror-model"
import { NodeSelection, type EditorState, type Transaction } from "prosemirror-state"
import { addBookmark, bookmarkIndex, canAddBookmark, deleteBookmark, renameBookmark } from "../../fields/bookmarks"
import { FieldEvaluation, pagedFieldHost, parseFieldCode, updatableField, type FieldHost } from "../../fields/evaluate"
import {
	addTocBookmarks,
	buildToc,
	partHasToc,
	takenElsewhere,
	tocAt,
	tocEntries,
	tocInstruction,
	tocOptions,
	tocRegions,
	type TocRegion
} from "../../fields/toc"
import type { PagedEditor } from "../../paged/editor"
import { markable } from "../../review/track"
import { AttrChanges } from "../attrs"
import { markCommand, markInternal } from "../invariants/index"
import type { DocumentCommandContext, DocumentHandlers, FeatureHandlers } from "../registry"
import { NO_MARKERS, NO_SOURCE, docxSchema } from "../schema"
import type { DocxEditorSession } from "../session"

// Fields, bookmarks and tables of contents (document commands: an update reads the laid-out pages). Updates
// are Word's manual ones: Update Field (F9) on the selection or the field at the caret, Update All on the
// story, Update Table on a table of contents. Each waits for layout to settle once, computes every result
// from the pages as they are and applies them as one step (no layout per field); a table rebuilt or inserted
// lays out once more and gets its page numbers in a second step that undoes with the first. Updates are
// not recorded as revisions (a result is the field's bookkeeping); an insertion is, while tracking.

const nodes = docxSchema.nodes

// An undo step a command began: the document it left, and an empty transaction a later step names as the one
// it is appended to (the step's own would keep every intermediate document while layout catches up).
interface Joined {
	readonly doc: Node
	readonly tr: Transaction
}

function writable(context: DocumentCommandContext): boolean {
	return !context.editor.body.session.readOnly && !context.session.readOnly
}

function inParagraph(state: EditorState): boolean {
	const { $from, $to } = state.selection

	return $from.parent.type === nodes.paragraph && $to.parent.type === nodes.paragraph
}

// The field atom at the caret (selected, or just before or after it).
function fieldAtCaret(state: EditorState): { node: Node; pos: number } | undefined {
	const selection = state.selection

	if (selection instanceof NodeSelection)
		return selection.node.type === nodes.field ? { node: selection.node, pos: selection.from } : undefined
	if (!selection.empty) return undefined

	const $pos = selection.$from
	const after = $pos.nodeAfter
	const before = $pos.nodeBefore

	if (after?.type === nodes.field) return { node: after, pos: $pos.pos }
	if (before?.type === nodes.field) return { node: before, pos: $pos.pos - before.nodeSize }

	return undefined
}

// Whether the selection is on a table of contents: a locked block of one of the file's selected, or a paragraph
// holding field characters of one the editor made.
function onToc(state: EditorState): boolean {
	const { $from } = state.selection
	const selected = $from.nodeAfter

	if (selected?.type === nodes.locked_block && selected.attrs["reason"] === "fieldSpansBlocks") return true

	let found = false

	$from.parent.forEach(child => {
		if (child.type === nodes.field_char) found = true
	})

	return found
}

// Fields are recorded as tracked insertions where tracked changes mark field characters and field atoms (see
// review/track.ts markable); until then an insertion of one is applied untracked rather than half recorded.
const FIELDS_TRACKED = markable(nodes.field_char.create({ kind: "end" })) && markable(nodes.field.create())

function inserting(session: DocxEditorSession, tr: Transaction): Transaction {
	return session.context.review.tracking && !FIELDS_TRACKED ? markCommand(tr) : tr
}

// Stories a table of contents was inserted into this session (Update Table stays offered there).
const INSERTED = new WeakSet<DocxEditorSession>()

function hasToc(session: DocxEditorSession): boolean {
	return INSERTED.has(session) || partHasToc(session.story.part) || onToc(session.state)
}

// Applies `tr` to `session` as part of the undo step `joined` began, when nothing changed the document since.
function dispatchJoined(session: DocxEditorSession, tr: Transaction, joined: Joined | undefined): Joined | undefined {
	if (joined !== undefined && session.state.doc === joined.doc) tr.setMeta("appendedTransaction", joined.tr)
	else closeHistory(tr)

	return session.dispatch(tr) ? { doc: session.state.doc, tr: session.state.tr } : undefined
}

// New results for the fields of [from, to) of the session's document.
function fieldChanges(host: FieldHost, session: DocxEditorSession, from: number, to: number): AttrChanges {
	const evaluation = new FieldEvaluation(host, session, session.context.identity.now())
	const changes = new AttrChanges()

	session.state.doc.nodesBetween(from, to, (node, pos) => {
		if (node.type !== nodes.field) return !node.isLeaf
		if (!updatableField(session, node)) return false

		const result = evaluation.result(node, pos)

		if (result !== undefined && result !== node.attrs["result"]) changes.set(node, pos, { result })

		return false
	})

	return changes
}

function applyChanges(session: DocxEditorSession, changes: AttrChanges, joined: Joined | undefined): Joined | undefined {
	if (changes.empty) return joined

	const tr = session.state.tr

	changes.step(tr)

	return dispatchJoined(session, markCommand(tr), joined)
}

// The page number a heading is on, as shown.
function headingPage(host: FieldHost, session: DocxEditorSession, pos: number): string {
	const pages = host.pages
	const index = pages?.pageAt(session, pos + 1)

	return (index === undefined ? undefined : pages?.page(index))?.numberText ?? ""
}

// Where a new table goes: before the caret's paragraph when the caret is at its start, after it at its end,
// else between its halves (a section break stays with the second).
function insertionPoint(tr: Transaction): number | undefined {
	const { $from } = tr.selection

	if ($from.parent.type !== nodes.paragraph) return undefined
	if ($from.parentOffset === 0) return $from.before()
	if ($from.parentOffset === $from.parent.content.size && ($from.parent.attrs["sectPr"] as number) === NO_SOURCE) return $from.after()

	const pos = $from.pos

	tr.split(pos)

	return pos + 1
}

// The page numbers of the table at `pos`, once layout caught up with it.
async function refreshToc(editor: PagedEditor, session: DocxEditorSession, pos: number, joined: Joined): Promise<Joined> {
	await editor.whenSettled()

	const region = tocRegions(session).find(candidate => candidate.from <= pos && pos < candidate.to)

	if (region === undefined || session.readOnly) return joined

	return applyChanges(session, fieldChanges(pagedFieldHost(editor), session, region.from, region.to), joined) ?? joined
}

// A table of contents with `instruction` in place of `region` (else at the selection), its entries' page
// numbers given once the pages are laid out again.
async function placeToc(
	editor: PagedEditor,
	session: DocxEditorSession,
	instruction: string,
	region: TocRegion | undefined,
	joined: Joined | undefined
): Promise<Joined | undefined> {
	const options = tocOptions(parseFieldCode(instruction))

	if (!options.supported) return undefined

	const host = pagedFieldHost(editor)
	const entries = tocEntries(host, session, options, region?.from ?? -1, region?.to ?? -1)
	const tr = session.state.tr
	let at: number | undefined = region?.from

	if (region === undefined) {
		if (!tr.selection.empty) tr.deleteSelection()

		at = insertionPoint(tr)
	}

	if (at === undefined) return undefined

	const build = buildToc({
		session,
		host,
		instruction,
		entries,
		pos: at,
		templates: region?.templates ?? new Map<number, number>(),
		markers: region?.markers ?? NO_MARKERS,
		sectPr: region?.sectPr ?? NO_SOURCE,
		taken: takenElsewhere(host.sessions, session),
		page: entry => headingPage(host, session, entry.pos)
	})

	const steps = tr.steps.length

	if (region === undefined) tr.insert(at, build.paragraphs)
	else tr.replaceWith(region.from, region.to, build.paragraphs)

	addTocBookmarks(tr, build)

	// A table of the file is locked blocks no edit may delete: replaced by the session's own step.
	const first = dispatchJoined(session, region === undefined ? inserting(session, tr.scrollIntoView()) : markInternal(tr), joined)

	if (first === undefined) return undefined
	if (region === undefined) INSERTED.add(session)

	return refreshToc(editor, session, tr.mapping.slice(steps).map(at, -1), first)
}

// Field results of the selection (or the field at the caret), or of the whole story, brought up to date, with
// the tables of contents of the file there.
async function update(editor: PagedEditor, all: boolean): Promise<boolean> {
	const session = editor.session

	await editor.whenSettled()

	if (session.readOnly || editor.session !== session) return false

	const state = session.state
	const caret = all ? undefined : fieldAtCaret(state)
	const from = all ? 0 : (caret?.pos ?? state.selection.from)
	const to = all ? state.doc.content.size : caret === undefined ? state.selection.to : caret.pos + caret.node.nodeSize
	const tables = tocRegions(session).filter(
		region => region.source && region.updatable && region.from < Math.max(to, from + 1) && region.to > from
	)
	let joined = applyChanges(session, fieldChanges(pagedFieldHost(editor), session, from, to), undefined)

	// The file's tables are rebuilt (their entries are bytes); the last first, so earlier ones keep their places.
	for (const table of tables.reverse()) {
		await editor.whenSettled()

		const current = tocRegions(session).find(region => region.source && region.from === table.from && region.to === table.to)

		if (current !== undefined) joined = (await placeToc(editor, session, current.instruction, current, joined)) ?? joined
	}

	return true
}

async function updateToc(editor: PagedEditor, pagesOnly: boolean): Promise<boolean> {
	const session = editor.session

	await editor.whenSettled()

	if (session.readOnly || editor.session !== session) return false

	const region = tocAt(session)

	if (region === undefined || !region.updatable) return false

	// Page numbers alone are fields of the editor's tables; the file's is rebuilt whole (its entries are bytes).
	if (pagesOnly && !region.source) {
		applyChanges(session, fieldChanges(pagedFieldHost(editor), session, region.from, region.to), undefined)

		return true
	}

	return (await placeToc(editor, session, region.instruction, region, undefined)) !== undefined
}

async function insertToc(editor: PagedEditor, levels: number, hyperlinks: boolean): Promise<boolean> {
	const session = editor.session

	await editor.whenSettled()

	if (session.readOnly || editor.session !== session || !inParagraph(session.state)) return false

	return (await placeToc(editor, session, tocInstruction(levels, hyperlinks), undefined, undefined)) !== undefined
}

// A new field at the selection, its result computed from the pages as they are.
function insertField(editor: PagedEditor, instruction: string): boolean {
	const session = editor.session
	const code = parseFieldCode(instruction)

	if (session.readOnly || !inParagraph(session.state) || code.type === "") return false

	if (code.type === "REF" || code.type === "PAGEREF") {
		const name = code.args[0]?.toLowerCase() ?? ""
		const body = editor.body.session

		if (
			!bookmarkIndex(session.state.doc, session.story.part).byName.has(name) &&
			!bookmarkIndex(body.state.doc, body.story.part).byName.has(name)
		)
			return false
	}

	const field = nodes.field.create({
		src: NO_SOURCE,
		end: NO_SOURCE,
		simple: false,
		type: code.type,
		instruction: ` ${instruction.trim()} `,
		result: ""
	})
	const tr = session.state.tr.replaceSelectionWith(field, false)
	const pos = tr.selection.from - field.nodeSize

	if (tr.doc.nodeAt(pos)?.type !== nodes.field) return false

	const result = new FieldEvaluation(pagedFieldHost(editor), session, session.context.identity.now(), tr.doc).result(field, pos)

	if (result !== undefined && result !== "") tr.setNodeAttribute(pos, "result", result)

	return session.dispatch(inserting(session, tr.scrollIntoView()))
}

const insertable = (context: DocumentCommandContext): boolean => writable(context) && inParagraph(context.session.state)

const document: DocumentHandlers = {
	"field.update": {
		enabled: context => {
			const state = context.session.state

			return writable(context) && (!state.selection.empty || fieldAtCaret(state) !== undefined || onToc(state))
		},
		run: context => update(context.editor, false)
	},
	"field.updateAll": {
		enabled: writable,
		run: context => update(context.editor, true)
	},
	"toc.insert": {
		enabled: context => insertable(context) && context.session === context.editor.body.session,
		run: (context, command) => insertToc(context.editor, command.levels, command.hyperlinks)
	},
	"toc.update": {
		enabled: context => writable(context) && hasToc(context.session),
		run: (context, command) => updateToc(context.editor, command.pagesOnly)
	},
	"insert.pageNumber": {
		enabled: insertable,
		run: context => insertField(context.editor, "PAGE \\* MERGEFORMAT")
	},
	"insert.pageCount": {
		enabled: insertable,
		run: context => insertField(context.editor, "NUMPAGES \\* MERGEFORMAT")
	},
	"insert.date": {
		enabled: insertable,
		run: (context, command) => insertField(context.editor, `DATE \\@ "${command.format}"`)
	},
	"insert.field": {
		enabled: insertable,
		run: (context, command) => insertField(context.editor, command.instruction)
	},
	"bookmark.add": {
		enabled: context => writable(context) && canAddBookmark(context.session),
		run: (context, command) => addBookmark(context.session, context.editor.sessions, command.name)
	},
	"bookmark.delete": {
		enabled: writable,
		run: (context, command) => deleteBookmark(context.session, command.name)
	},
	"bookmark.rename": {
		enabled: writable,
		run: (context, command) => renameBookmark(context.session, context.editor.sessions, command.name, command.to)
	}
}

export const FIELD_HANDLERS: FeatureHandlers = { document }
