import { Fragment, Slice } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import type { DocxEditContext } from "../edit/context"
import { builtInStyleId } from "../edit/formats"
import { freshSlice, markInternal } from "../edit/invariants/index"
import { docxSchema } from "../edit/schema"
import type { DocxEditorSession } from "../edit/session"
import type { HeaderFooterSet, Section } from "../model/index"
import type { PagedEditor } from "../paged/editor"
import { hasStyleDefinition } from "../write/styles"
import {
	bodySectionsOf,
	docSectionChanges,
	editSections,
	evenAndOddHeaders,
	pageOfCaret,
	registerSectionStory,
	sectionDocument,
	sectionNode,
	setEvenAndOddHeaders,
	setTitlePage,
	type DocxHeaderKind,
	type DocxHeaderType,
	type HeaderReferenceEdits,
	type SectionState
} from "./sections"

// Headers and footers: which one a page shows (its section's, of the type the page takes: the first page of a
// section with w:titlePg, an even page where the document has different even and odd ones, else the default),
// entering it, making one where a section shows none, linking a section's to the previous section's (it has
// none of its own) or unlinking it (its own, a copy of what it showed, as Word copies it), and removing one.

const nodes = docxSchema.nodes

// Where the header or footer being edited is shown, for the "Header — Close" chip and its toggles.
export interface DocxHeaderContext {
	readonly kind: DocxHeaderKind
	// The section of the page the caret is on (0-based), of `sections`.
	readonly section: number
	readonly sections: number
	readonly type: DocxHeaderType
	// The section's first page shows its own header and footer (w:titlePg).
	readonly differentFirst: boolean
	// Even pages show their own (document-wide w:evenAndOddHeaders).
	readonly differentOddEven: boolean
	// The section has none of its own of this type: it shows the previous section's ("Same as Previous").
	readonly linkedToPrevious: boolean
	// The first section links to nothing.
	readonly canLink: boolean
}

// The paged editor of each document's context, as the document commands see it (the snapshot's story context
// reads the pages through it).
const EDITORS = new WeakMap<DocxEditContext, PagedEditor>()

export function bindEditor(context: DocxEditContext, editor: PagedEditor): void {
	if (EDITORS.get(context) === editor) return

	EDITORS.set(context, editor)

	// A snapshot made before the editor was known is made again with what it shows.
	const kind = editor.session.story.story.kind

	if (kind === "header" || kind === "footer") editor.session.invalidate()
}

export function boundEditor(context: DocxEditContext): PagedEditor | undefined {
	return EDITORS.get(context)
}

function setOf(section: Section | SectionState, kind: DocxHeaderKind, effective: boolean): HeaderFooterSet {
	if (kind === "header") return effective ? section.effectiveHeaders : section.headers

	return effective ? section.effectiveFooters : section.footers
}

// Whether section `section` (as layout lays it out) shows a different first page.
function titlePageOf(editor: PagedEditor, section: Section): boolean {
	const body = editor.body.session
	const document = sectionDocument(section) ?? body.story.part.document

	return section.sectPr >= 0 && body.context.store.section(document, section.sectPr).titlePage === true
}

// The type of header and footer page `page` shows, and its section (as layout laid it out).
export function shownType(editor: PagedEditor, page: number): { readonly section: number; readonly type: DocxHeaderType } | undefined {
	const list = editor.page(page)
	const sections = editor.edits.model.sections
	const section = list === undefined ? undefined : sections[list.section]

	if (list === undefined || section === undefined) return undefined

	const first = page === 0 || editor.page(page - 1)?.section !== list.section

	if (first && titlePageOf(editor, section)) return { section: list.section, type: "first" }
	if (evenAndOddHeaders(editor.body.session.context, editor.body.session.state.doc) && list.number % 2 === 0)
		return { section: list.section, type: "even" }

	return { section: list.section, type: "default" }
}

// The page the caret is on: the active story's (a header shows on many: the one it was entered on).
function caretPage(editor: PagedEditor): number {
	return pageOfCaret(editor) ?? 0
}

// The header or footer context of the story being edited, or null when it is neither.
export function headerContext(editor: PagedEditor): DocxHeaderContext | null {
	const kind = editor.session.story.story.kind

	if (kind !== "header" && kind !== "footer") return null

	const shown = shownType(editor, caretPage(editor))
	const sections = editor.edits.model.sections
	const section = shown === undefined ? undefined : sections[shown.section]

	if (shown === undefined || section === undefined) return null

	return {
		kind,
		section: shown.section,
		sections: sections.length,
		type: shown.type,
		differentFirst: titlePageOf(editor, section),
		differentOddEven: evenAndOddHeaders(editor.body.session.context, editor.body.session.state.doc),
		linkedToPrevious: shown.section > 0 && setOf(section, kind, false)[shown.type] < 0,
		canLink: shown.section > 0
	}
}

// ---- commands -------------------------------------------------------------------------------------------------

function writable(editor: PagedEditor): boolean {
	return !editor.body.session.readOnly && editor.pageCount > 0
}

export function canEditHeader(editor: PagedEditor): boolean {
	return writable(editor)
}

// A new header or footer story holding `content` (one empty paragraph of the Header or Footer style by
// default), outside the undo history: the section edit referring to it is what undoes.
function newStory(editor: PagedEditor, kind: DocxHeaderKind, content: Fragment | undefined): number | undefined {
	const session = editor.addStory(kind)

	if (session === undefined) return undefined

	const context = session.context
	const style = context.styles.byName(kind, "paragraph")?.id ?? (hasStyleDefinition(kind) ? builtInStyleId(context, kind) : null)
	const blocks = content ?? Fragment.from(nodes.paragraph.create({ styleId: style }))
	const tr = session.state.tr.replaceWith(0, session.state.doc.content.size, blocks)

	if (!session.dispatch(markInternal(tr).setMeta("addToHistory", false))) return undefined

	registerSectionStory(context, session.story.story)

	return session.story.story.index
}

// A story's content made new (as a paste into another part makes it): what an unlinked section gets.
function copiedContent(editor: PagedEditor, story: number): Fragment | undefined {
	const session: DocxEditorSession | undefined = editor.storySession(story)
	const doc = session?.state.doc

	if (doc === undefined) return undefined

	const fresh = freshSlice(new Slice(doc.content, 0, 0)).content

	return fresh.childCount === 0 ? undefined : fresh
}

function references(
	kind: DocxHeaderKind,
	type: DocxHeaderType,
	story: number | null
): { headers?: HeaderReferenceEdits; footers?: HeaderReferenceEdits } {
	return kind === "header" ? { headers: { [type]: story } } : { footers: { [type]: story } }
}

function withReference(
	edit: { headers?: HeaderReferenceEdits; footers?: HeaderReferenceEdits },
	kind: DocxHeaderKind,
	type: DocxHeaderType,
	story: number | null
): { headers?: HeaderReferenceEdits; footers?: HeaderReferenceEdits } {
	const key = kind === "header" ? "headers" : "footers"

	return { [key]: { ...edit[key], ...references(kind, type, story)[key] } }
}

// Enters the header (or footer) of the caret's page, making one where its section shows none: the reference
// goes to the first section of the run of sections showing none of the type, so all of them show it (they are
// linked), as Word does.
export function editHeader(editor: PagedEditor, kind: DocxHeaderKind, page: number = caretPage(editor)): boolean {
	if (!canEditHeader(editor)) return false

	const shown = shownType(editor, page)

	if (shown === undefined) return false

	const { states } = bodySectionsOf(editor)
	const state = states[shown.section]

	if (state === undefined) return false

	const effective = setOf(state, kind, true)[shown.type]

	if (effective >= 0) return enter(editor, effective, page)

	let first = shown.section

	while (first > 0 && setOf(states[first - 1] as SectionState, kind, false)[shown.type] < 0) first--

	const story = newStory(editor, kind, undefined)

	if (story === undefined) return false

	const set = editSections(editor, [first], (_state, edit) => ({ ...edit, ...withReference(edit, kind, shown.type, story) }), false)

	return set && enter(editor, story, page)
}

function enter(editor: PagedEditor, story: number, page: number): boolean {
	if (!editor.enterStory(story, page)) return false

	const session = editor.session

	return session.dispatch(session.state.tr.setSelection(TextSelection.atStart(session.state.doc)).scrollIntoView())
}

// Whether the header or footer being edited can be linked to (or unlinked from) the previous section's.
export function canLink(editor: PagedEditor): boolean {
	const kind = editor.session.story.story.kind

	if ((kind !== "header" && kind !== "footer") || !writable(editor)) return false

	const shown = shownType(editor, caretPage(editor))

	return shown !== undefined && shown.section > 0
}

// Link to previous, toggled for the header or footer being edited: a section with its own loses it and shows
// the previous section's; one showing the previous section's gets its own, a copy of it.
export function toggleLink(editor: PagedEditor): boolean {
	if (!canLink(editor)) return false

	const kind = editor.session.story.story.kind as DocxHeaderKind
	const page = caretPage(editor)
	const shown = shownType(editor, page)
	const { states } = bodySectionsOf(editor)
	const state = shown === undefined ? undefined : states[shown.section]

	if (shown === undefined || state === undefined) return false

	const own = setOf(state, kind, false)[shown.type]

	if (own >= 0) {
		const linked = editSections(
			editor,
			[shown.section],
			(_state, edit) => ({ ...edit, ...withReference(edit, kind, shown.type, null) }),
			false
		)
		const now = bodySectionsOf(editor).states[shown.section]
		const showing = now === undefined ? -1 : setOf(now, kind, true)[shown.type]

		if (!linked) return false
		if (showing >= 0) return enter(editor, showing, page)

		editor.exitStory()

		return true
	}

	const previous = setOf(state, kind, true)[shown.type]
	const story = newStory(editor, kind, previous >= 0 ? copiedContent(editor, previous) : undefined)

	if (story === undefined) return false

	const unlinked = editSections(
		editor,
		[shown.section],
		(_state, edit) => ({ ...edit, ...withReference(edit, kind, shown.type, story) }),
		false
	)

	return unlinked && enter(editor, story, page)
}

// Whether the caret's page shows a header (footer) to remove.
export function canRemove(editor: PagedEditor, kind: DocxHeaderKind): boolean {
	if (!writable(editor)) return false

	const shown = shownType(editor, caretPage(editor))
	const section = shown === undefined ? undefined : editor.edits.model.sections[shown.section]

	return shown !== undefined && section !== undefined && setOf(section, kind, true)[shown.type] >= 0
}

// Removes the header (footer) the caret's page shows: every section referring to it loses the reference, so the
// sections linked to it lose it too.
export function removeHeader(editor: PagedEditor, kind: DocxHeaderKind): boolean {
	if (!canRemove(editor, kind)) return false

	const page = caretPage(editor)
	const shown = shownType(editor, page)
	const { states } = bodySectionsOf(editor)
	const state = shown === undefined ? undefined : states[shown.section]

	if (shown === undefined || state === undefined) return false

	const story = setOf(state, kind, true)[shown.type]
	const holders = states.filter(candidate => setOf(candidate, kind, false)[shown.type] === story).map(candidate => candidate.index)
	const leaving = editor.session.story.story.index === story

	if (!editSections(editor, holders, (_state, edit) => ({ ...edit, ...withReference(edit, kind, shown.type, null) }), false)) return false
	if (leaving) editor.exitStory()

	return true
}

// After a switch that changes which header and footer a page shows: the one the page edited on shows now is
// entered (made where it has none, as Word shows it ready to type in).
function reenter(editor: PagedEditor, kind: DocxHeaderKind | undefined, page: number, done: boolean): boolean {
	if (done && kind !== undefined) editHeader(editor, kind, page)

	return done
}

function editedKind(editor: PagedEditor): DocxHeaderKind | undefined {
	const kind = editor.session.story.story.kind

	return kind === "header" || kind === "footer" ? kind : undefined
}

export function toggleTitlePage(editor: PagedEditor): boolean {
	const page = caretPage(editor)
	const shown = shownType(editor, page)
	const { states } = bodySectionsOf(editor)
	const state = shown === undefined ? undefined : states[shown.section]
	const body = editor.body.session
	const node = state === undefined ? -1 : sectionNode(body.story.part, state.edit?.base ?? state.key)
	const current =
		state?.edit?.titlePage ?? (node >= 0 ? body.context.store.section(body.story.part.document, node).titlePage === true : false)

	return reenter(editor, editedKind(editor), page, setTitlePage(editor, !current))
}

export function toggleEvenAndOdd(editor: PagedEditor): boolean {
	const body = editor.body.session
	const page = caretPage(editor)
	const on = !(docSectionChanges(body.state.doc).evenAndOddHeaders ?? body.context.model.settings.evenAndOddHeaders)

	return reenter(editor, editedKind(editor), page, setEvenAndOddHeaders(editor, on))
}
