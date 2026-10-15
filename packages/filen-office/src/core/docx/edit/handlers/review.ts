import { Selection, TextSelection } from "prosemirror-state"
import { storyHasTrackedEquations } from "../../math/edit"
import type { PagedEditor } from "../../paged/editor"
import { RevisionView } from "../../model/index"
import {
	movedChange,
	rememberMovedChange,
	resolveChanges,
	reviewIndexWanted,
	storyReviewIndex,
	wantReviewIndex,
	type ReviewScope
} from "../../review/accept"
import { ReviewKind } from "../../review/index"
import { trackTransaction } from "../../review/track"
import type { DocxMarkupView } from "../commands/review"
import type { DocumentCommandContext, DocumentCommandHandler, DocxTransactionTransformer, FeatureHandlers } from "../registry"
import type { DocxEditorSession } from "../session"

// Tracked changes: recording (the transformer), the track switch, accept and reject (at the selection, by id,
// all), moving between changes, and the markup view.

const VIEWS: Readonly<Record<DocxMarkupView, RevisionView>> = {
	all: RevisionView.Markup,
	simple: RevisionView.Final,
	none: RevisionView.Final,
	original: RevisionView.Original
}

function revisionCount(session: DocxEditorSession): number {
	return reviewIndexWanted(session) ? storyReviewIndex(session).revisionCount : 0
}

// The context changed outside a transaction (tracking switched): every session makes its snapshot again.
function refresh(editor: PagedEditor): void {
	for (const session of editor.sessions) session.invalidate()
}

// The stories of a model (by index) whose source has tracked changes (inside its equations too): read once per
// model.
const changedStoriesOf = new WeakMap<object, readonly number[]>()

function changedStories(editor: PagedEditor): readonly number[] {
	const model = editor.body.session.context.model
	let found = changedStoriesOf.get(model)

	if (found !== undefined) return found

	const out = new Set<number>()

	for (const story of model.stories) {
		const part = model.parts[story.part]

		if (part === undefined || story.parent >= 0 || story.item < 0) continue

		// Tracked changes inside its equations (one item each, in no revision table).
		if (storyHasTrackedEquations(part, story.item)) {
			out.add(story.index)

			continue
		}

		if (part.revisions.count === 0) continue

		const document = part.document
		const start = document.start(part.node(story.item))
		const end = document.end(part.node(story.item))

		for (let k = 0; k < part.revisions.count; k++) {
			const at = document.start(part.revisions.node[k] ?? -1)

			if (at >= start && at < end) {
				out.add(story.index)

				break
			}
		}
	}

	found = [...out]
	changedStoriesOf.set(model, found)

	return found
}

// Whether the document has changes: a session's index counts its story (resolved ones are gone), the source
// counts stories no session indexes.
function hasChanges(editor: PagedEditor): boolean {
	const indexed = new Set<number>()

	for (const session of editor.sessions) {
		if (!reviewIndexWanted(session)) continue
		if (storyReviewIndex(session).revisionCount > 0) return true

		indexed.add(session.story.story.index)
	}

	return changedStories(editor).some(index => !indexed.has(index))
}

// Every editable session with changes: those the editor has, and the stories whose source has changes.
function sessionsWithChanges(editor: PagedEditor): DocxEditorSession[] {
	const out = editor.sessions.filter(session => !session.readOnly)
	const opened = new Set(out.map(session => session.story.story.index))

	for (const index of changedStories(editor)) {
		if (opened.has(index)) continue

		const session = editor.storySession(index)

		if (session !== undefined && !session.readOnly) out.push(session)
	}

	return out
}

function resolve(sessions: readonly DocxEditorSession[], scope: ReviewScope, accept: boolean): boolean {
	let done = false

	for (const session of sessions) {
		if (!reviewIndexWanted(session)) continue

		const outcome = resolveChanges(session, storyReviewIndex(session), scope, accept)

		if (outcome.tr !== null && session.dispatch(outcome.tr)) done = true
	}

	return done
}

const resolveOne = <T extends "review.accept" | "review.reject">(accept: boolean): DocumentCommandHandler<T> => ({
	enabled: context => revisionCount(context.session) > 0,
	run(context, command) {
		const selection = context.session.state.selection
		const id = (command as { readonly id?: string }).id
		const scope: ReviewScope = id !== undefined ? { id } : { from: selection.from, to: selection.to }

		return resolve([context.session], scope, accept)
	}
})

const resolveAll = <T extends "review.acceptAll" | "review.rejectAll">(accept: boolean): DocumentCommandHandler<T> => ({
	enabled: context => hasChanges(context.editor),
	run: context => resolve(sessionsWithChanges(context.editor), { all: true }, accept)
})

// Selects the next (previous) change of the story being edited, whole (from its first piece to its last).
// Changes go in the order of their first pieces (one inside another's span, as a paragraph's property change
// inside a deletion across paragraphs, after it); past the last (first) one, the first (last) again. The
// change at the selection (the one a move selected, or the narrowest the caret is inside) is skipped.
const move = <T extends "review.next" | "review.prev">(forward: boolean): DocumentCommandHandler<T> => ({
	enabled: context => revisionCount(context.session) > 0,
	run(context: DocumentCommandContext) {
		const session = context.session
		const state = session.state
		const index = storyReviewIndex(session)
		const spans = new Map<string, { from: number; to: number }>()

		for (const entry of index.entries()) {
			if (entry.kind <= ReviewKind.COMMENT_REFERENCE) continue

			const span = spans.get(entry.id)

			if (span === undefined) spans.set(entry.id, { from: entry.from, to: entry.to })
			else {
				span.from = Math.min(span.from, entry.from)
				span.to = Math.max(span.to, entry.to)
			}
		}

		const changes = [...spans].sort((a, b) => a[1].from - b[1].from || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))

		if (changes.length === 0) return false

		const { from, to } = state.selection
		const last = movedChange(session)
		let current: string | undefined

		if (last !== undefined && spans.has(last)) current = last
		else {
			const here = index.changeAt(from)

			if (here !== undefined && (from < to || (here.from < from && Math.max(here.to, here.from + 1) > from))) current = here.id
		}

		const at = changes.findIndex(([id]) => id === current)
		let next: number

		if (at >= 0) next = (at + (forward ? 1 : -1) + changes.length) % changes.length
		else {
			next = forward ? changes.findIndex(([, span]) => span.from >= to) : changes.findLastIndex(([, span]) => span.from < from)

			if (next < 0) next = forward ? 0 : changes.length - 1
		}

		const target = changes[next]

		if (target === undefined) return false

		const [id, span] = target
		const doc = state.doc
		const size = doc.content.size
		const start = Math.min(size, span.from)
		const end = Math.min(size, span.to)
		const $from = doc.resolve(start)
		// A paragraph mark's entry ends past its paragraph.
		const stop = doc.resolve(end).parent.inlineContent || end === start ? end : end - 1
		const selection =
			$from.parent.inlineContent && doc.resolve(stop).parent.inlineContent
				? TextSelection.create(doc, start, stop)
				: Selection.near($from)

		const tr = state.tr.setSelection(selection).scrollIntoView()

		rememberMovedChange(session, tr.doc, tr.selection, id)

		return session.dispatch(tr)
	}
})

export const REVIEW_HANDLERS: FeatureHandlers = {
	document: {
		"review.track": {
			enabled: () => true,
			run(context, command) {
				const review = context.context.review
				const on = command.on ?? !review.tracking

				// Recording names its author (the app asks for it first).
				if (on && (context.context.identity.author ?? "") === "") return false
				if (review.tracking === on) return true

				review.tracking = on

				if (on) for (const session of context.editor.sessions) wantReviewIndex(session)

				refresh(context.editor)

				return true
			}
		},
		"review.accept": resolveOne(true),
		"review.reject": resolveOne(false),
		"review.acceptAll": resolveAll(true),
		"review.rejectAll": resolveAll(false),
		"review.next": move(true),
		"review.prev": move(false),
		"view.markup": {
			enabled: () => true,
			run(context, command) {
				context.editor.setView(VIEWS[command.mode])

				return true
			}
		}
	}
}

// Records edits as tracked changes while the document's review.tracking is on.
export const REVIEW_TRANSFORMER: DocxTransactionTransformer | null = trackTransaction
