import type { CommentInfo, DocxModel } from "../../model/index"
import { commentRecords, commentsAt } from "../../review/comments"
import type { DocxEditorSession } from "../session"
import type { SnapshotInput } from "./index"

// The comment whose anchor holds the selection's head: what a margin balloon highlights and what the
// comment commands act on. Found from the story's review index (no document scan), the same object while
// the answer is the same.

export interface DocxCommentContext {
	readonly id: string
	// The thread it belongs to (its first comment's id).
	readonly thread: string
	// Its thread is resolved.
	readonly done: boolean
	// Every comment whose range holds the caret, innermost first.
	readonly ids: readonly string[]
}

interface Cached {
	readonly records: unknown
	readonly head: number
	readonly version: number
	readonly result: DocxCommentContext | null
}

const CACHE = new WeakMap<DocxEditorSession, Cached>()
const SOURCES = new WeakMap<DocxModel, ReadonlyMap<string, CommentInfo>>()

function sourceComments(model: DocxModel): ReadonlyMap<string, CommentInfo> {
	let out = SOURCES.get(model)

	if (out === undefined) {
		out = new Map(model.comments.map(comment => [comment.id, comment]))
		SOURCES.set(model, out)
	}

	return out
}

function same(a: DocxCommentContext | null | undefined, id: string, thread: string, done: boolean, ids: readonly string[]): boolean {
	return (
		a !== null &&
		a !== undefined &&
		a.id === id &&
		a.thread === thread &&
		a.done === done &&
		a.ids.length === ids.length &&
		a.ids.every((value, k) => value === ids[k])
	)
}

export function commentOf(input: SnapshotInput): DocxCommentContext | null {
	const session = input.session
	const model = input.context.model

	// Neither the document nor the editor has a comment: nothing to look for (and no index to build).
	if (model.comments.length === 0 && ![...input.context.newStories.values()].some(story => story.kind === "comment")) return null

	const index = session.reviewIndex
	const head = input.state.selection.head
	const records = input.state.doc.attrs["comments"]
	const cached = CACHE.get(session)

	if (cached !== undefined && cached.head === head && cached.version === index.version && cached.records === records) return cached.result

	const ids = commentsAt(index, head)
	const id = ids[0]
	let result: DocxCommentContext | null = null

	if (id !== undefined) {
		const own = commentRecords(input.state.doc)
		const sources = sourceComments(model)
		let thread = id

		for (const seen = new Set([id]); ;) {
			const parent = own[thread]?.parent ?? sources.get(thread)?.parentId

			if (parent === undefined || seen.has(parent)) break

			seen.add(parent)
			thread = parent
		}

		const done = own[thread]?.done ?? sources.get(thread)?.done ?? false

		result = same(cached?.result, id, thread, done, ids)
			? (cached?.result ?? null)
			: Object.freeze({ id, thread, done, ids: Object.freeze([...ids]) })
	}

	CACHE.set(session, { records, head, version: index.version, result })

	return result
}
