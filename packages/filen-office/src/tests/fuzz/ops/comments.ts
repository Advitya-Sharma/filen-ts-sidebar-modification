import type { Node } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { markCommand, markInternal } from "../../../core/docx/edit/invariants/index"
import {
	commentRecords,
	docxComments,
	type CommentHost,
	type DocxComment,
	type DocxCommentRecord
} from "../../../core/docx/review/comments"
import { docxSchema, type DocxEditorSession, type PagedEditor } from "../harness"
import { inlinePosition, nearPosition } from "./base"
import type { EditOp, OpContext, Rand, ScriptTarget } from "./index"

// Comments: add, reply, edit, resolve, reopen, delete and step through them under the paged editor (whose
// pages must stay those of the saved file); comment records and anchors made raw for the writer and Gate 2;
// and a resolve and a delete in the scripted suite.

const nodes = docxSchema.nodes
const AUTHOR = "Fuzz Author"
const TEXTS = ["Looks good", "Why?\nSecond paragraph", "Tab\there", "Kraków — naïve", "", "x"]
// Parts a comment edit may write besides the main part.
const COMMENT_WRITES = ["comments", "commentsExtended", "commentsIds", "commentsExtensible", "people"]

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

// The paged editor with an author for what records one (the app asks the user; the fuzz names one).
function editorOf(context: OpContext): PagedEditor | undefined {
	const editor = context.editor

	if (editor !== undefined) editor.body.session.context.identity.author ??= AUTHOR

	return editor
}

function comments(editor: PagedEditor): DocxComment[] {
	return docxComments(editor)
		.threads()
		.flatMap(thread => [thread.comment, ...thread.replies])
}

function commentOp(name: string, weight: number, run: (context: OpContext, editor: PagedEditor) => string): EditOp {
	return {
		name,
		harness: "paged",
		weight,
		run(context) {
			const editor = editorOf(context)

			return editor === undefined ? "noop" : run(context, editor)
		}
	}
}

function onComment(
	name: string,
	can: (comment: DocxComment) => boolean,
	command: (comment: DocxComment, rand: Rand) => Parameters<OpContext["execute"]>[1]
): EditOp {
	return commentOp(name, 1, (context, editor) => {
		const candidates = comments(editor).filter(can)

		if (candidates.length === 0) return "noop"

		const comment = pick(context.rand, candidates)
		const executed = command(comment, context.rand)

		return `${name} ${comment.id} -> ${context.execute(name, executed)}`
	})
}

// ---- damage: records and anchors made raw ----

function sessionHost(session: DocxEditorSession): CommentHost {
	return { body: { session }, sessions: [session], storySession: story => (story === session.story.story.index ? session : undefined) }
}

function sourceIds(session: DocxEditorSession): string[] {
	return session.context.model.comments.map(comment => comment.id)
}

function recordDamage(context: OpContext): string {
	const rand = context.rand
	const session = context.session
	const ids = [...sourceIds(session), "999", String(Math.floor(rand() * 5))]
	const id = pick(rand, ids)
	const record: DocxCommentRecord = pick(rand, [{ done: true }, { done: false }, { deleted: true }, { parent: pick(rand, ids) }, {}])
	const records = { ...commentRecords(session.state.doc), [id]: record }

	return `commentRecord ${id} ${JSON.stringify(record)} -> ${context.transact("commentRecord", tr => markCommand(tr.setDocAttribute("comments", records)))}`
}

function anchorDamage(context: OpContext): string {
	const rand = context.rand
	const session = context.session
	const doc = session.state.doc
	const at = inlinePosition(rand, doc)

	if (at === undefined) return "noop"

	const id = pick(rand, [...sourceIds(session), "777", String(Math.floor(rand() * 4))])
	const kind = pick(rand, ["commentRangeStart", "commentRangeEnd", "reference"] as const)
	const node: Node = kind === "reference" ? nodes.comment_ref.create({ id }) : nodes.marker.create({ kind, id })
	const internal = rand() < 0.5

	return `commentAnchor ${kind} ${id}@${at} -> ${context.transact("commentAnchor", tr => {
		tr.insert(at, node)

		return internal ? markInternal(tr) : markCommand(tr)
	})}`
}

// Resolves, reopens or deletes a source comment anchored in the session's story, as the commands do.
function storeDamage(context: OpContext): string {
	const rand = context.rand
	const store = docxComments(sessionHost(context.session))
	const threads = store.threads()

	if (threads.length === 0) return "noop"

	const thread = pick(rand, threads)
	const action = pick(rand, ["resolve", "reopen", "delete"] as const)
	const done = action === "delete" ? store.remove(thread.id) : store.resolve(thread.id, action === "resolve")

	return `commentStore ${action} ${thread.id} -> ${String(done)}`
}

// ---- scripted suite: on the body's own comments ----

function scriptHost(target: ScriptTarget): CommentHost {
	const body = target.session

	return {
		body: { session: body },
		get sessions() {
			const comment = target.story("comment")

			return comment === undefined ? [body] : [body, comment]
		},
		storySession(story) {
			if (story === body.story.story.index) return body

			const comment = target.story("comment")

			return comment?.story.story.index === story ? comment : undefined
		}
	}
}

function bodyThreads(target: ScriptTarget) {
	const host = scriptHost(target)
	const store = docxComments(host)
	const body = target.session.story.story.index

	return { store, threads: store.threads().filter(thread => thread.comment.story === body) }
}

// Within one paragraph: the scripted suite allows one changed block per edit.
function inOneParagraph(session: DocxEditorSession, from: number, to: number): boolean {
	const doc = session.state.doc

	return doc.resolve(from).sameParent(doc.resolve(Math.min(to, doc.content.size)))
}

export const commentOps: readonly EditOp[] = [
	commentOp("commentAdd", 3, (context, editor) => {
		const session = editor.session
		const doc = session.state.doc
		const a = inlinePosition(context.rand, doc)

		if (a === undefined) return "noop"

		const b = context.rand() < 0.3 ? a : (nearPosition(context.rand, doc, a) ?? a)

		session.dispatch(session.state.tr.setSelection(TextSelection.between(doc.resolve(Math.min(a, b)), doc.resolve(Math.max(a, b)))))

		return `commentAdd ${String(Math.min(a, b))}..${String(Math.max(a, b))} -> ${context.execute("comment.add", { type: "comment.add", text: pick(context.rand, TEXTS) })}`
	}),
	onComment(
		"commentReply",
		comment => comment.canReply,
		(comment, rand) => ({ type: "comment.reply", id: comment.id, text: pick(rand, TEXTS) })
	),
	onComment(
		"commentEdit",
		comment => comment.canEdit,
		(comment, rand) => ({ type: "comment.edit", id: comment.id, text: pick(rand, TEXTS) })
	),
	onComment(
		"commentResolve",
		comment => comment.canResolve && !comment.done,
		comment => ({ type: "comment.resolve", id: comment.id })
	),
	onComment(
		"commentReopen",
		comment => comment.canResolve && comment.done,
		comment => ({ type: "comment.reopen", id: comment.id })
	),
	onComment(
		"commentDelete",
		comment => comment.canDelete,
		comment => ({ type: "comment.delete", id: comment.id })
	),
	commentOp("commentStep", 1, context => {
		const type = context.rand() < 0.5 ? "comment.next" : "comment.prev"

		return `${type} -> ${context.execute(type, { type })}`
	}),
	commentOp("commentUndo", 1, context => `comment undo -> ${context.execute("history.undo", { type: "history.undo" })}`),
	{ name: "commentRecord", harness: "damage", weight: 1, run: recordDamage },
	{ name: "commentAnchor", harness: "damage", weight: 1, run: anchorDamage },
	{ name: "commentStore", harness: "damage", weight: 2, run: storeDamage },
	{
		name: "commentResolveScripted",
		harness: "damage",
		weight: 0,
		run: storeDamage,
		writes: COMMENT_WRITES,
		script: {
			order: 600,
			run(target) {
				const { store, threads } = bodyThreads(target)
				const thread = threads.find(candidate => candidate.comment.canResolve && !candidate.done)

				return thread !== undefined && store.resolve(thread.id, true)
			}
		}
	},
	{
		name: "commentDeleteScripted",
		harness: "damage",
		weight: 0,
		run: storeDamage,
		writes: COMMENT_WRITES,
		script: {
			order: 610,
			run(target) {
				const { store, threads } = bodyThreads(target)
				const thread = [...threads].reverse().find(candidate => {
					const anchor = store.anchor(candidate.id)

					return (
						candidate.comment.canDelete &&
						candidate.replies.length === 0 &&
						anchor !== undefined &&
						inOneParagraph(target.session, anchor.from, Math.max(anchor.to, anchor.reference))
					)
				})

				return thread !== undefined && store.remove(thread.id)
			}
		}
	}
]
