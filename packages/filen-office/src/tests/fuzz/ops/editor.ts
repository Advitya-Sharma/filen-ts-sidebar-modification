import {
	addStory,
	enterStory,
	exitStory,
	setView,
	storyIndex,
	toggleSuspended,
	type DocxEditorSession,
	type NewStoryKind
} from "../harness"
import type { EditOp, Rand } from "./index"

// The paged editor's own features: switching the revision view in place, suspending layout while another
// surface shows the document, and new (still unreferenced) notes, comments, headers and footers.

const VIEWS = ["markup", "original", "final"] as const
const NEW_STORIES: readonly NewStoryKind[] = ["footnote", "endnote", "comment", "header", "footer"]
const WORDS = ["Filen", "note", "header", "footer", "comment", "x"]

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

function typeInto(session: DocxEditorSession, rand: Rand): void {
	let at: number | undefined

	session.state.doc.descendants((node, pos) => {
		if (at === undefined && node.isTextblock) at = pos + 1

		return at === undefined
	})

	if (at !== undefined) session.dispatch(session.state.tr.insertText(`${pick(rand, WORDS)} ${pick(rand, WORDS)} `, at))
}

export const editorOps: readonly EditOp[] = [
	{
		name: "view",
		harness: "paged",
		run(context) {
			if (context.editor === undefined) return "noop"

			const view = pick(context.rand, VIEWS)

			return `view ${view} ${String(setView(context.editor, view))}`
		}
	},
	{
		name: "suspend",
		harness: "paged",
		run(context) {
			if (context.editor === undefined) return "noop"

			return toggleSuspended(context.editor) ? "suspend" : "resume"
		}
	},
	// Nothing refers to the new story, so the save leaves it out and no page changes.
	{
		name: "addStory",
		harness: "paged",
		run(context) {
			const editor = context.editor

			if (editor === undefined) return "noop"

			const kind = pick(context.rand, NEW_STORIES)
			const session = addStory(editor, kind)

			if (session === undefined) return `addStory ${kind} refused`

			if (kind !== "comment" && enterStory(editor, storyIndex(session))) {
				typeInto(editor.session, context.rand)
				exitStory(editor)

				return `addStory ${kind}#${storyIndex(session)} entered`
			}

			typeInto(session, context.rand)

			return `addStory ${kind}#${storyIndex(session)}`
		}
	}
]
