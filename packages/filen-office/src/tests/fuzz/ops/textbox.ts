import { TextSelection } from "prosemirror-state"
import { enterStory, exitStory, storiesOf, type DocxEditorSession } from "../harness"
import { textblocks } from "./base"
import type { EditOp, OpContext, Rand } from "./index"

// Text boxes: typing into and deleting from the document's text boxes under the paged editor (whose pages
// must stay those of the saved file: the box written into its mc:Choice and its VML fallback alike), and the
// first text box of the body edited in the scripted suite. The damage harness's story sessions (stories
// mode) edit text boxes with the base ops.

const WORDS = ["Boxed", "text", "Kraków", "—", "naïve", "x"]

function pick<T>(rand: Rand, list: readonly T[]): T {
	return list[Math.floor(rand() * list.length)] as T
}

// Enters a random text box that can be edited; false when there is none.
function enterTextbox(context: OpContext): boolean {
	const editor = context.editor

	if (editor === undefined) return false

	const boxes = storiesOf(editor).filter(story => story.kind === "textbox" && story.index < editor.edits.storyCount)

	for (let tries = 0; tries < 3 && boxes.length > 0; tries++) {
		const story = pick(context.rand, boxes)

		if (enterStory(editor, story.index) && !editor.session.readOnly) return true
	}

	if (editor.session.story.story.kind === "textbox") exitStory(editor)

	return false
}

function caret(session: DocxEditorSession, rand: Rand, span: boolean): { from: number; to: number } | undefined {
	const blocks = textblocks(session.state.doc)

	if (blocks.length === 0) return undefined

	const block = pick(rand, blocks)
	const from = block.pos + 1 + Math.floor(rand() * (block.node.content.size + 1))
	const to = span ? Math.min(block.pos + 1 + block.node.content.size, from + 1 + Math.floor(rand() * 6)) : from

	return { from, to }
}

function inTextbox(name: string, weight: number, edit: (context: OpContext, session: DocxEditorSession) => string): EditOp {
	return {
		name,
		harness: "paged",
		weight,
		run(context) {
			const editor = context.editor

			if (editor === undefined || !enterTextbox(context)) return "noop"

			const story = editor.session.story.story.index
			const result = edit(context, editor.session)

			// Sometimes the caret stays in the box for the next op.
			if (context.rand() < 0.6) exitStory(editor)

			return `textbox#${String(story)} ${result}`
		}
	}
}

export const textboxOps: readonly EditOp[] = [
	inTextbox("textboxType", 3, (context, session) => {
		const at = caret(session, context.rand, context.rand() < 0.2)

		if (at === undefined) return "noop"

		const text = context.rand() < 0.15 ? "" : `${pick(context.rand, WORDS)} `
		const tr = session.state.tr.setSelection(TextSelection.create(session.state.doc, at.from, at.to))

		return `type ${JSON.stringify(text)}@${String(at.from)} -> ${String(session.dispatch(text === "" ? tr.deleteSelection() : tr.insertText(text)))}`
	}),
	inTextbox("textboxSplit", 1, (context, session) => {
		const at = caret(session, context.rand, false)

		if (at === undefined) return "noop"

		return `split@${String(at.from)} -> ${String(session.dispatch(session.state.tr.split(at.from)))}`
	}),
	{
		name: "textboxScripted",
		harness: "damage",
		weight: 0,
		run: () => "noop",
		script: {
			order: 720,
			run(target) {
				const story = target.story("textbox")

				if (story === undefined || story.readOnly || story.story.story.parent !== target.session.story.story.index) return false

				const first = textblocks(story.state.doc)[0]

				return first !== undefined && story.dispatch(story.state.tr.insertText("Boxed by the edit suite. ", first.pos + 1))
			}
		}
	}
]
