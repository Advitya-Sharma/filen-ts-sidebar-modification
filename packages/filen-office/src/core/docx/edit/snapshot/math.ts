import type { Node } from "prosemirror-model"
import { NodeSelection } from "prosemirror-state"
import { mathNodeLatex } from "../../math/edit"
import { docxSchema } from "../schema"
import type { DocxEditorSession } from "../session"
import type { SnapshotInput } from "./index"

// The equation the selection is (a node selection of it), as LaTeX. The same object while the selected math
// node and its position are; the LaTeX is read once per node.

export interface DocxEquationContext {
	readonly latex: string
	// Editing it as LaTeX loses formatting the equation has.
	readonly lossy: boolean
	// On a line of its own (m:oMathPara).
	readonly display: boolean
	// The equation's position in the story's document.
	readonly pos: number
}

const previous = new WeakMap<DocxEditorSession, { readonly node: Node; readonly value: DocxEquationContext }>()

export function equationOf(input: SnapshotInput): DocxEquationContext | null {
	const selection = input.state.selection

	if (!(selection instanceof NodeSelection) || selection.node.type !== docxSchema.nodes.math) return null

	const node = selection.node
	const cached = previous.get(input.session)

	if (cached?.node === node && cached.value.pos === selection.from) return cached.value

	const value: DocxEquationContext = Object.freeze({ ...mathNodeLatex(node, input.session.story.part), pos: selection.from })

	previous.set(input.session, { node, value })

	return value
}
