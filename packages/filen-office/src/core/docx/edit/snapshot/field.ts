import type { Node } from "prosemirror-model"
import { NodeSelection } from "prosemirror-state"
import { parseFieldCode, updatableField } from "../../fields/evaluate"
import { sourceTocOf } from "../../fields/toc"
import { docxSchema } from "../schema"
import type { SnapshotInput } from "./index"

// The field the selection is on: a field selected or at the caret, or a table of contents (a locked block of
// one of the file's selected, or the paragraph holding the characters of one the editor made). Read from the
// selection's own nodes only; the same object while the field is.

export interface DocxFieldContext {
	// The field type (PAGE, REF, TOC...).
	readonly type: string
	// An update recomputes its result (Update Field), or rebuilds the table (Update Table).
	readonly updatable: boolean
	readonly toc: boolean
	readonly instruction: string
}

const nodes = docxSchema.nodes
const CONTEXTS = new WeakMap<Node, DocxFieldContext>()

function contextOf(node: Node, make: () => DocxFieldContext): DocxFieldContext {
	let context = CONTEXTS.get(node)

	if (context === undefined) {
		context = Object.freeze(make())
		CONTEXTS.set(node, context)
	}

	return context
}

export function fieldOf(input: SnapshotInput): DocxFieldContext | null {
	const session = input.session
	const selection = input.state.selection
	const $from = selection.$from
	let field: Node | null = selection instanceof NodeSelection ? selection.node : null

	if (field === null && selection.empty) field = $from.nodeAfter?.type === nodes.field ? $from.nodeAfter : $from.nodeBefore

	if (field?.type === nodes.field) {
		const node = field

		return contextOf(node, () => ({
			type: node.attrs["type"] as string,
			updatable: updatableField(session, node),
			toc: false,
			instruction: node.attrs["instruction"] as string
		}))
	}

	const after = $from.nodeAfter

	if (after?.type === nodes.locked_block && after.attrs["reason"] === "fieldSpansBlocks") {
		const toc = sourceTocOf(session.story.part, after.attrs["src"] as number)

		return toc === undefined
			? null
			: contextOf(after, () => ({ type: "TOC", updatable: toc.updatable, toc: true, instruction: toc.instruction }))
	}

	let begin: Node | undefined

	$from.parent.forEach(child => {
		if (begin === undefined && child.type === nodes.field_char && child.attrs["kind"] === "begin") begin = child
	})

	if (begin === undefined) return null

	const node = begin
	const instruction = node.attrs["instruction"] as string
	const type = parseFieldCode(instruction).type

	return contextOf(node, () => ({ type, updatable: type === "TOC", toc: type === "TOC", instruction }))
}
