import type { Node } from "prosemirror-model"
import { NodeSelection } from "prosemirror-state"
import { parseFieldCode, updatableField } from "../../fields/evaluate"
import { sourceTocOf, tocOptions, type TocRefusal } from "../../fields/toc"
import { docxSchema } from "../schema"
import type { SnapshotInput } from "./index"

// The field the selection is on: a field selected or at the caret, or a table of contents (a locked block of
// one of the file's selected, or the paragraph holding the characters of one built as paragraphs). Read from the
// selection's own nodes only; the same object while the field is.

// Why a table of contents cannot be updated: see TocRefusal; "trackedLockedTable": one of the file's kept as
// locked blocks while changes are tracked (its old entries cannot be recorded deleted: update it with tracking
// off).
export type DocxFieldRefusal = TocRefusal | "trackedLockedTable"

export interface DocxFieldContext {
	// The field type (PAGE, REF, TOC...).
	readonly type: string
	// An update recomputes its result (Update Field), or rebuilds the table (Update Table).
	readonly updatable: boolean
	readonly toc: boolean
	readonly instruction: string
	// Why a table of contents cannot be updated, null when it can (or for other fields).
	readonly refusal: DocxFieldRefusal | null
}

const nodes = docxSchema.nodes
// Per node, the context with changes not tracked and tracked.
const CONTEXTS = new WeakMap<Node, [DocxFieldContext | undefined, DocxFieldContext | undefined]>()

function contextOf(node: Node, tracking: boolean, make: () => DocxFieldContext): DocxFieldContext {
	let pair = CONTEXTS.get(node)

	if (pair === undefined) {
		pair = [undefined, undefined]
		CONTEXTS.set(node, pair)
	}

	const index = tracking ? 1 : 0
	let context = pair[index]

	if (context === undefined) {
		context = Object.freeze(make())
		pair[index] = context
	}

	return context
}

export function fieldOf(input: SnapshotInput): DocxFieldContext | null {
	const session = input.session
	const tracking = input.context.review.tracking
	const selection = input.state.selection
	const $from = selection.$from
	let field: Node | null = selection instanceof NodeSelection ? selection.node : null

	if (field === null && selection.empty) field = $from.nodeAfter?.type === nodes.field ? $from.nodeAfter : $from.nodeBefore

	if (field?.type === nodes.field) {
		const node = field

		return contextOf(node, false, () => ({
			type: node.attrs["type"] as string,
			updatable: updatableField(session, node),
			toc: false,
			instruction: node.attrs["instruction"] as string,
			refusal: null
		}))
	}

	const after = $from.nodeAfter

	if (after?.type === nodes.locked_block && after.attrs["reason"] === "fieldSpansBlocks") {
		const toc = sourceTocOf(session.story.part, after.attrs["src"] as number)

		if (toc === undefined) return null

		const refusal = toc.refusal ?? (tracking ? "trackedLockedTable" : null)

		return contextOf(after, tracking, () => ({
			type: "TOC",
			updatable: refusal === null,
			toc: true,
			instruction: toc.instruction,
			refusal
		}))
	}

	let begin: Node | undefined

	$from.parent.forEach(child => {
		if (begin === undefined && child.type === nodes.field_char && child.attrs["kind"] === "begin") begin = child
	})

	if (begin === undefined) return null

	const node = begin
	const instruction = node.attrs["instruction"] as string
	const type = parseFieldCode(instruction).type

	const toc = type === "TOC"
	const refusal = toc && !tocOptions(parseFieldCode(instruction)).supported ? "unsupportedTable" : null

	return contextOf(node, false, () => ({ type, updatable: toc && refusal === null, toc, instruction, refusal }))
}
