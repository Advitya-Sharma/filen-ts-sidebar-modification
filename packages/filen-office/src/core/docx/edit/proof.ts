import type { Node } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import type { DocxDocBuilder, StoryDoc } from "./build"
import { markInternal } from "./invariants/index"
import { NO_SOURCE, docxSchema, type LockReason } from "./schema"

// Gate 1: every source block must come out of the edit path as it went in. A proof forces a block through
// the writer and compares the result with the block's source at XML-meaning level; a block that fails is
// locked (kept as its bytes, read-only) and the rest of the document stays editable.
export interface BlockProof {
	// Undefined when `node` (a source paragraph or table of `story`) round-trips; otherwise why it cannot.
	check(node: Node, story: StoryDoc): LockReason | undefined
	// check() yielding whenever `due` says so (inside a long table), for a host that proves in slices.
	steps?(node: Node, story: StoryDoc, due: () => boolean): Generator<void, LockReason | undefined>
}

export interface LockRequest {
	readonly pos: number
	readonly reason: LockReason
}

const nodes = docxSchema.nodes

// Content the model does not read (an element it kept opaque) may depend on its neighbours in ways an
// edited paragraph cannot reproduce; until the writer's proof covers a block, such blocks are locked.
export const conservativeProof: BlockProof = {
	check(node) {
		if (node.type !== nodes.paragraph) return undefined

		let unmodelled = false

		node.forEach(child => {
			if (
				!unmodelled &&
				(child.type === nodes.opaque_inline || child.type === nodes.run_object) &&
				child.attrs["kind"] === "Opaque"
			) {
				unmodelled = true
			}
		})

		return unmodelled ? "unmodelled" : undefined
	}
}

// Proves the source blocks in [from, to) of a story document: paragraphs, and tables as a whole before
// their cells' content (a failed table is locked whole). Only blocks still pristine are checked.
export function proveBlocks(story: StoryDoc, doc: Node, proof: BlockProof, from: number = 0, to: number = doc.content.size): LockRequest[] {
	const steps = proveSteps(story, doc, proof, from, to, undefined)

	for (;;) {
		const step = steps.next()

		if (step.done === true) return step.value
	}
}

// proveBlocks yielding whenever `due` says so: between blocks, and inside one where the proof checks in steps.
export function proveBlocksSteps(
	story: StoryDoc,
	doc: Node,
	proof: BlockProof,
	due: () => boolean,
	from: number = 0,
	to: number = doc.content.size
): Generator<void, LockRequest[]> {
	return proveSteps(story, doc, proof, from, to, due)
}

interface Frame {
	readonly parent: Node
	// The parent's content start, and the range of its content walked.
	readonly start: number
	readonly from: number
	readonly to: number
	index: number
	pos: number
}

// The walk of doc.nodesBetween(from, to), with a stack of the containers open.
function* proveSteps(
	story: StoryDoc,
	doc: Node,
	proof: BlockProof,
	from: number,
	to: number,
	due: (() => boolean) | undefined
): Generator<void, LockRequest[]> {
	const out: LockRequest[] = []
	const stack: Frame[] = [{ parent: doc, start: 0, from, to, index: 0, pos: 0 }]

	while (stack.length > 0) {
		const frame = stack[stack.length - 1] as Frame

		if (frame.pos >= frame.to || frame.index >= frame.parent.childCount) {
			stack.pop()

			continue
		}

		const node = frame.parent.child(frame.index++)
		const pos = frame.pos

		frame.pos += node.nodeSize

		if (frame.pos <= frame.from || node.isInline || node.type === nodes.locked_block) continue

		const at = frame.start + pos

		if (node.type === nodes.paragraph || node.type === nodes.table) {
			const source = (node.attrs["src"] as number | undefined) ?? NO_SOURCE

			if (source === NO_SOURCE || (node.type === nodes.paragraph && !story.pristine.has(node))) continue

			const reason = due !== undefined && proof.steps !== undefined ? yield* proof.steps(node, story, due) : proof.check(node, story)

			if (due?.() === true) yield

			if (reason !== undefined) {
				out.push({ pos: at, reason })

				continue
			}

			if (node.type !== nodes.table) continue
		}

		if (node.content.size > 0) {
			stack.push({
				parent: node,
				start: at + 1,
				from: Math.max(0, frame.from - pos - 1),
				to: Math.min(node.content.size, frame.to - pos - 1),
				index: 0,
				pos: 0
			})
		}
	}

	return out
}

// Replaces the requested blocks with locked blocks, outside the undo history. A locked block writes the
// same bytes as the block it replaces, so the containers around it stay copyable (see StoryDoc).
export function lockBlocks(
	state: EditorState,
	story: StoryDoc,
	builder: DocxDocBuilder,
	requests: readonly LockRequest[]
): Transaction | null {
	if (requests.length === 0) return null

	const tr = state.tr

	for (const request of [...requests].sort((a, b) => b.pos - a.pos)) {
		const node = tr.doc.nodeAt(request.pos)

		if (node === null || typeof node.attrs["src"] !== "number" || node.attrs["src"] === NO_SOURCE) continue

		tr.replaceWith(request.pos, request.pos + node.nodeSize, builder.lockedBlock(story, node, request.reason))
	}

	if (!tr.docChanged) return null

	return markInternal(tr).setMeta("addToHistory", false)
}
