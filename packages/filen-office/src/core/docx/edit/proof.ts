import type { Node } from "prosemirror-model"
import type { EditorState, Transaction } from "prosemirror-state"
import type { StoryPart } from "../model/index"
import type { DocxDocBuilder, StoryDoc } from "./build"
import { markInternal } from "./invariants"
import { NO_SOURCE, docxSchema, type LockReason } from "./schema"

// Gate 1: every source block must come out of the edit path as it went in. A proof forces a block through
// the writer and compares the result with the block's source at XML-meaning level; a block that fails is
// locked (kept as its bytes, read-only) and the rest of the document stays editable.
export interface BlockProof {
	// Undefined when `node` (a source paragraph or table of `part`) round-trips; otherwise why it cannot.
	check(node: Node, part: StoryPart): LockReason | undefined
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
	const out: LockRequest[] = []

	doc.nodesBetween(from, to, (node, pos) => {
		if (node.isInline || node.type === nodes.locked_block) return false

		const source = (node.attrs["src"] as number | undefined) ?? NO_SOURCE

		if (node.type === nodes.paragraph || node.type === nodes.table) {
			if (source === NO_SOURCE || (node.type === nodes.paragraph && !story.pristine.has(node))) return false

			const reason = proof.check(node, story.part)

			if (reason !== undefined) {
				out.push({ pos, reason })

				return false
			}

			return node.type === nodes.table
		}

		return true
	})

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
