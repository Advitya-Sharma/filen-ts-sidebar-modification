import type { Node } from "prosemirror-model"
import type { StoryDoc } from "../../docx/edit/build"
import type { BlockProof } from "../../docx/edit/proof"
import { NO_SOURCE, type LockReason } from "../../docx/edit/schema"

// Gate 1 for an RTF document in the editor: the docx writer's proof (the block survives the edit path as
// WordprocessingML) and the RTF writer's (proveRtf: the block written as RTF reads back as it was). A top-level
// block either refuses is locked; the rest of the document stays editable.

export interface RtfLockedBlock {
	// The translation's part and the block element's offset in it.
	readonly part: string
	readonly offset: number
}

export function rtfBlockProof(locked: readonly RtfLockedBlock[], docx?: BlockProof): BlockProof {
	const set = new Set(locked.map(block => `${block.part}#${block.offset}`))
	const rtf = (node: Node, story: StoryDoc): LockReason | undefined => {
		const src = node.attrs["src"] as number

		if (src === NO_SOURCE || set.size === 0) return undefined

		const part = story.part
		const name = part.name.replace(/^\//u, "")

		return set.has(`${name}#${part.document.start(part.node(src))}`) ? "unmodelled" : undefined
	}

	return {
		check(node, story) {
			return rtf(node, story) ?? docx?.check(node, story)
		},
		...(docx?.steps === undefined
			? {}
			: {
					*steps(node: Node, story: StoryDoc, due: () => boolean): Generator<void, LockReason | undefined> {
						return rtf(node, story) ?? (yield* (docx.steps as NonNullable<BlockProof["steps"]>)(node, story, due))
					}
				})
	}
}
