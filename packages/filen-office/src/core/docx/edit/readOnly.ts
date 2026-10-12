import type { Node } from "prosemirror-model"
import type { Story, StoryPart } from "../model/index"
import type { DocxGateReason } from "../package"
import { textboxCopies, type TextboxProblem } from "../write/nested"
import type { DocxEditContext } from "./context"
import { docxSchema } from "./schema"

// Text boxes are edited where the model reads them and written into the copy Word keeps for older readers
// (the VML in mc:Fallback, see write/nested.ts). "linkedTextbox": its text flows on into other boxes;
// "textboxFallback": its copies cannot be matched up; "lockedTextbox": the block holding it is locked (kept as
// its bytes, which a text box rewritten inside would change). All stay view-only.
export type DocxReadOnlyReason = DocxGateReason | "host" | TextboxProblem | "lockedTextbox"

// Text boxes found in a locked block of the story holding them (see textboxHolderLocked).
const lockedTextboxes = new WeakSet<Story>()

// Whether text box `story` stands in a locked block of `holder` (the document of the story holding it, of
// part `part`), noted for readOnlyReasons. Ask before opening the text box's session.
export function textboxHolderLocked(holder: Node, part: StoryPart, story: Story): boolean {
	const owner = story.owner
	let locked = false
	const visit = (node: Node): boolean => {
		if (locked || node.isInline) return false

		const own = node.attrs["src"]

		// Only blocks around the text box's drawing are entered (its source item lies inside theirs); a locked
		// block always has a source.
		if (typeof own !== "number" || own < 0) return !node.isTextblock
		if (owner < own || owner >= part.end(own)) return false

		if (node.type === docxSchema.nodes.locked_block) locked = true

		return !locked && !node.isTextblock
	}

	holder.descendants(visit)

	if (locked) lockedTextboxes.add(story)

	return locked
}

// Why a session of `story` is read-only, in order: the document's gate reasons, the host's mode, the story's
// own. Empty when it may be edited.
export function readOnlyReasons(context: DocxEditContext, story: Story, host: boolean): readonly DocxReadOnlyReason[] {
	const out: DocxReadOnlyReason[] = [...context.gate]

	if (host) out.push("host")

	if (story.kind === "textbox") {
		const part = context.model.parts[story.part]
		const copies = part === undefined ? "textboxFallback" : textboxCopies(part, part.node(story.item))

		if (typeof copies === "string") out.push(copies)
		if (lockedTextboxes.has(story)) out.push("lockedTextbox")
	}

	return out
}
