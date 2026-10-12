import type { StoryDoc } from "../edit/build"
import type { StoryPart } from "../model/index"
import type { Out } from "./markup"
import type { StoryDraft } from "./story"
import { DocxWriteError } from "./values"

// A story container rewritten inside source bytes another story copies (a text box in a body paragraph):
// source bytes [from, to) of the part are written as `bytes` instead.
export interface NestedSplice {
	readonly from: number
	readonly to: number
	readonly bytes: Uint8Array
}

// The rewritten containers inside source bytes [from, to) of `part`, ascending and apart (WriteEnvironment.nested).
export type NestedStories = (part: StoryPart, from: number, to: number) => readonly NestedSplice[] | undefined

// Source bytes [from, to) of a story's part as written, with the rewritten containers inside them spliced
// in. `trusted`: the range is whole source nodes (see Out.copy); with a splice it is new markup to gate 2.
export function copySource(out: Out, part: StoryPart, from: number, to: number, trusted: boolean, nested: NestedStories | undefined): void {
	const bytes = part.document.bytes
	const splices = nested?.(part, from, to)

	if (splices === undefined || splices.length === 0) {
		if (trusted) out.copy(bytes, from, to)
		else out.bytes(bytes, from, to)

		return
	}

	let at = from

	for (const splice of splices) {
		if (splice.from < at || splice.to < splice.from || splice.to > to)
			throw new DocxWriteError("unsupported", `Nested stories of ${part.name} overlap`)

		out.bytes(bytes, at, splice.from)
		out.bytes(splice.bytes, 0, splice.bytes.length)
		at = splice.to
	}

	out.bytes(bytes, at, to)
}

// How a save writes stories that sit inside other stories: the sessions in the order they are written
// (inner containers first), which are written into the copies of their outer story rather than into their
// part, and env.nested fed with each inner story's draft.
export interface NestedPlan {
	readonly order: readonly number[]
	inner(index: number): boolean
	record(index: number, draft: StoryDraft): void
	readonly nested: NestedStories
}

// No story is written inside another yet: text boxes are view-only.
export function planNested(_stories: readonly StoryDoc[]): NestedPlan | undefined {
	return undefined
}
