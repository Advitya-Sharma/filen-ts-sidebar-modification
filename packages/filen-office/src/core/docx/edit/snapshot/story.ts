import type { StoryKind } from "../../model/index"
import type { SnapshotInput } from "./index"

// The story being edited.

export interface DocxStoryContext {
	readonly kind: StoryKind
	// Its index among the document's stories (what PagedEditor.enterStory takes).
	readonly index: number
}

const STORIES = new WeakMap<object, DocxStoryContext>()

// One object per story, so snapshots compare it by identity.
export function storyOf(input: SnapshotInput): DocxStoryContext {
	const story = input.session.story.story
	let out = STORIES.get(story)

	if (out === undefined) {
		out = Object.freeze({ kind: story.kind, index: story.index })
		STORIES.set(story, out)
	}

	return out
}
