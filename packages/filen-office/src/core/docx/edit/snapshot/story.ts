import type { StoryKind } from "../../model/index"
import { boundEditor, headerContext, type DocxHeaderContext } from "../../stories/headers"
import type { SnapshotInput } from "./index"

// The story being edited: for a header or footer, where the page the caret is on shows it (what the floating
// "Header — Close" chip shows and toggles); for a note, which one.

export interface DocxStoryContext {
	readonly kind: StoryKind
	// Its index among the document's stories (what PagedEditor.enterStory takes).
	readonly index: number
	// Always set by the snapshot (null for other stories); optional for contexts an app builds itself.
	readonly header?: DocxHeaderContext | null
	readonly note?: { readonly kind: "footnote" | "endnote"; readonly id: string } | null
}

export type { DocxHeaderContext }

const STORIES = new WeakMap<object, { key: string; value: DocxStoryContext }>()

function keyOf(header: DocxHeaderContext | null): string {
	if (header === null) return ""

	return [
		header.kind,
		header.section,
		header.sections,
		header.type,
		header.differentFirst,
		header.differentOddEven,
		header.linkedToPrevious,
		header.canLink
	].join(",")
}

// One object per story and placement, so snapshots compare it by identity.
export function storyOf(input: SnapshotInput): DocxStoryContext {
	const story = input.session.story.story
	const editor = story.kind === "header" || story.kind === "footer" ? boundEditor(input.context) : undefined
	const header = editor !== undefined && editor.session === input.session ? headerContext(editor) : null
	const key = keyOf(header)
	const known = STORIES.get(story)

	if (known !== undefined && known.key === key) return known.value

	const value: DocxStoryContext = Object.freeze({
		kind: story.kind,
		index: story.index,
		header,
		note:
			(story.kind === "footnote" || story.kind === "endnote") && story.id !== undefined
				? Object.freeze({ kind: story.kind, id: story.id })
				: null
	})

	STORIES.set(story, { key, value })

	return value
}
