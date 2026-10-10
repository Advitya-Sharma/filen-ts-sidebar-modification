export type StoryKind = "body" | "header" | "footer" | "footnote" | "endnote" | "comment" | "textbox"

export type NoteType = "normal" | "separator" | "continuationSeparator" | "continuationNotice"

// One flow of text laid out on its own: its content is the children of `item` in part `part`.
export interface Story {
	readonly index: number
	readonly kind: StoryKind
	readonly part: number
	readonly item: number
	// w:id of a note or comment.
	readonly id: string | undefined
	readonly noteType: NoteType | undefined
	// The story a text box sits in, and the drawing or VML item that holds it; -1 for top-level stories.
	readonly parent: number
	readonly owner: number
}

const NOTE_TYPES = new Set<string>(["normal", "separator", "continuationSeparator", "continuationNotice"])

export function noteType(value: string | undefined): NoteType {
	return value !== undefined && NOTE_TYPES.has(value) ? (value as NoteType) : "normal"
}
