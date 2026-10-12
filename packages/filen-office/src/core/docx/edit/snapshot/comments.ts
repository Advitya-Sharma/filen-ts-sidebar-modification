import type { SnapshotInput } from "./index"

// The comment whose anchor holds the selection.

export interface DocxCommentContext {
	readonly id: string
	readonly done: boolean
}

export function commentOf(_input: SnapshotInput): DocxCommentContext | null {
	return null
}
