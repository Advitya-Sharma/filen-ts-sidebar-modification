import type { SnapshotInput } from "./index"

// The review state: whether edits are tracked, whether recording one waits for the author's name, and how
// many changes and comments the story has (from its ReviewIndex, never a document scan).

export interface DocxReviewSnapshot {
	readonly tracking: boolean
	readonly authorNeeded: boolean
	readonly revisionCount: number
	readonly commentCount: number
}

export const NO_REVIEW: DocxReviewSnapshot = Object.freeze({ tracking: false, authorNeeded: false, revisionCount: 0, commentCount: 0 })

export function reviewOf(_input: SnapshotInput): DocxReviewSnapshot {
	return NO_REVIEW
}
