import { lastSkipped, movedChange, reviewIndexWanted, storyReviewIndex } from "../../review/accept"
import type { DocxEditorSession } from "../session"
import type { SnapshotInput } from "./index"

// The review state: whether edits are tracked, whether recording one waits for the author's name, and how
// many changes and comments the story has (from its ReviewIndex, never a document scan).

export interface DocxReviewSnapshot {
	readonly tracking: boolean
	readonly authorNeeded: boolean
	readonly revisionCount: number
	readonly commentCount: number
	// Changes the story's last accept or reject left as they were: inside locked blocks, cell insertions and
	// deletions, paragraph marks that cannot join (a section break, no paragraph after).
	readonly skipped: number
	// The id of the tracked change at the selection (what review.accept / review.reject with `id` resolve, and
	// what review.next / review.prev move to), or null.
	readonly current: string | null
}

export const NO_REVIEW: DocxReviewSnapshot = Object.freeze({
	tracking: false,
	authorNeeded: false,
	revisionCount: 0,
	commentCount: 0,
	skipped: 0,
	current: null
})

const last = new WeakMap<DocxEditorSession, DocxReviewSnapshot>()

export function reviewOf(input: SnapshotInput): DocxReviewSnapshot {
	const session = input.session
	const tracking = input.context.review.tracking
	const authorNeeded = tracking && !input.readOnly && (input.context.identity.author ?? "") === ""
	const index = reviewIndexWanted(session) ? storyReviewIndex(session) : undefined
	const revisionCount = index?.revisionCount ?? 0
	const commentCount = index?.commentCount ?? 0
	const skipped = lastSkipped(session)
	const current =
		index === undefined || revisionCount === 0
			? null
			: (movedChange(session, input.state) ?? index.changeAt(input.state.selection.from)?.id ?? null)
	const previous = last.get(session) ?? NO_REVIEW

	if (
		previous.tracking === tracking &&
		previous.authorNeeded === authorNeeded &&
		previous.revisionCount === revisionCount &&
		previous.commentCount === commentCount &&
		previous.skipped === skipped &&
		previous.current === current
	)
		return previous

	const snapshot = Object.freeze({ tracking, authorNeeded, revisionCount, commentCount, skipped, current })

	last.set(session, snapshot)

	return snapshot
}
