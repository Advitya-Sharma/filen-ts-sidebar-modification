import { compareStamps, deriveWatchState, isResumablePosition, progressMatchesFile, type ItemRec } from "@filen/shared/videos"
import { isCumulativeStackProgress, locateInStack } from "@/features/videos/lib/player/watchTarget.logic"

// Where playback starts. A valid resume point is taken without asking (the merged state already holds the
// newest position from every device) and announced with a "Resumed at … · Start over" pill; "start" plays
// from the beginning.

export type StartRequest = "resume" | "start"

export const RESUME_PILL_MS = 8000

export interface ResumePlan {
	part: number
	// Seconds within `part`.
	position: number
	// Show the pill (playback continues where it was left, not from the start).
	resumed: boolean
}

export function planStart(params: {
	request: StartRequest
	rec: ItemRec | undefined
	// The first part's file: the progress fingerprint.
	file: { uuid: string; size: number }
	// One per part; null where not probed yet.
	partDurations: readonly (number | null)[]
}): ResumePlan {
	const { request, rec, file, partDurations } = params

	if (request === "start" || rec?.p === undefined) {
		return { part: 0, position: 0, resumed: false }
	}

	const progress = rec.p

	if (partDurations.length > 1) {
		return planStackStart(rec, progress, file, partDurations)
	}

	const state = deriveWatchState({
		rec,
		file: { uuid: file.uuid, size: file.size, durationSec: partDurations[0] }
	})

	return state.inProgress ? { part: 0, position: progress.pos, resumed: progress.pos > 0 } : { part: 0, position: 0, resumed: false }
}

// A stack's part need not be resumable on its own: a per-part record before the last part carries no duration.
const STACK_MIN_RESUME_SECONDS = 30

function planStackStart(
	rec: ItemRec,
	progress: NonNullable<ItemRec["p"]>,
	file: { uuid: string; size: number },
	partDurations: readonly (number | null)[]
): ResumePlan {
	const none: ResumePlan = { part: 0, position: 0, resumed: false }

	if (
		(rec.w !== undefined && compareStamps(progress, rec.w) <= 0) ||
		!progressMatchesFile({ progress, file: { uuid: file.uuid, size: file.size } })
	) {
		return none
	}

	if (isCumulativeStackProgress(partDurations, progress.dur) && !isResumablePosition({ pos: progress.pos, dur: progress.dur })) {
		return none
	}

	const located = locateInStack(partDurations, progress)
	const length = partDurations[located.part] ?? null

	if ((located.part === 0 && located.pos < STACK_MIN_RESUME_SECONDS) || (length !== null && located.pos >= length - 1)) {
		return none
	}

	return { part: located.part, position: located.pos, resumed: true }
}
