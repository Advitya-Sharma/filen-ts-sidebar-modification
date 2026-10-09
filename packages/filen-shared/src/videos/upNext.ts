export const UP_NEXT_MIN_RUNTIME_SECONDS = 10 * 60
export const UP_NEXT_MIN_REMAINING_SECONDS = 20
export const UP_NEXT_CHECK_INTERVAL_MS = 700
// Consecutive autoplayed episodes with no control interaction before the player asks.
export const STILL_WATCHING_LIMIT = 3

// Lead time before the end: 30 s, 35 s from 40 min, 40 s from 50 min.
export function upNextLeadSeconds(runtimeSeconds: number): number {
	if (runtimeSeconds >= 50 * 60) {
		return 40
	}

	return runtimeSeconds >= 40 * 60 ? 35 : 30
}

// Playback position at which the Up Next card appears; null for runtimes under 10 minutes.
// A known credits start earlier than the tier shows it sooner.
export function upNextShowAt(duration: number, creditsStart?: number | null): number | null {
	if (!Number.isFinite(duration) || duration < UP_NEXT_MIN_RUNTIME_SECONDS) {
		return null
	}

	const tier = duration - upNextLeadSeconds(duration)

	if (creditsStart !== undefined && creditsStart !== null && Number.isFinite(creditsStart) && creditsStart > 0 && creditsStart < tier) {
		return creditsStart
	}

	return tier
}

// Evaluated at most every `UP_NEXT_CHECK_INTERVAL_MS`. Never with less than 20 s left.
export function shouldShowUpNext(position: number, duration: number, creditsStart?: number | null): boolean {
	const showAt = upNextShowAt(duration, creditsStart)

	return showAt !== null && position >= showAt && duration - position >= UP_NEXT_MIN_REMAINING_SECONDS
}

export type StillWatchingEvent = "autoplay" | "interaction" | "confirm"

// Counter step: an autoplayed episode counts up, any control interaction or the Continue button resets it.
export function stillWatchingStep(count: number, event: StillWatchingEvent): number {
	return event === "autoplay" ? count + 1 : 0
}

// True when the next item must wait behind Continue / Back instead of autoplaying.
export function needsStillWatchingPrompt(consecutiveAutoplays: number): boolean {
	return consecutiveAutoplays >= STILL_WATCHING_LIMIT
}
