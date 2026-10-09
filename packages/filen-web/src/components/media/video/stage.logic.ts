// The video stage's rules: when its controls show, when a stall becomes a spinner and then a message, where a
// double tap seeks. Plain and DOM-free.

// The Videos player hides its controls after 3 s untouched (the preview keeps its 2.5 s).
export const STAGE_CONTROLS_IDLE_MS = 3000
// A stall shorter than this shows nothing.
export const STAGE_SPINNER_DELAY_MS = 400
// A stall this long says so and offers Retry.
export const STAGE_STILL_LOADING_MS = 15_000
// Two taps this close are a double tap.
export const DOUBLE_TAP_MS = 300

export interface StageChapter {
	start: number
	title: string | null
}

export interface StageVisibility {
	paused: boolean
	waiting: boolean
	// A pointer moved over the stage within the idle period.
	pointerActive: boolean
	keyboardFocus: boolean
	// A menu, popover or panel of the stage is open.
	popupOpen: boolean
	scrubbing: boolean
	// A card asks for the user (still watching, post-play, autoplay blocked).
	blocking: boolean
}

export function stageControlsVisible(state: StageVisibility): boolean {
	return (
		state.paused || state.waiting || state.pointerActive || state.keyboardFocus || state.popupOpen || state.scrubbing || state.blocking
	)
}

// The chapter `time` falls in, from start-ordered chapters.
export function chapterAt(chapters: readonly StageChapter[], time: number): StageChapter | null {
	let found: StageChapter | null = null

	for (const chapter of chapters) {
		if (chapter.start <= time) {
			found = chapter
		} else {
			break
		}
	}

	return found
}

// Where a double tap seeks: back in the left third, forward in the right one, nothing in the middle.
export function doubleTapZone(x: number, width: number): -1 | 0 | 1 {
	if (width <= 0) {
		return 0
	}

	const ratio = x / width

	return ratio < 1 / 3 ? -1 : ratio > 2 / 3 ? 1 : 0
}

export type StallState = "none" | "spinner" | "stillLoading"

// What a stall that began `waitingFor` ms ago shows.
export function stallState(waitingFor: number | null): StallState {
	if (waitingFor === null || waitingFor < STAGE_SPINNER_DELAY_MS) {
		return "none"
	}

	return waitingFor >= STAGE_STILL_LOADING_MS ? "stillLoading" : "spinner"
}
