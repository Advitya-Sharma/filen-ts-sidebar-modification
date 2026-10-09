import type { PlayerKeyTarget } from "@/lib/media/playerKeys.logic"

// The full-screen video player's keyboard. Single keys and Shift combinations are the player's; chords with
// Ctrl, Cmd or Alt stay the browser's and the app's. A slider keeps its own arrows, Home/End and Page keys,
// a button its own Space/Enter, a menu or field everything.

export const VIDEO_SEEK_STEP_SECONDS = 10
export const VIDEO_VOLUME_STEP = 0.05
export const SUBTITLE_DELAY_STEP_SECONDS = 0.1
export const PLAYBACK_RATES = [0.5, 0.75, 1, 1.25, 1.5, 1.75, 2] as const

export type VideoKeyAction =
	| { type: "togglePlay" }
	| { type: "seekBy"; seconds: number }
	| { type: "seekToFraction"; fraction: number }
	| { type: "seekToStart" }
	| { type: "seekToEnd" }
	| { type: "chapter"; direction: 1 | -1 }
	| { type: "frameStep"; direction: 1 | -1 }
	| { type: "volumeBy"; delta: number }
	| { type: "toggleMute" }
	| { type: "toggleFullscreen" }
	| { type: "toggleSubtitles" }
	| { type: "skip" }
	| { type: "next" }
	| { type: "previous" }
	| { type: "rate"; direction: 1 | -1 }
	| { type: "subtitleDelay"; seconds: number }
	| { type: "togglePictureInPicture" }
	| { type: "shortcuts" }
	| { type: "dismiss" }
	// Enter: the shown card's primary action (Play now, Skip), else play/pause.
	| { type: "confirm" }

export interface VideoKeyInput {
	key: string
	target: PlayerKeyTarget
	// Ctrl, Cmd or Alt held.
	modified: boolean
	shift: boolean
}

const SLIDER_KEYS: ReadonlySet<string> = new Set(["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End", "PageUp", "PageDown"])

export function videoKeyAction({ key, target, modified, shift }: VideoKeyInput): VideoKeyAction | null {
	if (modified || target === "menu" || target === "field") {
		return null
	}

	if (target === "slider" && SLIDER_KEYS.has(key)) {
		return null
	}

	if ((key === " " || key === "Enter") && target === "button") {
		return null
	}

	// Shifted punctuation arrives as its own character on most layouts ("<", ">", "?"); letters as capitals.
	switch (key) {
		case "<":
			return { type: "rate", direction: -1 }
		case ">":
			return { type: "rate", direction: 1 }
		case "?":
			return { type: "shortcuts" }
		case "Escape":
			return { type: "dismiss" }
		default:
			break
	}

	const lower = key.length === 1 ? key.toLowerCase() : key

	if (shift) {
		switch (lower) {
			case "n":
				return { type: "next" }
			case "p":
				return { type: "previous" }
			default:
				return null
		}
	}

	if (lower.length === 1 && lower >= "0" && lower <= "9") {
		return { type: "seekToFraction", fraction: Number(lower) / 10 }
	}

	switch (lower) {
		case " ":
		case "k":
			return { type: "togglePlay" }
		case "Enter":
			return { type: "confirm" }
		case "ArrowLeft":
		case "j":
			return { type: "seekBy", seconds: -VIDEO_SEEK_STEP_SECONDS }
		case "ArrowRight":
		case "l":
			return { type: "seekBy", seconds: VIDEO_SEEK_STEP_SECONDS }
		case "ArrowUp":
			return { type: "volumeBy", delta: VIDEO_VOLUME_STEP }
		case "ArrowDown":
			return { type: "volumeBy", delta: -VIDEO_VOLUME_STEP }
		case "Home":
			return { type: "seekToStart" }
		case "End":
			return { type: "seekToEnd" }
		case "PageUp":
			return { type: "chapter", direction: -1 }
		case "PageDown":
			return { type: "chapter", direction: 1 }
		case ",":
			return { type: "frameStep", direction: -1 }
		case ".":
			return { type: "frameStep", direction: 1 }
		case "m":
			return { type: "toggleMute" }
		case "f":
			return { type: "toggleFullscreen" }
		case "c":
			return { type: "toggleSubtitles" }
		case "s":
			return { type: "skip" }
		case "g":
			return { type: "subtitleDelay", seconds: -SUBTITLE_DELAY_STEP_SECONDS }
		case "h":
			return { type: "subtitleDelay", seconds: SUBTITLE_DELAY_STEP_SECONDS }
		case "i":
			return { type: "togglePictureInPicture" }
		default:
			return null
	}
}

// The next playback rate up or down the list, from wherever the current one sits.
export function stepPlaybackRate(current: number, direction: 1 | -1): number {
	if (direction > 0) {
		return PLAYBACK_RATES.find(rate => rate > current + 1e-6) ?? PLAYBACK_RATES[PLAYBACK_RATES.length - 1] ?? current
	}

	let lower: number = PLAYBACK_RATES[0]

	for (const rate of PLAYBACK_RATES) {
		if (rate < current - 1e-6) {
			lower = rate
		}
	}

	return lower
}

// Chapter starts to step through with PageUp/PageDown. Going back from within the first seconds of a chapter
// lands on the one before it, like a track list.
export const CHAPTER_BACK_GRACE_SECONDS = 3

export function chapterTarget(starts: readonly number[], time: number, direction: 1 | -1): number | null {
	if (starts.length === 0) {
		return null
	}

	if (direction > 0) {
		return starts.find(start => start > time + 0.5) ?? null
	}

	let target: number | null = null

	for (const start of starts) {
		if (start <= time - CHAPTER_BACK_GRACE_SECONDS) {
			target = start
		}
	}

	return target ?? 0
}

// Seconds per frame from the frame rate when known, else a common 24 fps.
export function frameDuration(fps: number | null): number {
	return fps !== null && fps > 0 && Number.isFinite(fps) ? 1 / fps : 1 / 24
}
