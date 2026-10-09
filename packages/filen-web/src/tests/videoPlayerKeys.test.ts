import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { chapterTarget, frameDuration, stepPlaybackRate, videoKeyAction, type VideoKeyInput } from "@/lib/media/videoKeys.logic"
import { SEEK_COALESCE_MS, SeekCoalescer } from "@/lib/media/seekCoalescer.logic"

function key(input: Partial<VideoKeyInput> & { key: string }) {
	return videoKeyAction({ target: "surface", modified: false, shift: false, ...input })
}

describe("videoKeyAction", () => {
	it("maps the player's keys on the picture", () => {
		expect(key({ key: " " })).toEqual({ type: "togglePlay" })
		expect(key({ key: "k" })).toEqual({ type: "togglePlay" })
		expect(key({ key: "K" })).toEqual({ type: "togglePlay" })
		expect(key({ key: "Enter" })).toEqual({ type: "confirm" })
		expect(key({ key: "ArrowLeft" })).toEqual({ type: "seekBy", seconds: -10 })
		expect(key({ key: "j" })).toEqual({ type: "seekBy", seconds: -10 })
		expect(key({ key: "ArrowRight" })).toEqual({ type: "seekBy", seconds: 10 })
		expect(key({ key: "l" })).toEqual({ type: "seekBy", seconds: 10 })
		expect(key({ key: "ArrowUp" })).toEqual({ type: "volumeBy", delta: 0.05 })
		expect(key({ key: "ArrowDown" })).toEqual({ type: "volumeBy", delta: -0.05 })
		expect(key({ key: "m" })).toEqual({ type: "toggleMute" })
		expect(key({ key: "f" })).toEqual({ type: "toggleFullscreen" })
		expect(key({ key: "c" })).toEqual({ type: "toggleSubtitles" })
		expect(key({ key: "s" })).toEqual({ type: "skip" })
		expect(key({ key: "i" })).toEqual({ type: "togglePictureInPicture" })
		expect(key({ key: "g" })).toEqual({ type: "subtitleDelay", seconds: -0.1 })
		expect(key({ key: "h" })).toEqual({ type: "subtitleDelay", seconds: 0.1 })
		expect(key({ key: "Home" })).toEqual({ type: "seekToStart" })
		expect(key({ key: "End" })).toEqual({ type: "seekToEnd" })
		expect(key({ key: "PageDown" })).toEqual({ type: "chapter", direction: 1 })
		expect(key({ key: "PageUp" })).toEqual({ type: "chapter", direction: -1 })
		expect(key({ key: "," })).toEqual({ type: "frameStep", direction: -1 })
		expect(key({ key: "." })).toEqual({ type: "frameStep", direction: 1 })
		expect(key({ key: "Escape" })).toEqual({ type: "dismiss" })
	})

	it("maps digits to tenths of the file", () => {
		expect(key({ key: "0" })).toEqual({ type: "seekToFraction", fraction: 0 })
		expect(key({ key: "5" })).toEqual({ type: "seekToFraction", fraction: 0.5 })
		expect(key({ key: "9" })).toEqual({ type: "seekToFraction", fraction: 0.9 })
	})

	it("takes Shift combinations but leaves Ctrl, Cmd and Alt chords alone", () => {
		expect(key({ key: "N", shift: true })).toEqual({ type: "next" })
		expect(key({ key: "P", shift: true })).toEqual({ type: "previous" })
		expect(key({ key: "<", shift: true })).toEqual({ type: "rate", direction: -1 })
		expect(key({ key: ">", shift: true })).toEqual({ type: "rate", direction: 1 })
		expect(key({ key: "?", shift: true })).toEqual({ type: "shortcuts" })
		expect(key({ key: "K", shift: true })).toBeNull()
		expect(key({ key: "k", modified: true })).toBeNull()
		expect(key({ key: "ArrowRight", modified: true })).toBeNull()
	})

	it("leaves a slider its arrows, a button its Space and Enter, and menus and fields everything", () => {
		expect(key({ key: "ArrowLeft", target: "slider" })).toBeNull()
		expect(key({ key: "Home", target: "slider" })).toBeNull()
		expect(key({ key: "PageUp", target: "slider" })).toBeNull()
		expect(key({ key: "m", target: "slider" })).toEqual({ type: "toggleMute" })
		expect(key({ key: " ", target: "button" })).toBeNull()
		expect(key({ key: "Enter", target: "button" })).toBeNull()
		expect(key({ key: "k", target: "button" })).toEqual({ type: "togglePlay" })
		expect(key({ key: "k", target: "menu" })).toBeNull()
		expect(key({ key: "k", target: "field" })).toBeNull()
	})
})

describe("player key helpers", () => {
	it("steps the playback rate along the list from wherever it is", () => {
		expect(stepPlaybackRate(1, 1)).toBe(1.25)
		expect(stepPlaybackRate(1, -1)).toBe(0.75)
		expect(stepPlaybackRate(2, 1)).toBe(2)
		expect(stepPlaybackRate(0.5, -1)).toBe(0.5)
		expect(stepPlaybackRate(1.1, 1)).toBe(1.25)
		expect(stepPlaybackRate(1.1, -1)).toBe(1)
	})

	it("finds the next and previous chapter, going back past a chapter's first seconds", () => {
		const starts = [0, 60, 300, 900]

		expect(chapterTarget(starts, 100, 1)).toBe(300)
		expect(chapterTarget(starts, 100, -1)).toBe(60)
		expect(chapterTarget(starts, 61, -1)).toBe(0)
		expect(chapterTarget(starts, 950, 1)).toBeNull()
		expect(chapterTarget([], 10, 1)).toBeNull()
	})

	it("uses the frame rate for a frame step, 24 fps when unknown", () => {
		expect(frameDuration(25)).toBeCloseTo(0.04)
		expect(frameDuration(null)).toBeCloseTo(1 / 24)
	})
})

describe("SeekCoalescer", () => {
	let now = 0
	let position = 100
	let seeks: number[] = []
	let pending: (number | null)[] = []

	function make(): SeekCoalescer {
		return new SeekCoalescer({
			current: () => position,
			duration: () => 600,
			seek: seconds => {
				seeks.push(seconds)
				position = seconds
			},
			onPending: seconds => {
				pending.push(seconds)
			},
			now: () => now,
			setTimer: (callback, ms) => setTimeout(callback, ms),
			clearTimer: timer => {
				clearTimeout(timer)
			}
		})
	}

	beforeEach(() => {
		vi.useFakeTimers()
		now = 0
		position = 100
		seeks = []
		pending = []
	})

	afterEach(() => {
		vi.useRealTimers()
	})

	it("seeks at once on the first press", () => {
		make().seekBy(10)

		expect(seeks).toEqual([110])
		expect(pending).toEqual([])
	})

	it("collapses repeats within the window into one trailing seek", () => {
		const coalescer = make()

		coalescer.seekBy(10)
		now += 50
		coalescer.seekBy(10)
		now += 50
		coalescer.seekBy(10)
		now += 50
		coalescer.seekBy(10)

		expect(seeks).toEqual([110])
		expect(coalescer.target()).toBe(140)
		expect(pending.at(-1)).toBe(140)

		vi.advanceTimersByTime(SEEK_COALESCE_MS)

		expect(seeks).toEqual([110, 140])
		expect(pending.at(-1)).toBeNull()
	})

	it("starts a new window after a pause in the presses", () => {
		const coalescer = make()

		coalescer.seekBy(-10)
		now += SEEK_COALESCE_MS + 1
		coalescer.seekBy(-10)

		expect(seeks).toEqual([90, 80])
	})

	it("clamps to the file and drops a pending seek on cancel", () => {
		const coalescer = make()

		coalescer.seekTo(700)
		now += 10
		coalescer.seekBy(-1000)
		coalescer.cancel()
		vi.advanceTimersByTime(SEEK_COALESCE_MS)

		expect(seeks).toEqual([600])
		expect(pending.at(-1)).toBeNull()
	})
})
