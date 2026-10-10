// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, render } from "@testing-library/react"
import { createElement } from "react"
import { CueBuilder, type CueStore } from "@filen/shared/videos"

vi.mock("@/lib/storage/adapter", () => ({ kvGetJson: () => Promise.resolve(null), kvSetJson: () => Promise.resolve() }))

import { SubtitleOverlay } from "@/components/media/video/subtitleOverlay"

// A text track as browsers have it: a disabled track exposes no cues.
class FakeTrack {
	public mode: TextTrackMode = "disabled"
	public added = 0
	private readonly list: unknown[] = []

	public get cues(): unknown[] | null {
		return this.mode === "disabled" ? null : this.list
	}

	public addCue(cue: unknown): void {
		this.added++
		this.list.push(cue)
	}

	public removeCue(cue: unknown): void {
		this.list.splice(this.list.indexOf(cue), 1)
	}

	public get stored(): number {
		return this.list.length
	}

	public addEventListener = (): undefined => undefined
	public removeEventListener = (): undefined => undefined
}

function store(count: number): CueStore {
	const builder = new CueBuilder()

	for (let i = 0; i < count; i++) {
		builder.add(i * 2, i * 2 + 1, `cue ${String(i)}`)
	}

	return builder.finish()
}

let tracks: FakeTrack[]
let video: HTMLVideoElement

beforeEach(() => {
	vi.useFakeTimers()
	tracks = []
	video = Object.assign(new EventTarget(), {
		currentTime: 0,
		addTextTrack: () => {
			const track = new FakeTrack()

			tracks.push(track)

			return track
		}
	}) as unknown as HTMLVideoElement
	vi.stubGlobal(
		"VTTCue",
		class {
			public line: number | "auto" = "auto"
			public snapToLines = true
			public align = "center"
			public startTime: number
			public endTime: number
			public text: string
			public constructor(startTime: number, endTime: number, text: string) {
				this.startTime = startTime
				this.endTime = endTime
				this.text = text
			}
		}
	)
})

afterEach(() => {
	cleanup()
	vi.useRealTimers()
	vi.unstubAllGlobals()
})

describe("SubtitleOverlay's text track", () => {
	it("adds embedded cues as they grow at once, and refills only after a delay change settles", () => {
		const view = render(createElement(SubtitleOverlay, { video, cues: store(2), offset: 0 }))

		expect(tracks[0]?.stored).toBe(2)

		// A new block of cues: only its cues go in, with no wait (picture-in-picture shows the track itself).
		view.rerender(createElement(SubtitleOverlay, { video, cues: store(5), offset: 0 }))
		expect(tracks[0]?.stored).toBe(5)
		expect(tracks[0]?.added).toBe(5)

		view.rerender(createElement(SubtitleOverlay, { video, cues: store(5), offset: 0.5 }))
		view.rerender(createElement(SubtitleOverlay, { video, cues: store(6), offset: 1 }))
		expect(tracks.reduce((sum, track) => sum + track.added, 0)).toBe(5)

		act(() => {
			vi.advanceTimersByTime(300)
		})

		expect(tracks.reduce((sum, track) => sum + track.stored, 0)).toBe(6)
		expect(tracks.reduce((sum, track) => sum + track.added, 0)).toBe(11)
	})
})
