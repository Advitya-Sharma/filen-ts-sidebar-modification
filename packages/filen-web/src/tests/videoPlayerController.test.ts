import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { PlaybackSnapshot } from "@/lib/media/playback/openPlayback"
import type { VideoProbe } from "@/lib/media/probe/types"
import type { WatchItemTarget } from "@/features/videos/lib/player/watchTarget.logic"

const mocks = vi.hoisted(() => {
	const playbacks: { status: string; listeners: Set<() => void>; disposed: boolean; startTime: number }[] = []

	return {
		playbacks,
		openPlayback: vi.fn(),
		watchOpen: vi.fn(() => Promise.resolve()),
		reportProgress: vi.fn(),
		reportStop: vi.fn(() => Promise.resolve()),
		dropPlayback: vi.fn(),
		playableFiles: vi.fn()
	}
})

vi.mock("@/lib/media/playback/openPlayback", () => ({ openPlayback: mocks.openPlayback }))
vi.mock("@/features/videos/lib/index", () => ({ videoLibrary: { playableFiles: mocks.playableFiles } }))
vi.mock("@/features/videos/lib/probeService", () => ({
	probeService: { request: () => Promise.resolve({ probe: PROBE }) }
}))
vi.mock("@/features/videos/lib/watchState", () => ({
	watchState: {
		open: mocks.watchOpen,
		getSnapshot: () => ({ doc: { v: 1, items: {}, series: {} } }),
		reportProgress: mocks.reportProgress,
		reportStop: mocks.reportStop,
		dropPlayback: mocks.dropPlayback,
		setSeriesLanguages: () => Promise.resolve()
	}
}))
vi.mock("@/features/videos/lib/frames", () => ({ captureFrameFromVideo: () => Promise.resolve(null) }))
vi.mock("@/features/videos/lib/artwork", () => ({ storeCapturedFrame: () => Promise.resolve() }))

import { PlayerController, WATCH_STATE_WAIT_MS } from "@/features/videos/lib/player/playerController"

const PROBE: VideoProbe = {
	v: 1,
	container: "mp4",
	durationSec: 3600,
	index: "moov",
	video: null,
	audio: [],
	subs: [],
	chapters: [],
	cover: null,
	coverLand: null,
	partial: false
}

const TARGET: WatchItemTarget = {
	key: "movie",
	titleId: "title",
	titleName: "Movie",
	kind: "movie",
	itemClass: "movie",
	versionLabel: "",
	parts: [
		{
			file: { stableUuid: "movie", uuid: "movie-u", name: "Movie.mp4", size: 1000 },
			container: "mp4",
			unsupported: false,
			height: 1080,
			subtitles: [],
			nfo: null,
			thumb: null,
			edl: null
		}
	],
	episode: null,
	seriesDirUuid: null
}

// A media element as far as the controller reads it.
class FakeVideo extends EventTarget {
	public readonly HAVE_METADATA = 1
	public currentTime = 0
	public duration = Number.NaN
	public readyState = 0
	public paused = true
	public ended = false
	public seeking = false
	public refusePlay = false
	public play = vi.fn(() => {
		if (this.refusePlay) {
			return Promise.reject(new DOMException("no activation", "NotAllowedError"))
		}

		this.paused = false
		this.dispatchEvent(new Event("play"))

		return Promise.resolve()
	})
	public pause(): void {
		this.paused = true
		this.dispatchEvent(new Event("pause"))
	}

	public loaded(time: number): void {
		this.readyState = 4
		this.duration = 3600
		this.currentTime = time
		this.dispatchEvent(new Event("loadedmetadata"))
	}

	public tick(time: number): void {
		this.currentTime = time
		this.dispatchEvent(new Event("timeupdate"))
	}
}

function playbackFor(index: number) {
	const entry = mocks.playbacks[index]

	if (entry === undefined) {
		throw new Error("no playback")
	}

	return entry
}

function setStatus(index: number, status: string): void {
	const entry = playbackFor(index)

	entry.status = status

	for (const listener of entry.listeners) {
		listener()
	}
}

async function flush(): Promise<void> {
	for (let i = 0; i < 10; i++) {
		await Promise.resolve()
	}
}

let video: FakeVideo
let controller: PlayerController

beforeEach(() => {
	mocks.playbacks.length = 0
	mocks.openPlayback.mockImplementation(
		(_video: unknown, _file: unknown, _probe: unknown, selection: { audioIndex: number }, options: { startTime: number }) => {
			const entry = { status: "opening", listeners: new Set<() => void>(), disposed: false, startTime: options.startTime }

			mocks.playbacks.push(entry)

			return Promise.resolve({
				getSnapshot: (): PlaybackSnapshot => ({
					status: entry.status as PlaybackSnapshot["status"],
					decision: { path: "direct", reasons: [], audio: { mode: "native" } },
					selection: { audioIndex: selection.audioIndex, subtitle: "none" },
					failure: entry.status === "failed" ? { trigger: "network", detail: "test" } : null,
					demotions: 0
				}),
				subscribe: (listener: () => void) => {
					entry.listeners.add(listener)

					return () => entry.listeners.delete(listener)
				},
				selectAudio: () => Promise.resolve(),
				dispose: () => {
					entry.disposed = true
				}
			})
		}
	)
	mocks.playableFiles.mockResolvedValue([{ uuid: "movie-u", stableUUID: "movie", size: 1000n }])
	video = new FakeVideo()
	controller = new PlayerController()
	controller.attach(video as unknown as HTMLVideoElement)
})

afterEach(() => {
	controller.dispose()
	vi.useRealTimers()
})

async function openPlaying(time: number): Promise<void> {
	await controller.open({ target: TARGET, start: "resume", autoplay: false })
	video.loaded(time)
	setStatus(0, "ready")
	video.paused = false
	video.tick(time)
}

describe("PlayerController and the watch state", () => {
	it("drops the engine's playback when an item closes before it played (never journals a fake 0)", async () => {
		await controller.open({ target: TARGET, start: "resume", autoplay: false })
		setStatus(0, "failed")
		controller.stop()

		expect(mocks.reportStop).not.toHaveBeenCalled()
		expect(mocks.reportProgress).not.toHaveBeenCalled()
		expect(mocks.dropPlayback).toHaveBeenCalledWith({ stableUuid: "movie" })
	})

	it("ends a failed item at the last position it really reported", async () => {
		await openPlaying(1200)
		setStatus(0, "failed")
		// A failed element reads 0 from now on.
		video.currentTime = 0
		video.tick(0)
		controller.stop()

		expect(mocks.reportProgress).toHaveBeenCalledTimes(1)
		expect(mocks.reportStop).toHaveBeenCalledWith(expect.objectContaining({ pos: 1200, dur: 3600 }))
		expect(mocks.dropPlayback).not.toHaveBeenCalled()
	})

	it("retries at the last real position, not at the element's 0", async () => {
		await openPlaying(1500)
		setStatus(0, "failed")
		video.currentTime = 0
		await controller.retry()

		expect(mocks.playbacks[1]?.startTime).toBe(1500)
		expect(mocks.reportStop).toHaveBeenCalledWith(expect.objectContaining({ pos: 1500 }))
	})

	it("retries at the resume point when nothing played yet", async () => {
		await controller.open({ target: TARGET, start: "resume", autoplay: false })
		setStatus(0, "failed")
		await controller.retry()

		expect(mocks.playbacks[1]?.startTime).toBe(0)
	})

	it("reports a finish (Up Next past the credits) as played to the end, once", async () => {
		await openPlaying(3500)
		controller.finish()
		controller.stop()

		expect(mocks.reportStop).toHaveBeenCalledTimes(1)
		expect(mocks.reportStop).toHaveBeenCalledWith(expect.objectContaining({ pos: 3500, ended: true }))
	})

	it("starts without waiting more than 1.5 s for the watch state, and not at all for a part switch", async () => {
		vi.useFakeTimers()
		mocks.watchOpen.mockReturnValue(new Promise(() => undefined))

		const opening = controller.open({ target: TARGET, start: "resume", autoplay: false })

		await vi.advanceTimersByTimeAsync(WATCH_STATE_WAIT_MS - 100)
		expect(mocks.openPlayback).not.toHaveBeenCalled()
		await vi.advanceTimersByTimeAsync(200)
		await opening
		expect(mocks.openPlayback).toHaveBeenCalledTimes(1)

		mocks.watchOpen.mockClear()
		await controller.open({ target: TARGET, start: "resume", autoplay: false }, { part: 0, position: 10 })
		expect(mocks.watchOpen).not.toHaveBeenCalled()
		expect(mocks.openPlayback).toHaveBeenCalledTimes(2)
	})
})

describe("PlayerController autoplay", () => {
	it("reports a refused play, clears it on the element's play, and plays through the controller", async () => {
		video.refusePlay = true
		await controller.open({ target: TARGET, start: "resume", autoplay: true })
		video.loaded(0)
		await flush()

		expect(controller.getSnapshot().autoplayBlocked).toBe(true)

		video.refusePlay = false
		controller.play()
		await flush()

		expect(video.play).toHaveBeenCalledTimes(2)
		expect(controller.getSnapshot().autoplayBlocked).toBe(false)
	})
})
