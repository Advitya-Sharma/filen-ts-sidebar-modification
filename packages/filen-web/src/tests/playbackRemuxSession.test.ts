import { afterEach, describe, expect, it, vi } from "vitest"
import type { RemuxSession, RemuxSessionOptions } from "@/lib/media/mseController"
import { IdrUnreachableError, SourceReadError } from "@/lib/media/remuxTypes"
import type { PlaybackFailure } from "@/lib/media/playback/types"
import { FakeVideo } from "@/tests/support/fakeVideo"

vi.mock("@/lib/media/mseController", () => ({ openRemuxSession: vi.fn() }))

const { openRemuxPlaybackSession, remuxOpenFailure, OPEN_PORT_TIMEOUT_MS } = await import("@/lib/media/playback/remuxSession")

afterEach(() => {
	vi.useRealTimers()
})

function setUp(
	attach: () => Promise<void> = () => Promise.resolve(),
	openPort: () => Promise<MessagePort> = () => Promise.resolve(new MessageChannel().port2)
) {
	const failures: PlaybackFailure[] = []
	const close = vi.fn()
	const opened: RemuxSessionOptions[] = []
	const remux = {
		attach: vi.fn(attach),
		selectAudio: vi.fn(() => Promise.resolve()),
		dispose: vi.fn()
	}
	const session = openRemuxPlaybackSession({
		path: "remux",
		source: { openPort, close },
		size: 1000,
		index: "cues",
		startTime: 7,
		audio: { index: 0, trackId: 2, mode: "copy" },
		idrOnlyStart: true,
		callbacks: { onFailure: failure => failures.push(failure), onReconnecting: () => undefined },
		open: options => {
			opened.push(options)

			return Promise.resolve(remux as unknown as RemuxSession)
		}
	})

	return { session, failures, close, opened, remux }
}

describe("remuxOpenFailure", () => {
	it.each([
		[new IdrUnreachableError(12), "open-gop-start"],
		[new SourceReadError(new Error("download failed")), "network"],
		[new DOMException("codecs", "NotSupportedError"), "media-error"],
		[new Error("remux worker is gone (terminated)"), "worker-error"]
	])("%s -> %s", (error, trigger) => {
		expect(remuxOpenFailure(error).trigger).toBe(trigger)
	})

	it("recognises errors by name, as they arrive from the worker", () => {
		const crossed = Object.assign(new Error("no IDR picture"), { name: "IdrUnreachableError" })

		expect(remuxOpenFailure(crossed).trigger).toBe("open-gop-start")
	})
})

describe("remux playback session", () => {
	it("opens the remux session with the plan's start, audio and IDR rule", async () => {
		const { session, opened, remux } = setUp()
		const video = new FakeVideo()

		await session.attach(video.asElement())

		expect(opened[0]).toMatchObject({
			size: 1000,
			index: "cues",
			startTime: 7,
			idrOnlyStart: true,
			audio: { trackId: 2, mode: "copy" }
		})
		expect(remux.attach).toHaveBeenCalledWith(video)
	})

	it("maps the controller's failures onto triggers, reporting only the first", async () => {
		const { session, opened, failures } = setUp()

		await session.attach(new FakeVideo().asElement())
		opened[0]?.onFailure("keyframeIntervalTooLarge", "a 200 MB segment")
		opened[0]?.onFailure("remux", "later")

		expect(failures).toEqual([{ trigger: "keyframe-interval", detail: "a 200 MB segment" }])
	})

	it("reports an attach that throws", async () => {
		const { session, failures } = setUp(() => Promise.reject(new IdrUnreachableError(40)))

		await session.attach(new FakeVideo().asElement())

		expect(failures[0]?.trigger).toBe("open-gop-start")
	})

	it("switches audio in place with the ladder's mode, but cannot drop audio", async () => {
		const { session, remux } = setUp()

		await session.attach(new FakeVideo().asElement())

		await expect(session.selectAudio({ index: 1, trackId: 3, mode: "transcode" })).resolves.toBe(true)
		expect(remux.selectAudio).toHaveBeenCalledWith(3, "transcode")
		await expect(session.selectAudio({ index: -1, trackId: -1, mode: "none" })).resolves.toBe(false)
	})

	it("disposes the controller and stops the file's reads", async () => {
		const { session, close, remux } = setUp()

		await session.attach(new FakeVideo().asElement())
		session.dispose()
		session.dispose()

		expect(remux.dispose).toHaveBeenCalledTimes(1)
		expect(close).toHaveBeenCalledTimes(1)
	})

	it("gives up on a byte source that never answers (its worker is gone) as a lost source", async () => {
		vi.useFakeTimers()

		const { session, failures, opened, close } = setUp(undefined, () => new Promise<MessagePort>(() => undefined))
		const attached = session.attach(new FakeVideo().asElement())

		await vi.advanceTimersByTimeAsync(OPEN_PORT_TIMEOUT_MS)
		await attached

		expect(failures[0]?.trigger).toBe("network")
		expect(opened).toHaveLength(0)
		// The source is let go, so a port that opens late frees its client at once.
		expect(close).toHaveBeenCalledOnce()
	})
})
