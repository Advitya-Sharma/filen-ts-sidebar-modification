import { afterEach, describe, expect, it, vi } from "vitest"
import {
	directErrorAction,
	KEEPALIVE_IDLE_PAUSE_MS,
	RECOVERY_COOLDOWN_MS,
	RECOVERY_TIMEOUT_MS
} from "@/lib/media/playback/directSession.logic"
import { openDirectSession, type DirectSource } from "@/lib/media/playback/directSession"
import type { PlaybackFailure } from "@/lib/media/playback/types"
import { FakeVideo } from "@/tests/support/fakeVideo"

afterEach(() => {
	vi.useRealTimers()
})

describe("directErrorAction", () => {
	it.each([
		[{ code: 1, hadMetadata: true, recovering: false, lastRecoveryAt: null }, "ignore"],
		[{ code: 3, hadMetadata: true, recovering: false, lastRecoveryAt: null }, "demote"],
		[{ code: 2, hadMetadata: true, recovering: false, lastRecoveryAt: null }, "recover"],
		[{ code: 2, hadMetadata: false, recovering: false, lastRecoveryAt: null }, "recover"],
		[{ code: 2, hadMetadata: true, recovering: false, lastRecoveryAt: 1000 }, "fail"],
		// MEDIA_ERR_SRC_NOT_SUPPORTED before metadata is the format or a fetch that failed at once: verified.
		[{ code: 4, hadMetadata: false, recovering: false, lastRecoveryAt: null }, "verify"],
		[{ code: 4, hadMetadata: true, recovering: false, lastRecoveryAt: null }, "recover"],
		[{ code: 4, hadMetadata: true, recovering: false, lastRecoveryAt: 1000 }, "fail"],
		// A recovery's own reload failing is the stream, never the format.
		[{ code: 3, hadMetadata: true, recovering: true, lastRecoveryAt: 1000 }, "fail"],
		[{ code: 4, hadMetadata: false, recovering: true, lastRecoveryAt: 1000 }, "fail"]
	] as const)("%o -> %s", (params, action) => {
		expect(directErrorAction({ ...params, now: 2000 })).toBe(action)
	})

	it("recovers again once the cooldown has passed", () => {
		expect(directErrorAction({ code: 2, hadMetadata: true, recovering: false, now: RECOVERY_COOLDOWN_MS, lastRecoveryAt: 0 })).toBe(
			"recover"
		)
	})
})

function setUp(startTime = 0, options: { reachable?: boolean; register?: (id: string) => Promise<string> } = {}) {
	const releases: string[] = []
	const registered: string[] = []
	let holds = 0
	const releaseAlive = vi.fn(() => {
		holds--
	})
	const register = vi.fn((id: string) => {
		registered.push(id)

		return options.register?.(id) ?? Promise.resolve(`/sw/download/${id}`)
	})
	const reachable = vi.fn(() => Promise.resolve(options.reachable ?? true))
	const source: DirectSource = {
		register,
		release: id => {
			releases.push(id)
		},
		holdAlive: () => {
			holds++

			return releaseAlive
		},
		reachable
	}
	const failures: PlaybackFailure[] = []
	const reconnecting: boolean[] = []
	let now = 0
	const session = openDirectSession({
		source,
		startTime,
		audioIndex: 0,
		callbacks: {
			onFailure: failure => failures.push(failure),
			onReconnecting: value => reconnecting.push(value)
		},
		now: () => now
	})
	const video = new FakeVideo()

	return {
		session,
		video,
		register,
		reachable,
		registered,
		releases,
		failures,
		reconnecting,
		holds: () => holds,
		releaseAlive,
		advance: (ms: number) => {
			now += ms
		}
	}
}

describe("direct session", () => {
	it("streams the registered URL and seeks to the start once metadata is in", async () => {
		const { session, video, registered } = setUp(42)

		await session.attach(video.asElement())

		expect(video.src).toBe(`/sw/download/${registered[0] ?? ""}`)
		expect(video.disableRemotePlayback).toBe(true)

		video.loadMetadata()

		expect(video.currentTime).toBe(42)
	})

	it("keeps the worker alive while mounted, paused too, and lets go after a long pause", async () => {
		vi.useFakeTimers()

		const { session, video, holds, releaseAlive } = setUp()

		await session.attach(video.asElement())
		expect(holds()).toBe(1)

		await video.play()
		video.pause()
		vi.advanceTimersByTime(KEEPALIVE_IDLE_PAUSE_MS - 1)
		expect(holds()).toBe(1)
		vi.advanceTimersByTime(1)
		expect(holds()).toBe(0)

		// Playing again holds it again.
		await video.play()
		expect(holds()).toBe(1)

		session.dispose()
		expect(holds()).toBe(0)
		expect(releaseAlive).toHaveBeenCalledTimes(2)
	})

	it("recovers a lost stream at the saved position and play state, once per cooldown", async () => {
		const { session, video, register, registered, reconnecting, failures, advance } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		video.currentTime = 95
		advance(10_000)
		// The worker was terminated during a pause and the element's next range request 404ed.
		video.fail(2)
		await Promise.resolve()
		await Promise.resolve()

		expect(reconnecting).toEqual([true])
		expect(register).toHaveBeenCalledTimes(2)
		expect(registered[1]).toBe(registered[0])
		expect(video.loads).toBe(1)

		video.loadMetadata()
		expect(video.currentTime).toBe(95)
		expect(video.paused).toBe(false)
		video.canPlay()
		expect(reconnecting).toEqual([true, false])

		// Broken again right away: no second recovery inside the cooldown.
		advance(1000)
		video.fail(2)
		expect(failures).toEqual([{ trigger: "network", detail: "MediaError 2" }])
	})

	it("keeps a paused player paused across a recovery", async () => {
		const { session, video } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		video.currentTime = 30
		video.fail(4)
		await Promise.resolve()
		await Promise.resolve()
		video.loadMetadata()

		expect(video.currentTime).toBe(30)
		expect(video.paused).toBe(true)
		expect(video.plays).toBe(0)
	})

	it("fails a recovery that does not bring the stream back in time", async () => {
		vi.useFakeTimers()

		const { session, video, failures } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		video.fail(2)
		await vi.advanceTimersByTimeAsync(RECOVERY_TIMEOUT_MS)

		expect(failures).toEqual([{ trigger: "network", detail: "the stream did not come back" }])
	})

	it("reports a decode error as a demotion", async () => {
		const decode = setUp()

		await decode.session.attach(decode.video.asElement())
		decode.video.loadMetadata()
		decode.video.fail(3)
		expect(decode.failures).toEqual([{ trigger: "media-error", detail: "MediaError 3" }])
	})

	it("blames the format for a refusal before metadata only when the stream answers a range request", async () => {
		const format = setUp()

		await format.session.attach(format.video.asElement())
		format.video.fail(4)
		await vi.waitFor(() => {
			expect(format.failures).toEqual([{ trigger: "media-error", detail: "MediaError 4" }])
		})
		expect(format.reachable).toHaveBeenCalledWith(`/sw/download/${format.registered[0] ?? ""}`)
		expect(format.register).toHaveBeenCalledTimes(1)

		// Chrome says the same when the fetch failed at once: that is the stream, recovered, never learned.
		const fetch = setUp(12, { reachable: false })

		await fetch.session.attach(fetch.video.asElement())
		fetch.video.fail(4)
		await vi.waitFor(() => {
			expect(fetch.reconnecting).toEqual([true])
		})
		expect(fetch.register).toHaveBeenCalledTimes(2)
		fetch.video.loadMetadata()
		expect(fetch.video.currentTime).toBe(12)

		// The reload failing too is the stream as well.
		fetch.video.fail(4)
		expect(fetch.failures).toEqual([{ trigger: "network", detail: "MediaError 4" }])
	})

	it("reports a registration that fails as the stream being unavailable, not the file", async () => {
		const { session, video, failures } = setUp(0, { register: () => Promise.reject(new Error("no active service worker")) })

		await expect(session.attach(video.asElement())).resolves.toBeUndefined()
		expect(failures).toEqual([{ trigger: "stream-unavailable", detail: "no active service worker" }])
		expect(video.src).toBe("")
	})

	it("honours a pause during a reconnect, and lets the worker go after the idle pause again", async () => {
		vi.useFakeTimers()

		const { session, video, holds } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		video.fail(2)
		await vi.advanceTimersByTimeAsync(0)
		// The reload's own pause, then the user's.
		await vi.advanceTimersByTimeAsync(0)
		video.paused = false
		video.pause()
		video.loadMetadata()
		video.canPlay()

		expect(video.paused).toBe(true)
		expect(video.plays).toBe(1)
		expect(holds()).toBe(1)
		await vi.advanceTimersByTimeAsync(KEEPALIVE_IDLE_PAUSE_MS)
		expect(holds()).toBe(0)
		session.dispose()
	})

	it("empties the element before releasing its registration", async () => {
		const { session, video, releases, registered } = setUp()
		const order: string[] = []

		await session.attach(video.asElement())
		video.addEventListener("emptied", () => order.push(`emptied ${String(releases.length)}`))
		session.dispose()

		expect(order).toEqual(["emptied 0"])
		expect(releases).toEqual([registered[0]])
		expect(video.src).toBe("")
	})

	it("hands a registration made after dispose straight back", async () => {
		const { session, video, releases, registered } = setUp()
		const attached = session.attach(video.asElement())

		session.dispose()
		await attached

		expect(releases).toEqual([registered[0]])
		expect(video.src).toBe("")
	})
})
