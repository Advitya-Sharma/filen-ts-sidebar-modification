import { afterEach, describe, expect, it, vi } from "vitest"
import {
	directErrorAction,
	KEEPALIVE_IDLE_PAUSE_MS,
	RECOVERY_COOLDOWN_MS,
	RECOVERY_TIMEOUT_MS,
	STREAM_STALL_MS,
	STREAM_STALL_POLL_MS,
	stallAction,
	expectsProgress
} from "@/lib/media/playback/directSession.logic"
import { openDirectSession, type DirectSource } from "@/lib/media/playback/directSession"
import type { PlaybackFailure } from "@/lib/media/playback/types"
import type { PreviewStreamStatus } from "@/lib/sw/protocol"
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

function setUp(
	startTime = 0,
	options: {
		status?: PreviewStreamStatus | (() => Promise<PreviewStreamStatus>)
		register?: (id: string) => Promise<string>
		hidden?: () => boolean
	} = {}
) {
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
	const answer = options.status ?? "present"
	const status = vi.fn((_id: string) => (typeof answer === "function" ? answer() : Promise.resolve(answer)))
	const gone = new Map<string, () => void>()
	const source: DirectSource = {
		onGone: (id, listener) => {
			gone.set(id, listener)

			return () => gone.delete(id)
		},
		register,
		release: id => {
			releases.push(id)
		},
		holdAlive: () => {
			holds++

			return releaseAlive
		},
		status
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
		now: () => now,
		hidden: options.hidden ?? (() => false)
	})
	const video = new FakeVideo()

	return {
		session,
		video,
		register,
		status,
		registered,
		releases,
		failures,
		reconnecting,
		holds: () => holds,
		releaseAlive,
		gone,
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
		// Played again once it can, after the seek.
		video.canPlay()
		expect(video.paused).toBe(false)
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

	it("gives the route up for the next rung when a recovery does not bring the stream back, learning nothing", async () => {
		vi.useFakeTimers()

		const { session, video, failures } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		video.fail(2)
		await vi.advanceTimersByTimeAsync(RECOVERY_TIMEOUT_MS)

		expect(failures).toEqual([{ trigger: "stream-unavailable", detail: "the stream did not come back", learnable: false }])
	})

	it("plays on after a recovery when the browser paused the element as its stream broke (Chrome)", async () => {
		const { session, video, advance } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		video.currentTime = 40
		advance(10_000)
		video.pause()
		advance(200)
		video.fail(2)
		await Promise.resolve()
		await Promise.resolve()
		video.loadMetadata()
		video.canPlay()

		expect(video.currentTime).toBe(40)
		expect(video.paused).toBe(false)
	})

	it("plays on after a recovery when the browser's pause comes after the error (Chrome's order)", async () => {
		const { session, video, advance } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		video.currentTime = 40
		advance(10_000)
		// The element is paused and errored first, its pause event queued after the error.
		video.paused = true
		video.fail(2)
		video.emit("pause")
		await Promise.resolve()
		await Promise.resolve()
		video.loadMetadata()
		video.canPlay()

		expect(video.paused).toBe(false)
	})

	it("keeps a user's pause well before the break", async () => {
		const { session, video, advance } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		video.pause()
		advance(30_000)
		video.fail(2)
		await Promise.resolve()
		await Promise.resolve()
		video.loadMetadata()
		video.canPlay()

		expect(video.paused).toBe(true)
	})
})

describe("direct session stall watchdog", () => {
	async function stalled(options: { status: PreviewStreamStatus; hidden?: () => boolean }) {
		vi.useFakeTimers()

		const harness = setUp(0, { status: options.status, ...(options.hidden === undefined ? {} : { hidden: options.hidden }) })
		let now = 0
		const advance = async (ms: number): Promise<void> => {
			for (let step = 0; step < ms; step += STREAM_STALL_POLL_MS) {
				now += STREAM_STALL_POLL_MS
				harness.advance(STREAM_STALL_POLL_MS)
				await vi.advanceTimersByTimeAsync(STREAM_STALL_POLL_MS)
			}
		}

		await harness.session.attach(harness.video.asElement())
		harness.video.loadMetadata()
		await harness.video.play()
		// A WebKit 404 loop: playing, the playhead frozen, readyState left at HAVE_FUTURE_DATA, no MediaError.
		harness.video.readyState = 3

		return { ...harness, advanceStall: advance, now: () => now }
	}

	it("re-registers once when the worker lost the stream, then gives the route up without learning it", async () => {
		const { advanceStall, register, registered, status, failures, reconnecting, video } = await stalled({ status: "lost" })

		await advanceStall(STREAM_STALL_MS + STREAM_STALL_POLL_MS)

		expect(status).toHaveBeenCalledTimes(1)
		expect(status).toHaveBeenCalledWith(registered[0])
		expect(register).toHaveBeenCalledTimes(2)
		expect(reconnecting).toEqual([true])

		// The reload played, then stalled the same way.
		video.loadMetadata()
		video.canPlay()
		video.readyState = 3
		await advanceStall(STREAM_STALL_MS + STREAM_STALL_POLL_MS)

		expect(failures).toEqual([{ trigger: "stream-unavailable", detail: "the stream stalled", learnable: false }])
	})

	it("recovers a stalled stream again once the cooldown has passed since the last recovery", async () => {
		const { advanceStall, advance, register, failures, reconnecting, video } = await stalled({ status: "missing" })

		await advanceStall(STREAM_STALL_MS + STREAM_STALL_POLL_MS)
		video.loadMetadata()
		video.canPlay()
		video.readyState = 3
		// A long stretch of playback since that recovery.
		advance(RECOVERY_COOLDOWN_MS)
		await advanceStall(STREAM_STALL_MS + STREAM_STALL_POLL_MS)

		expect(register).toHaveBeenCalledTimes(3)
		expect(reconnecting).toEqual([true, false, true])
		expect(failures).toEqual([])
	})

	it("waits on a stream the worker still serves (only slow)", async () => {
		const { advanceStall, register, failures } = await stalled({ status: "present" })

		await advanceStall(3 * STREAM_STALL_MS)

		expect(register).toHaveBeenCalledTimes(1)
		expect(failures).toEqual([])
	})

	it("waits on a stream when the worker is older than the page and cannot be asked", async () => {
		const { advanceStall, register, failures, reconnecting } = await stalled({ status: "unknown" })

		await advanceStall(3 * STREAM_STALL_MS)

		expect(register).toHaveBeenCalledTimes(1)
		expect(reconnecting).toEqual([])
		expect(failures).toEqual([])
	})

	it("blames the format for a refusal before metadata when the worker is older than the page", async () => {
		const old = setUp(0, { status: "unknown" })

		await old.session.attach(old.video.asElement())
		old.video.fail(4)
		await vi.waitFor(() => {
			expect(old.failures).toEqual([{ trigger: "media-error", detail: "MediaError 4" }])
		})
		expect(old.register).toHaveBeenCalledTimes(1)
		expect(old.reconnecting).toEqual([])
	})

	it("says nothing while the page is hidden", async () => {
		let hidden = true
		const { advanceStall, status } = await stalled({ status: "missing", hidden: () => hidden })

		await advanceStall(2 * STREAM_STALL_MS)
		hidden = false

		expect(status).not.toHaveBeenCalled()
	})

	it("reports a decode error as a demotion", async () => {
		const decode = setUp()

		await decode.session.attach(decode.video.asElement())
		decode.video.loadMetadata()
		decode.video.fail(3)
		expect(decode.failures).toEqual([{ trigger: "media-error", detail: "MediaError 3" }])
	})

	it("blames the format for a refusal before metadata only when the worker still serves the stream", async () => {
		const format = setUp()

		await format.session.attach(format.video.asElement())
		format.video.fail(4)
		await vi.waitFor(() => {
			expect(format.failures).toEqual([{ trigger: "media-error", detail: "MediaError 4" }])
		})
		expect(format.status).toHaveBeenCalledWith(format.registered[0])
		expect(format.register).toHaveBeenCalledTimes(1)

		// Chrome says the same when the fetch failed at once: that is the stream, recovered, never learned.
		const fetch = setUp(12, { status: "missing" })

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

	it("drops a format verdict whose answer a recovery overtook", async () => {
		const answer = Promise.withResolvers<PreviewStreamStatus>()
		const raced = setUp(0, { status: () => answer.promise })

		await raced.session.attach(raced.video.asElement())
		raced.video.fail(4)
		await Promise.resolve()
		// The worker reports the stream gone while the question is out; that recovery reloads and plays.
		raced.gone.get(raced.registered[0] ?? "")?.()
		await vi.waitFor(() => {
			expect(raced.register).toHaveBeenCalledTimes(2)
		})
		raced.video.loadMetadata()
		raced.video.canPlay()
		answer.resolve("present")
		await new Promise(resolve => setTimeout(resolve, 0))

		expect(raced.failures).toEqual([])
		expect(raced.reconnecting).toEqual([true, false])
	})

	it("drops a stall verdict whose answer a recovery overtook", async () => {
		vi.useFakeTimers()

		const answer = Promise.withResolvers<PreviewStreamStatus>()
		const harness = setUp(0, { status: () => answer.promise })

		await harness.session.attach(harness.video.asElement())
		harness.video.loadMetadata()
		await harness.video.play()
		harness.video.readyState = 3

		for (let step = 0; step <= STREAM_STALL_MS; step += STREAM_STALL_POLL_MS) {
			harness.advance(STREAM_STALL_POLL_MS)
			await vi.advanceTimersByTimeAsync(STREAM_STALL_POLL_MS)
		}

		expect(harness.status).toHaveBeenCalledTimes(1)
		harness.video.fail(2)
		await vi.advanceTimersByTimeAsync(0)
		harness.video.loadMetadata()
		harness.video.canPlay()
		answer.resolve("lost")
		await vi.advanceTimersByTimeAsync(0)

		expect(harness.register).toHaveBeenCalledTimes(2)
		expect(harness.failures).toEqual([])
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

describe("stall decisions", () => {
	it("expects a playing or seeking playhead to move, whatever readyState claims (WebKit's 404 loop holds it at 3)", () => {
		expect(expectsProgress({ paused: false, seeking: false, ended: false })).toBe(true)
		expect(expectsProgress({ paused: true, seeking: true, ended: false })).toBe(true)
		expect(expectsProgress({ paused: true, seeking: false, ended: false })).toBe(false)
		expect(expectsProgress({ paused: false, seeking: false, ended: true })).toBe(false)
	})

	it("re-registers, then demotes within the cooldown; a stream the worker serves is only slow", () => {
		expect(stallAction({ status: "present", now: 0, lastRecoveryAt: null })).toBe("wait")
		expect(stallAction({ status: "unknown", now: 0, lastRecoveryAt: null })).toBe("wait")
		expect(stallAction({ status: "missing", now: 0, lastRecoveryAt: null })).toBe("recover")
		expect(stallAction({ status: "lost", now: 0, lastRecoveryAt: null })).toBe("recover")
		expect(stallAction({ status: "lost", now: RECOVERY_COOLDOWN_MS - 1, lastRecoveryAt: 0 })).toBe("demote")
	})

	it("recovers again once the cooldown has passed since the last recovery", () => {
		expect(stallAction({ status: "missing", now: RECOVERY_COOLDOWN_MS, lastRecoveryAt: 0 })).toBe("recover")
	})
})

describe("direct session told its stream is gone", () => {
	it("takes the element's own error for the same break, landing after the report, as part of that recovery", async () => {
		const { session, video, registered, gone, failures, reconnecting } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		gone.get(registered[0] ?? "")?.()
		video.fail(2)
		await Promise.resolve()
		await Promise.resolve()
		video.loadMetadata()
		video.canPlay()

		expect(failures).toEqual([])
		expect(reconnecting).toEqual([true, false])
		expect(video.paused).toBe(false)
	})

	it("re-registers at once when the worker reports the stream gone, then gives the route up, never learned", async () => {
		const { session, video, register, registered, gone, failures, reconnecting } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		gone.get(registered[0] ?? "")?.()
		await Promise.resolve()
		await Promise.resolve()

		expect(register).toHaveBeenCalledTimes(2)
		expect(reconnecting).toEqual([true])

		video.loadMetadata()
		video.canPlay()
		gone.get(registered[0] ?? "")?.()

		expect(failures).toEqual([{ trigger: "stream-unavailable", detail: "the stream stalled", learnable: false }])

		session.dispose()
		expect(gone.size).toBe(0)
	})

	it("re-registers again for a stream reported gone once the cooldown has passed", async () => {
		const { session, video, register, registered, gone, failures, reconnecting, advance } = setUp()

		await session.attach(video.asElement())
		video.loadMetadata()
		await video.play()
		gone.get(registered[0] ?? "")?.()
		await vi.waitFor(() => {
			expect(register).toHaveBeenCalledTimes(2)
		})
		video.loadMetadata()
		video.canPlay()
		advance(RECOVERY_COOLDOWN_MS)
		gone.get(registered[0] ?? "")?.()
		await vi.waitFor(() => {
			expect(register).toHaveBeenCalledTimes(3)
		})

		expect(reconnecting).toEqual([true, false, true])
		expect(failures).toEqual([])
		session.dispose()
	})
})
