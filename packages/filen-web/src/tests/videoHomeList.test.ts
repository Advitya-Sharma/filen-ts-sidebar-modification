// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { VideoHomeItem, VideoIndexUpdate } from "@/workers/videoIndex/view.logic"

// The main thread's copy of the library's videos: held only while a screen reading it is mounted, patched by the
// worker's diffs, and its subscription retried when the worker call fails.

const { sdkApi, listener } = vi.hoisted(() => ({
	sdkApi: {
		videoIndexSubscribe: vi.fn(),
		videoIndexSetWriter: vi.fn(),
		videoHomeSubscribe: vi.fn<(enabled: boolean) => Promise<void>>()
	},
	listener: { current: null as ((update: VideoIndexUpdate) => void) | null }
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi }))
vi.mock("@/queries/client", async () => ({ queryClient: (await import("@/tests/testQueryClient")).createTestQueryClient() }))
vi.mock("@/queries/persist", () => ({ persister: { persistQuery: () => Promise.resolve() } }))

import { videoLibrary } from "@/features/videos/lib/index"
import { episodeItem } from "@/tests/support/videoHome"

function send(update: VideoIndexUpdate): void {
	listener.current?.(update)
}

beforeEach(() => {
	vi.useFakeTimers()
	videoLibrary.reset()
	sdkApi.videoIndexSubscribe.mockImplementation((callback: (update: VideoIndexUpdate) => void) => {
		listener.current = callback
	})
	sdkApi.videoHomeSubscribe.mockResolvedValue(undefined)
})

afterEach(() => {
	vi.useRealTimers()
})

describe("library video list", () => {
	it("subscribes while a screen holds it, applies the worker's diffs, and lets go a while after the last release", async () => {
		const first = videoLibrary.retainHome()
		const second = videoLibrary.retainHome()

		expect(sdkApi.videoHomeSubscribe.mock.calls).toEqual([[true]])

		const pilot: VideoHomeItem = episodeItem("lost", 1, 1)
		const next: VideoHomeItem = episodeItem("lost", 1, 2)

		send({ type: "home", reset: true, upserts: [pilot, next], removed: [] })

		expect(videoLibrary.getHomeItems()?.map(item => item.uuid)).toEqual([pilot.uuid, next.uuid])

		// A replaced episode file: same stable uuid, new uuid.
		send({ type: "home", reset: false, upserts: [{ ...next, uuid: "rotated" }], removed: [pilot.stableUuid] })

		expect(videoLibrary.getHomeItems()?.map(item => item.uuid)).toEqual(["rotated"])

		first()
		second()
		await vi.advanceTimersByTimeAsync(29_000)

		// Back within the grace (a screen switch): nothing is torn down.
		const again = videoLibrary.retainHome()

		await vi.advanceTimersByTimeAsync(60_000)

		expect(sdkApi.videoHomeSubscribe.mock.calls).toEqual([[true]])

		again()
		await vi.advanceTimersByTimeAsync(30_000)

		expect(sdkApi.videoHomeSubscribe.mock.calls).toEqual([[true], [false]])
		expect(videoLibrary.getHomeItems()).toBeNull()
	})

	it("retries a failed subscription with backoff, then on the next focus", async () => {
		sdkApi.videoHomeSubscribe.mockRejectedValue(new Error("worker"))

		const release = videoLibrary.retainHome()

		await vi.advanceTimersByTimeAsync(0)
		expect(sdkApi.videoHomeSubscribe).toHaveBeenCalledTimes(1)

		await vi.advanceTimersByTimeAsync(1_000)
		expect(sdkApi.videoHomeSubscribe).toHaveBeenCalledTimes(2)

		await vi.advanceTimersByTimeAsync(2_000 + 4_000 + 8_000 + 16_000)
		expect(sdkApi.videoHomeSubscribe).toHaveBeenCalledTimes(6)

		await vi.advanceTimersByTimeAsync(120_000)
		expect(sdkApi.videoHomeSubscribe).toHaveBeenCalledTimes(6)

		sdkApi.videoHomeSubscribe.mockResolvedValue(undefined)
		window.dispatchEvent(new Event("focus"))
		await vi.advanceTimersByTimeAsync(0)

		expect(sdkApi.videoHomeSubscribe).toHaveBeenCalledTimes(7)
		expect(sdkApi.videoHomeSubscribe).toHaveBeenLastCalledWith(true)

		release()
	})
})
