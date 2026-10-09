import { describe, expect, it } from "vitest"
import {
	bufferedAheadOf,
	DECODE_STALL_MS,
	DemotionWatchLogic,
	FIRST_FRAME_TIMEOUT_MS,
	METADATA_TIMEOUT_MS,
	type WatchConfig,
	type WatchSample
} from "@/lib/media/playback/demotionWatch.logic"

const CONFIG: WatchConfig = { metadataTimeout: true, frameCallback: true, expectAudio: true }

function sample(overrides: Partial<WatchSample>): WatchSample {
	return {
		now: 0,
		paused: false,
		seeking: false,
		readyState: 4,
		currentTime: 0,
		bufferedAhead: 10,
		audioDecodedBytes: 1000,
		...overrides
	}
}

describe("DemotionWatchLogic", () => {
	it("fires no-metadata only for a stream still without metadata at the deadline, paused or not", () => {
		const logic = new DemotionWatchLogic(CONFIG)

		logic.start(0)
		expect(logic.needsClock).toBe(true)
		expect(logic.sample(sample({ now: METADATA_TIMEOUT_MS - 1, paused: true, readyState: 0 }))).toBeNull()
		expect(logic.sample(sample({ now: METADATA_TIMEOUT_MS, paused: true, readyState: 0 }))).toBe("no-metadata")
		// Once only.
		expect(logic.sample(sample({ now: METADATA_TIMEOUT_MS + 1000, paused: true, readyState: 0 }))).toBeNull()

		const loaded = new DemotionWatchLogic(CONFIG)

		loaded.start(0)
		loaded.loadedMetadata()
		expect(loaded.needsClock).toBe(false)
		expect(loaded.sample(sample({ now: 60_000 }))).toBeNull()

		const remux = new DemotionWatchLogic({ ...CONFIG, metadataTimeout: false })

		remux.start(0)
		expect(remux.sample(sample({ now: 60_000, readyState: 0, bufferedAhead: 0 }))).toBeNull()
	})

	it("fires no-first-frame when canplay while playing brings no frame in time", () => {
		const logic = new DemotionWatchLogic(CONFIG)

		logic.start(0)
		logic.loadedMetadata()
		logic.canPlayEvent(1000, false)
		expect(logic.sample(sample({ now: 1000 + FIRST_FRAME_TIMEOUT_MS - 1 }))).toBeNull()
		expect(logic.sample(sample({ now: 1000 + FIRST_FRAME_TIMEOUT_MS }))).toBe("no-first-frame")
	})

	it("waits for playback before timing the first frame, and stands down once one is presented", () => {
		const paused = new DemotionWatchLogic(CONFIG)

		paused.start(0)
		paused.loadedMetadata()
		paused.canPlayEvent(1000, true)
		expect(paused.sample(sample({ now: 60_000, paused: true }))).toBeNull()
		paused.playing(60_000)
		expect(paused.sample(sample({ now: 60_000 + FIRST_FRAME_TIMEOUT_MS }))).toBe("no-first-frame")

		const shown = new DemotionWatchLogic(CONFIG)

		shown.start(0)
		shown.loadedMetadata()
		shown.canPlayEvent(1000, false)
		shown.frame()
		expect(shown.sample(sample({ now: 60_000 }))).toBeNull()

		const noCallback = new DemotionWatchLogic({ ...CONFIG, frameCallback: false })

		noCallback.start(0)
		noCallback.loadedMetadata()
		noCallback.canPlayEvent(0, false)
		expect(noCallback.sample(sample({ now: 60_000 }))).toBeNull()
	})

	it("fires decode-stall for data buffered ahead that the element will not play", () => {
		const logic = new DemotionWatchLogic(CONFIG)

		logic.start(0)
		logic.loadedMetadata()
		logic.frame()
		expect(logic.sample(sample({ now: 1000, readyState: 2, bufferedAhead: 5 }))).toBeNull()
		expect(logic.sample(sample({ now: 1000 + DECODE_STALL_MS - 1, readyState: 2, bufferedAhead: 5 }))).toBeNull()
		expect(logic.sample(sample({ now: 1000 + DECODE_STALL_MS, readyState: 2, bufferedAhead: 5 }))).toBe("decode-stall")
	})

	it("does not count a stall that is only waiting for data, a seek, or one that clears", () => {
		const logic = new DemotionWatchLogic(CONFIG)

		logic.start(0)
		logic.loadedMetadata()
		logic.frame()
		expect(logic.sample(sample({ now: 0, readyState: 2, bufferedAhead: 1 }))).toBeNull()
		expect(logic.sample(sample({ now: 10_000, readyState: 2, bufferedAhead: 1 }))).toBeNull()
		expect(logic.sample(sample({ now: 10_000, readyState: 2, bufferedAhead: 5, seeking: true }))).toBeNull()
		expect(logic.sample(sample({ now: 20_000, readyState: 2, bufferedAhead: 5 }))).toBeNull()
		expect(logic.sample(sample({ now: 21_000, readyState: 4, bufferedAhead: 5 }))).toBeNull()
		expect(logic.sample(sample({ now: 22_000, readyState: 2, bufferedAhead: 5 }))).toBeNull()
		expect(logic.sample(sample({ now: 24_000, readyState: 2, bufferedAhead: 5 }))).toBeNull()
	})

	it("fires silent-audio after three seconds played with nothing decoded, where the engine counts", () => {
		const logic = new DemotionWatchLogic(CONFIG)
		const run = (bytes: number | null, expectAudio = true) => {
			const watch = new DemotionWatchLogic({ ...CONFIG, expectAudio })
			const fired: (string | null)[] = []

			watch.start(0)
			watch.loadedMetadata()
			watch.frame()

			for (let step = 0; step <= 8; step++) {
				fired.push(watch.sample(sample({ now: step * 500, currentTime: step * 0.5, audioDecodedBytes: bytes })))
			}

			return fired.filter(trigger => trigger !== null)
		}

		expect(run(0)).toEqual(["silent-audio"])
		expect(run(4096)).toEqual([])
		expect(run(null)).toEqual([])
		expect(run(0, false)).toEqual([])

		// A seek's jump is not playback.
		logic.start(0)
		logic.loadedMetadata()
		logic.frame()
		expect(logic.sample(sample({ now: 0, currentTime: 0, audioDecodedBytes: 0 }))).toBeNull()
		expect(logic.sample(sample({ now: 500, currentTime: 600, audioDecodedBytes: 0 }))).toBeNull()
	})

	it("stops the frame and stall timers while the tab is hidden, starts them over when visible, and marks the run untrusted", () => {
		const logic = new DemotionWatchLogic(CONFIG)

		logic.start(0)
		logic.loadedMetadata()
		logic.canPlayEvent(0, false)
		logic.playing(0)
		expect(logic.trusted).toBe(true)

		logic.visibility(true, 1000)
		expect(logic.sample(sample({ now: 60_000, readyState: 2, bufferedAhead: 5 }))).toBeNull()

		logic.visibility(false, 60_000)
		expect(logic.trusted).toBe(false)
		expect(logic.sample(sample({ now: 60_000 + FIRST_FRAME_TIMEOUT_MS - 1 }))).toBeNull()
		expect(logic.sample(sample({ now: 60_000 + FIRST_FRAME_TIMEOUT_MS }))).toBe("no-first-frame")
	})

	it("keeps a run trusted when the tab was hidden only while paused, and moves the metadata deadline along", () => {
		const logic = new DemotionWatchLogic(CONFIG)

		logic.start(0)
		logic.visibility(true, 1000)
		logic.visibility(false, 31_000)
		expect(logic.trusted).toBe(true)
		expect(logic.sample(sample({ now: 31_000 + METADATA_TIMEOUT_MS - 1001, paused: true, readyState: 0 }))).toBeNull()
		expect(logic.sample(sample({ now: 31_000 + METADATA_TIMEOUT_MS - 1000, paused: true, readyState: 0 }))).toBe("no-metadata")
	})

	it("says nothing while suspended, and starts over on start", () => {
		const logic = new DemotionWatchLogic(CONFIG)

		logic.start(0)
		logic.suspend()
		expect(logic.needsClock).toBe(false)
		expect(logic.sample(sample({ now: 60_000, readyState: 0 }))).toBeNull()

		logic.start(60_000)
		expect(logic.sample(sample({ now: 60_000 + METADATA_TIMEOUT_MS, readyState: 0 }))).toBe("no-metadata")
	})
})

describe("bufferedAheadOf", () => {
	it("measures the range holding the playhead only", () => {
		expect(
			bufferedAheadOf(
				[
					[0, 5],
					[10, 20]
				],
				3
			)
		).toBe(2)
		expect(bufferedAheadOf([[10, 20]], 9.95)).toBeCloseTo(10.05)
		expect(bufferedAheadOf([[10, 20]], 5)).toBe(0)
	})
})
