import { describe, expect, it } from "vitest"
import {
	needsStillWatchingPrompt,
	shouldShowUpNext,
	STILL_WATCHING_LIMIT,
	stillWatchingStep,
	upNextLeadSeconds,
	upNextShowAt
} from "@filen/shared/videos"

const MIN = 60

describe("up next timing", () => {
	it.each([
		[10 * MIN, 30],
		[39 * MIN + 59, 30],
		[40 * MIN, 35],
		[49 * MIN + 59, 35],
		[50 * MIN, 40],
		[3 * 3600, 40]
	])("lead for %s s is %s s", (runtime, lead) => {
		expect(upNextLeadSeconds(runtime)).toBe(lead)
	})

	it("only applies to runtimes of at least 10 minutes", () => {
		expect(upNextShowAt(10 * MIN - 1)).toBeNull()
		expect(upNextShowAt(Number.NaN)).toBeNull()
		expect(upNextShowAt(10 * MIN)).toBe(10 * MIN - 30)
		expect(upNextShowAt(45 * MIN)).toBe(45 * MIN - 35)
		expect(upNextShowAt(60 * MIN)).toBe(60 * MIN - 40)
	})

	it("shows earlier at a known credits start, but never later than the tier", () => {
		expect(upNextShowAt(45 * MIN, 40 * MIN)).toBe(40 * MIN)
		expect(upNextShowAt(45 * MIN, 44 * MIN + 50)).toBe(45 * MIN - 35)
		expect(upNextShowAt(45 * MIN, null)).toBe(45 * MIN - 35)
		expect(upNextShowAt(45 * MIN, 0)).toBe(45 * MIN - 35)
	})

	it("fires inside the tier window and never with less than 20 s left", () => {
		const duration = 25 * MIN

		expect(shouldShowUpNext(duration - 31, duration)).toBe(false)
		expect(shouldShowUpNext(duration - 30, duration)).toBe(true)
		expect(shouldShowUpNext(duration - 20, duration)).toBe(true)
		expect(shouldShowUpNext(duration - 19.9, duration)).toBe(false)
		expect(shouldShowUpNext(duration, duration)).toBe(false)
		expect(shouldShowUpNext(100, 5 * MIN)).toBe(false)
		expect(shouldShowUpNext(duration - 120, duration, duration - 150)).toBe(true)
		expect(shouldShowUpNext(duration - 200, duration, duration - 150)).toBe(false)
	})
})

describe("still watching", () => {
	it("asks after three consecutive autoplays and resets on interaction", () => {
		let count = 0

		expect(STILL_WATCHING_LIMIT).toBe(3)

		count = stillWatchingStep(count, "autoplay")
		count = stillWatchingStep(count, "autoplay")
		expect(needsStillWatchingPrompt(count)).toBe(false)

		count = stillWatchingStep(count, "autoplay")
		expect(needsStillWatchingPrompt(count)).toBe(true)

		expect(stillWatchingStep(count, "confirm")).toBe(0)
		expect(stillWatchingStep(2, "interaction")).toBe(0)
		expect(needsStillWatchingPrompt(stillWatchingStep(stillWatchingStep(2, "interaction"), "autoplay"))).toBe(false)
	})
})
