import { describe, expect, it } from "vitest"
import { holdReadsWhilePlaying } from "@/lib/media/playbackReads"
import { RangeReadBudget } from "@/lib/media/rangeReads"

class FakeMedia extends EventTarget {
	paused = true
}

function media(paused = true): HTMLMediaElement {
	const element = new FakeMedia()

	element.paused = paused

	return element as unknown as HTMLMediaElement
}

describe("holdReadsWhilePlaying", () => {
	it("holds playback's claim while playing or waiting for data, and gives it back on pause, end or teardown", () => {
		const budget = new RangeReadBudget()
		const element = media()
		const dispose = holdReadsWhilePlaying(element, budget)

		expect(budget.playing).toBe(false)

		element.dispatchEvent(new Event("playing"))
		element.dispatchEvent(new Event("waiting"))
		expect(budget.playing).toBe(true)

		element.dispatchEvent(new Event("pause"))
		expect(budget.playing).toBe(false)

		element.dispatchEvent(new Event("waiting"))
		element.dispatchEvent(new Event("ended"))
		expect(budget.playing).toBe(false)

		element.dispatchEvent(new Event("playing"))
		dispose()
		expect(budget.playing).toBe(false)

		element.dispatchEvent(new Event("playing"))
		expect(budget.playing).toBe(false)
	})

	it("holds at once for an element already playing", () => {
		const budget = new RangeReadBudget()
		const dispose = holdReadsWhilePlaying(media(false), budget)

		expect(budget.playing).toBe(true)
		dispose()
		expect(budget.playing).toBe(false)
	})
})
