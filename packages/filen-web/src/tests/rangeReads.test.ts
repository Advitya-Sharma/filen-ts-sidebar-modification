import { afterEach, describe, expect, it, vi } from "vitest"
import { BACKGROUND_READS_DURING_PLAYBACK, RANGE_READ_LIMIT, RangeReadBudget } from "@/lib/media/rangeReads"

async function flush(): Promise<void> {
	for (let i = 0; i < 5; i++) {
		await Promise.resolve()
	}
}

function first(): number {
	return 0
}

afterEach(() => {
	vi.useRealTimers()
})

describe("RangeReadBudget", () => {
	it("caps background reads across features at six", async () => {
		const budget = new RangeReadBudget()
		let started = 0

		expect(RANGE_READ_LIMIT).toBe(6)

		for (let i = 0; i < 8; i++) {
			void budget.acquire(first).then(() => started++)
		}

		await flush()
		expect(started).toBe(6)

		budget.release()
		await flush()
		expect(started).toBe(7)
	})

	it("hands a freed permit to the lowest rank", async () => {
		const budget = new RangeReadBudget(1, 1, 50)
		const order: string[] = []

		await budget.acquire(first)
		void budget.acquire(() => 3).then(() => order.push("idle"))
		void budget.acquire(() => 1).then(() => order.push("visible"))

		budget.release()
		await flush()

		expect(order).toEqual(["visible"])
	})

	it("leaves background work two permits while playback runs, and the rest once it stops", async () => {
		const budget = new RangeReadBudget()
		const release = budget.holdPlayback()
		let started = 0

		expect(budget.playing).toBe(true)

		for (let i = 0; i < 5; i++) {
			void budget.acquire(first).then(() => started++)
		}

		await flush()
		expect(started).toBe(BACKGROUND_READS_DURING_PLAYBACK)

		release()
		release()
		await flush()

		expect(budget.playing).toBe(false)
		expect(started).toBe(5)
	})

	it("keeps the playback limit until every playback has ended", () => {
		const budget = new RangeReadBudget(6, 2, 50)
		const a = budget.holdPlayback()
		const b = budget.holdPlayback()

		a()
		expect(budget.playing).toBe(true)
		b()
		expect(budget.playing).toBe(false)
	})

	it("waits for playback to end, or for the waiter's signal", async () => {
		const budget = new RangeReadBudget()

		await expect(budget.whenPlaybackIdle()).resolves.toBeUndefined()

		const release = budget.holdPlayback()
		const waiting = budget.whenPlaybackIdle()
		const controller = new AbortController()
		const aborted = budget.whenPlaybackIdle(controller.signal)

		controller.abort()
		await expect(aborted).rejects.toThrow()

		release()
		await expect(waiting).resolves.toBeUndefined()
	})

	it("runs work in a permit and frees it after", async () => {
		const budget = new RangeReadBudget(1, 1, 50)

		await expect(budget.run(first, () => Promise.resolve(3))).resolves.toBe(3)
		await expect(budget.run(first, () => Promise.reject(new Error("read failed")))).rejects.toThrow("read failed")
		await expect(budget.run(first, () => Promise.resolve(4))).resolves.toBe(4)
	})
})
