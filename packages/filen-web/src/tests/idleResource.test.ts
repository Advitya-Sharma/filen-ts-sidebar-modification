import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { idleResource } from "@/lib/idleResource"

beforeEach(() => {
	vi.useFakeTimers()
})

afterEach(() => {
	vi.useRealTimers()
})

function harness() {
	let next = 0
	const create = vi.fn(() => ({ id: ++next }))
	const dispose = vi.fn<(resource: { id: number }) => void>()

	return { create, dispose, resource: idleResource(create, dispose, 1000) }
}

describe("idleResource", () => {
	it("creates nothing until first used", () => {
		const h = harness()

		expect(h.create).not.toHaveBeenCalled()
	})

	it("reuses one resource across uses inside the idle window", async () => {
		const h = harness()

		await h.resource.use(() => Promise.resolve())
		await vi.advanceTimersByTimeAsync(999)
		await h.resource.use(() => Promise.resolve())

		expect(h.create).toHaveBeenCalledTimes(1)
		expect(h.dispose).not.toHaveBeenCalled()
	})

	it("disposes once idle for the whole window, then creates a fresh one on the next use", async () => {
		const h = harness()

		await h.resource.use(() => Promise.resolve())
		await vi.advanceTimersByTimeAsync(1000)

		expect(h.dispose).toHaveBeenCalledWith({ id: 1 })

		await h.resource.use(resource => {
			expect(resource).toEqual({ id: 2 })

			return Promise.resolve()
		})
	})

	it("never disposes while a use is in flight, however long it takes", async () => {
		const h = harness()
		let finish: () => void = () => undefined
		const running = h.resource.use(
			() =>
				new Promise<void>(resolve => {
					finish = resolve
				})
		)

		await vi.advanceTimersByTimeAsync(5000)
		h.resource.disposeIfIdle()

		expect(h.dispose).not.toHaveBeenCalled()

		finish()
		await running
		await vi.advanceTimersByTimeAsync(1000)

		expect(h.dispose).toHaveBeenCalledTimes(1)
	})

	it("disposeIfIdle disposes an idle resource right away", async () => {
		const h = harness()

		await h.resource.use(() => Promise.resolve())
		h.resource.disposeIfIdle()

		expect(h.dispose).toHaveBeenCalledTimes(1)

		await vi.advanceTimersByTimeAsync(1000)

		expect(h.dispose).toHaveBeenCalledTimes(1)
	})

	it("arms the idle timer after a failed use too", async () => {
		const h = harness()

		await expect(h.resource.use(() => Promise.reject(new Error("boom")))).rejects.toThrow("boom")
		await vi.advanceTimersByTimeAsync(1000)

		expect(h.dispose).toHaveBeenCalledTimes(1)
	})

	it("caches nothing when create throws", async () => {
		const h = harness()
		h.create.mockImplementationOnce(() => {
			throw new Error("spawn failed")
		})

		await expect(h.resource.use(() => Promise.resolve())).rejects.toThrow("spawn failed")
		await h.resource.use(() => Promise.resolve())

		expect(h.create).toHaveBeenCalledTimes(2)
	})

	it("invalidate disposes a broken resource at once, even in use, and the next use gets a fresh one", async () => {
		const h = harness()
		const inUse = Promise.withResolvers<undefined>()
		const seen: { id: number }[] = []
		const using = h.resource.use(resource => {
			seen.push(resource)

			return inUse.promise
		})

		await Promise.resolve()

		const [first] = seen

		if (first === undefined) {
			throw new Error("not in use")
		}

		h.resource.invalidate(first)

		expect(h.dispose).toHaveBeenCalledWith({ id: 1 })

		await h.resource.use(resource => {
			expect(resource).toEqual({ id: 2 })

			return Promise.resolve()
		})

		inUse.resolve(undefined)
		await using
		// The invalidated one is not disposed twice.
		expect(h.dispose).toHaveBeenCalledTimes(1)
	})
})
