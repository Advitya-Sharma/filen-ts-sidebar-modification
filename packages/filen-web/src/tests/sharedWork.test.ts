import { describe, expect, it, vi } from "vitest"
import { onAbort } from "@/lib/abort"
import { SharedWork, type SharedWorkContext } from "@/lib/sharedWork"

describe("onAbort", () => {
	it("runs at once for a signal that already aborted, and once later otherwise", () => {
		const done = new AbortController()
		const later = new AbortController()
		const now = vi.fn()
		const next = vi.fn()

		done.abort()
		onAbort(done.signal, now)
		onAbort(later.signal, next)
		expect(now).toHaveBeenCalledTimes(1)
		expect(next).not.toHaveBeenCalled()

		later.abort()
		expect(next).toHaveBeenCalledTimes(1)
	})

	it("stops listening when unsubscribed, and does nothing without a signal", () => {
		const controller = new AbortController()
		const callback = vi.fn()

		onAbort(controller.signal, callback)()
		controller.abort()
		onAbort(undefined, callback)()

		expect(callback).not.toHaveBeenCalled()
	})
})

describe("SharedWork", () => {
	it("runs a key once while it is in flight, and again after it settled", async () => {
		const shared = new SharedWork<number>()
		const work = vi.fn(() => Promise.resolve(1))

		const [a, b] = await Promise.all([shared.run("k", {}, work), shared.run("k", {}, work)])

		expect([a, b]).toEqual([1, 1])
		expect(work).toHaveBeenCalledTimes(1)

		await shared.run("k", {}, work)
		expect(work).toHaveBeenCalledTimes(2)
	})

	it("aborts the work only once every caller has given up", () => {
		const shared = new SharedWork<number>()
		const seen: SharedWorkContext[] = []
		const finish = Promise.withResolvers<number>()
		const a = new AbortController()
		const b = new AbortController()

		void shared.run("k", { signal: a.signal }, context => {
			seen.push(context)

			return finish.promise
		})
		void shared.run("k", { signal: b.signal }, () => finish.promise)

		a.abort()
		expect(seen[0]?.signal.aborted).toBe(false)
		b.abort()
		expect(seen[0]?.signal.aborted).toBe(true)
		finish.resolve(0)
	})

	it("never aborts work a caller without a signal joined", () => {
		const shared = new SharedWork<number>()
		const seen: SharedWorkContext[] = []
		const a = new AbortController()

		void shared.run("k", { signal: a.signal }, context => {
			seen.push(context)

			return new Promise<number>(() => undefined)
		})
		void shared.run("k", {}, () => Promise.resolve(0))
		a.abort()

		expect(seen[0]?.signal.aborted).toBe(false)
	})

	it("ranks the work by its most urgent caller still waiting", () => {
		const shared = new SharedWork<number>()
		const seen: SharedWorkContext[] = []
		const urgent = new AbortController()

		void shared.run("k", { signal: new AbortController().signal, rank: () => 3 }, context => {
			seen.push(context)

			return new Promise<number>(() => undefined)
		})
		void shared.run("k", { signal: urgent.signal, rank: () => 1 }, () => Promise.resolve(0))

		expect(seen[0]?.rank()).toBe(1)
		urgent.abort()
		expect(seen[0]?.rank()).toBe(3)
	})
})
