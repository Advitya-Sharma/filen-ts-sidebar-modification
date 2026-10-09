import { describe, expect, it } from "vitest"
import { PendingRegistry } from "@/lib/sw/pendingRegistry"

// The service worker's bounded download registry (sw.ts's own state, extracted so it is testable under
// node). What matters: an entry that is still being read — or was read most recently — must survive a
// burst of new registrations, because dropping it 404s a stream mid-playback.

describe("PendingRegistry", () => {
	it("keeps entries until the ceiling is exceeded, then drops the oldest", () => {
		const registry = new PendingRegistry<string>(2)

		registry.set("a", "A")
		registry.set("b", "B")
		registry.set("c", "C")

		expect(registry.size).toBe(2)
		expect(registry.get("a")).toBeUndefined()
		expect(registry.get("b")).toBe("B")
		expect(registry.get("c")).toBe("C")
	})

	it("refreshes recency on read, so a re-read entry is not the next eviction candidate", () => {
		const registry = new PendingRegistry<string>(2)

		registry.set("a", "A")
		registry.set("b", "B")
		registry.get("a")
		registry.set("c", "C")

		expect(registry.get("a")).toBe("A")
		expect(registry.get("b")).toBeUndefined()
	})

	it("treats re-registering an existing id as fresh, so it outlives the next registration", () => {
		const registry = new PendingRegistry<string>(2)

		registry.set("a", "A")
		registry.set("b", "B")
		registry.set("a", "A2")
		registry.set("c", "C")

		expect(registry.get("a")).toBe("A2")
		expect(registry.get("b")).toBeUndefined()
	})

	it("never evicts an entry whose stream is still in flight", () => {
		const registry = new PendingRegistry<string>(2)

		registry.set("a", "A")
		registry.beginStream("a")
		registry.set("b", "B")
		registry.set("c", "C")

		expect(registry.get("a")).toBe("A")
		expect(registry.get("b")).toBeUndefined()

		registry.endStream("a")
		registry.set("d", "D")
		registry.set("e", "E")

		// Protection lasts only as long as the stream does.
		expect(registry.get("a")).toBeUndefined()
		expect(registry.get("e")).toBe("E")
	})

	it("keeps a protected entry rather than evicting a streaming one when everything is in flight", () => {
		const registry = new PendingRegistry<string>(1)

		registry.set("a", "A")
		registry.beginStream("a")
		registry.set("b", "B")

		expect(registry.get("a")).toBe("A")
		expect(registry.get("b")).toBe("B")
	})

	it("counts overlapping reads of one id (a range probe alongside the real fetch)", () => {
		const registry = new PendingRegistry<string>(4)

		registry.set("a", "A")
		registry.beginStream("a")
		registry.beginStream("a")

		expect(registry.activeStreams).toBe(2)

		registry.endStream("a")

		expect(registry.activeStreams).toBe(1)

		registry.set("b", "B")
		registry.set("c", "C")
		registry.set("d", "D")
		registry.set("e", "E")

		// Still streaming under the second read → still protected.
		expect(registry.get("a")).toBe("A")

		registry.endStream("a")

		expect(registry.activeStreams).toBe(0)
	})

	it("clears every entry on logout while in-flight pumps keep their own count", () => {
		const registry = new PendingRegistry<string>(4)

		registry.set("a", "A")
		registry.beginStream("a")
		registry.clear()

		expect(registry.size).toBe(0)
		expect(registry.get("a")).toBeUndefined()
		// The pump still holds its SDK handle and settles on its own — the update gate must stay closed.
		expect(registry.activeStreams).toBe(1)

		registry.endStream("a")

		expect(registry.activeStreams).toBe(0)
	})

	it("peeks at an entry without refreshing its recency", () => {
		const registry = new PendingRegistry<string>(2)

		registry.set("a", "A")
		registry.set("b", "B")
		expect(registry.peek("a")).toBe("A")
		expect(registry.peek("missing")).toBeUndefined()
		registry.set("c", "C")

		expect(registry.peek("a")).toBeUndefined()
		expect(registry.peek("b")).toBe("B")
	})

	it("drops one entry by id, leaving the others and a running stream's count alone", () => {
		const registry = new PendingRegistry<string>(4)

		registry.set("a", "A")
		registry.set("b", "B")
		registry.beginStream("a")
		registry.delete("a")
		registry.delete("missing")

		expect(registry.get("a")).toBeUndefined()
		expect(registry.get("b")).toBe("B")
		expect(registry.size).toBe(1)
		expect(registry.activeStreams).toBe(1)

		registry.endStream("a")

		expect(registry.activeStreams).toBe(0)
	})

	describe("buckets", () => {
		it("never lets a burst in one bucket evict an entry of another, paused or not", () => {
			const registry = new PendingRegistry<string>(2, { transient: 2 })

			registry.set("player", "P")

			for (let index = 0; index < 10; index++) {
				registry.set(`thumb-${String(index)}`, "T", "transient")
			}

			expect(registry.get("player")).toBe("P")
			expect(registry.size).toBe(3)
			expect(registry.get("thumb-7")).toBeUndefined()
			expect(registry.get("thumb-8")).toBe("T")
			expect(registry.get("thumb-9")).toBe("T")
		})

		it("evicts within the bucket by recency, sparing a streaming entry", () => {
			const registry = new PendingRegistry<string>(4, { transient: 2 })

			registry.set("a", "A", "transient")
			registry.beginStream("a")
			registry.set("b", "B", "transient")
			registry.set("c", "C", "transient")

			expect(registry.get("a")).toBe("A")
			expect(registry.get("b")).toBeUndefined()
			expect(registry.get("c")).toBe("C")
		})

		it("moves a re-registered id to its new bucket, and files an unknown bucket under the default one", () => {
			const registry = new PendingRegistry<string>(1, { transient: 1 })

			registry.set("a", "A", "transient")
			registry.set("a", "A2")
			registry.set("t", "T", "transient")

			expect(registry.get("a")).toBe("A2")
			expect(registry.get("t")).toBe("T")

			registry.set("x", "X", "nonsense")

			expect(registry.get("a")).toBeUndefined()
			expect(registry.get("x")).toBe("X")
			expect(registry.get("t")).toBe("T")
		})

		it("deletes and clears across buckets", () => {
			const registry = new PendingRegistry<string>(4, { transient: 4 })

			registry.set("a", "A")
			registry.set("b", "B", "transient")
			registry.delete("b")

			expect(registry.get("b")).toBeUndefined()
			expect(registry.size).toBe(1)

			registry.clear()

			expect(registry.size).toBe(0)
			expect(registry.get("a")).toBeUndefined()
		})
	})

	it("hands every entry that leaves to onDrop: replaced, evicted, deleted and cleared", () => {
		const dropped: string[] = []
		const registry = new PendingRegistry<string>(2, {}, entry => {
			dropped.push(entry)
		})

		registry.set("a", "A")
		registry.set("a", "A2")
		registry.set("b", "B")
		registry.set("c", "C")
		registry.delete("b")
		registry.delete("missing")
		registry.clear()

		expect(dropped).toEqual(["A", "A2", "B", "C"])
	})
})
