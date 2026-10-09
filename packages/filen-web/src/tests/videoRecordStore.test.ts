import { describe, expect, it, vi } from "vitest"
import { type } from "arktype"
import { ABSENT_TTL_MS, RecordStore, planCountEviction, type RecordStoreKv } from "@/features/videos/lib/recordStore.logic"
import { memoryKv } from "@/tests/support/memoryKv"

const schema = type({ value: "number", at: "number" })

type Row = typeof schema.infer

function store(
	kv: RecordStoreKv,
	options: { maxRecords?: number; memoryRecords?: number; writesPerSweep?: number; now?: () => number } = {}
): RecordStore<Row> {
	return new RecordStore({
		prefix: "test.v1.",
		schema,
		maxRecords: options.maxRecords ?? 100,
		memoryRecords: options.memoryRecords ?? 10,
		writesPerSweep: options.writesPerSweep ?? 1000,
		scope: "test",
		kv,
		...(options.now === undefined ? {} : { now: options.now })
	})
}

async function settle(): Promise<void> {
	for (let i = 0; i < 10; i++) {
		await Promise.resolve()
	}
}

describe("planCountEviction", () => {
	it("evicts the oldest writes down to the target once over the cap", () => {
		const records = [
			["a", 3],
			["b", 1],
			["c", 4],
			["d", 2]
		] as const

		expect(planCountEviction(records, 4, 2)).toEqual([])
		expect(planCountEviction(records, 3, 2)).toEqual(["b", "d"])
	})
})

describe("RecordStore", () => {
	it("reads a record once, then from memory; writes land in memory at once and in kv behind", async () => {
		const { kv, rows, reads } = memoryKv()
		const records = store(kv)

		rows.set("test.v1.x", { value: 1, at: 1 })

		expect(records.peek("x")).toBeUndefined()
		expect(await records.get("x")).toEqual({ value: 1, at: 1 })
		expect(await records.get("x")).toEqual({ value: 1, at: 1 })
		expect(await records.get("missing")).toBeNull()
		expect(await records.get("missing")).toBeNull()
		expect(reads()).toBe(2)

		records.put("y", { value: 2, at: 2 })
		expect(records.peek("y")).toEqual({ value: 2, at: 2 })
		await settle()
		expect(rows.get("test.v1.y")).toEqual({ value: 2, at: 2 })

		records.delete("y")
		expect(records.peek("y")).toBeNull()
		await settle()
		expect(rows.has("test.v1.y")).toBe(false)
	})

	it("coalesces concurrent reads of one id", async () => {
		const { kv, reads } = memoryKv()
		const records = store(kv)

		await Promise.all([records.get("x"), records.get("x"), records.get("x")])

		expect(reads()).toBe(1)
	})

	it("keeps a bounded memory, least recently used out first", async () => {
		const { kv, reads } = memoryKv()
		const records = store(kv, { memoryRecords: 2 })

		records.put("a", { value: 1, at: 1 })
		records.put("b", { value: 2, at: 2 })
		records.peek("a")
		records.put("c", { value: 3, at: 3 })

		expect(records.peek("b")).toBeUndefined()
		expect(records.peek("a")).toEqual({ value: 1, at: 1 })
		await settle()
		expect(await records.get("b")).toEqual({ value: 2, at: 2 })
		expect(reads()).toBe(1)
	})

	it("holds the prefix to its cap on the first write of a session, oldest first, dropping invalid rows", async () => {
		const { kv, rows } = memoryKv()
		const keys = vi.spyOn(kv, "kvKeys")
		const entries = vi.spyOn(kv, "kvEntriesJson")

		for (let i = 0; i < 10; i++) {
			rows.set(`test.v1.r${String(i)}`, { value: i, at: i })
		}

		rows.set("test.v1.broken", { nope: true })
		rows.set("other.v1.r0", { value: 0, at: 0 })

		const records = store(kv, { maxRecords: 10 })

		records.put("new", { value: 99, at: 99 })
		await settle()
		await records.sweep()

		// Eleven valid rows over a cap of ten go down to nine, the oldest first.
		expect([...rows.keys()].filter(key => key.startsWith("test.v1.")).sort()).toEqual(
			["new", "r2", "r3", "r4", "r5", "r6", "r7", "r8", "r9"].map(id => `test.v1.${id}`)
		)
		expect(rows.has("other.v1.r0")).toBe(true)
		expect(keys).toHaveBeenCalled()
		expect(entries).toHaveBeenCalled()
	})

	it("lists only keys while the prefix is under its cap", async () => {
		const { kv } = memoryKv()
		const entries = vi.spyOn(kv, "kvEntriesJson")
		const records = store(kv, { maxRecords: 10, writesPerSweep: 2 })

		records.put("a", { value: 1, at: 1 })
		records.put("b", { value: 1, at: 1 })
		records.put("c", { value: 1, at: 1 })
		await settle()

		expect(entries).not.toHaveBeenCalled()
	})

	it("forgets memory at reset, and a read in flight writes nothing back", async () => {
		const { kv, rows } = memoryKv()
		const records = store(kv)

		rows.set("test.v1.x", { value: 1, at: 1 })

		const reading = records.get("x")

		records.reset()
		await reading

		expect(records.peek("x")).toBeUndefined()
	})

	it("trusts a known absence only briefly, and refresh reads past it", async () => {
		const { kv, rows, reads } = memoryKv()
		let now = 0
		const records = store(kv, { now: () => now })

		expect(await records.get("x")).toBeNull()

		// Another tab writes the record.
		rows.set("test.v1.x", { value: 1, at: 1 })
		expect(await records.get("x")).toBeNull()
		expect(await records.refresh("x")).toEqual({ value: 1, at: 1 })
		expect(reads()).toBe(2)

		records.delete("x")
		await settle()
		rows.set("test.v1.x", { value: 2, at: 2 })
		now = ABSENT_TTL_MS + 1
		expect(await records.get("x")).toEqual({ value: 2, at: 2 })
	})

	it("keeps a row rewritten after the sweep planned to evict it", async () => {
		const { kv, rows } = memoryKv()
		const entries = kv.kvEntriesJson

		for (let i = 0; i < 4; i++) {
			rows.set(`test.v1.r${String(i)}`, { value: i, at: i })
		}

		// The rewrite lands between the plan and the delete.
		kv.kvEntriesJson = async (prefix, check) => {
			const planned = await entries(prefix, check)

			rows.set("test.v1.r0", { value: 0, at: 100 })

			return planned
		}

		await store(kv, { maxRecords: 3 }).sweep()

		expect(rows.has("test.v1.r0")).toBe(true)
		expect(rows.has("test.v1.r1")).toBe(false)
	})

	it("keeps at most eight deletes in flight however many rows go", async () => {
		const { kv, rows } = memoryKv()
		const remove = kv.kvDelete
		let inFlight = 0
		let peak = 0

		for (let i = 0; i < 100; i++) {
			rows.set(`test.v1.r${String(i)}`, { value: i, at: i })
		}

		kv.kvDelete = async key => {
			inFlight++
			peak = Math.max(peak, inFlight)
			await new Promise(resolve => setTimeout(resolve, 0))
			inFlight--

			return remove(key)
		}

		await store(kv, { maxRecords: 10 }).sweep()

		expect(rows.size).toBe(9)
		expect(peak).toBeLessThanOrEqual(8)
	})
})
