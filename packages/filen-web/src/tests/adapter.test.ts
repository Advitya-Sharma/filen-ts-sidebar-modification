import { beforeEach, describe, expect, it, vi } from "vitest"
import { type } from "arktype"
import { kvClear, kvDeleteIfAt, kvEntriesJson, kvEntriesJsonInSlices, kvGetJson, kvHas, kvSetJson } from "@/lib/storage/adapter"
import { log } from "@/lib/log"

// node vitest cannot provide navigator.locks/BroadcastChannel/real workers (leader.ts's actual
// election machinery), so the whole leader module is replaced with a Map-backed fake StorageApi —
// this tests the adapter facade only (envelope + arktype validation), never leader election itself
// (that needs a manual two-tab dev smoke test, plus a scripted e2e spec later).
const { fakeStore, fakeApi } = vi.hoisted(() => {
	const fakeStore = new Map<string, string>()
	const fakeApi = {
		open: () => Promise.resolve(undefined),
		kvGet: (key: string) => Promise.resolve(fakeStore.get(key) ?? null),
		kvSet: (key: string, value: string) => {
			fakeStore.set(key, value)
			return Promise.resolve()
		},
		kvDelete: (key: string) => {
			fakeStore.delete(key)
			return Promise.resolve()
		},
		kvKeys: (prefix: string) => Promise.resolve([...fakeStore.keys()].filter(k => k.startsWith(prefix))),
		kvEntries: (prefix: string) => Promise.resolve([...fakeStore.entries()].filter(([k]) => k.startsWith(prefix))),
		kvEntriesPage: (prefix: string, after: string | null, limit: number) =>
			Promise.resolve(
				[...fakeStore.entries()]
					.filter(([k]) => k.startsWith(prefix) && (after === null || k > after))
					.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
					.slice(0, limit)
			),
		kvDeletePrefix: (prefix: string) => {
			for (const key of [...fakeStore.keys()]) {
				if (key.startsWith(prefix)) {
					fakeStore.delete(key)
				}
			}

			return Promise.resolve()
		},
		wipe: () => {
			fakeStore.clear()

			return Promise.resolve()
		}
	}

	return { fakeStore, fakeApi }
})

vi.mock("@/lib/storage/leader", () => ({
	acquireStorage: () => Promise.resolve({ role: "leader" as const, api: fakeApi })
}))

beforeEach(() => {
	fakeStore.clear()
	vi.restoreAllMocks()
})

describe("storage adapter (Map-backed fake StorageApi)", () => {
	it("reads and decodes a large prefix in pages and slices, as the one-pass read would", async () => {
		const schema = type({ n: "number" })

		for (let n = 0; n < 20_000; n++) {
			await kvSetJson(`many.${String(n)}`, n % 997 === 0 ? { n: "not a number" } : { n })
		}

		let turns = 0
		const counter = setInterval(() => {
			turns++
		}, 0)
		const sliced = await kvEntriesJsonInSlices("many.", schema)

		clearInterval(counter)

		const byKey = (rows: [string, unknown][]) => new Map(rows)

		expect(byKey(sliced)).toEqual(byKey(await kvEntriesJson("many.", schema)))
		expect(sliced).toHaveLength(20_000 - 21)
		expect(turns).toBeGreaterThan(0)
	})

	it("reads a prefix in one read from a leader on an older build without paging", async () => {
		const schema = type({ n: "number" })

		for (let n = 0; n < 3; n++) {
			await kvSetJson(`few.${String(n)}`, { n })
		}

		const unknownMethod = new Error("Cannot read properties of undefined (reading 'apply')")

		unknownMethod.name = "TypeError"

		const page = vi.spyOn(fakeApi, "kvEntriesPage").mockRejectedValue(unknownMethod)
		const entries = vi.spyOn(fakeApi, "kvEntries")

		expect(new Map(await kvEntriesJsonInSlices("few.", schema))).toEqual(
			new Map([
				["few.0", { n: 0 }],
				["few.1", { n: 1 }],
				["few.2", { n: 2 }]
			])
		)
		expect(page).toHaveBeenCalledTimes(1)
		expect(entries).toHaveBeenCalledExactlyOnceWith("few.")

		page.mockRejectedValue(new Error("timeout"))

		await expect(kvEntriesJsonInSlices("few.", schema)).rejects.toThrow("timeout")
	})

	it("keeps only what the map makes of each row", async () => {
		await kvSetJson("m.a", { n: 1, big: "x".repeat(100) })
		await kvSetJson("m.b", { n: "bad" })

		expect(await kvEntriesJsonInSlices("m.", type({ n: "number" }), (key, value) => `${key}=${String(value.n)}`)).toEqual(["m.a=1"])
	})

	it("roundtrips a bigint through kvSetJson/kvGetJson", async () => {
		const schema = type({ n: "bigint" })

		await kvSetJson("k1", { n: 123456789012345678n })

		await expect(kvGetJson("k1", schema)).resolves.toEqual({ n: 123456789012345678n })
	})

	it("drops a schema-mismatched value: null, warns, never throws", async () => {
		const warnSpy = vi.spyOn(log, "warn").mockImplementation(() => undefined)
		const schema = type({ n: "bigint" })

		await kvSetJson("k2", { n: "not-a-bigint" })

		await expect(kvGetJson("k2", schema)).resolves.toBeNull()
		expect(warnSpy).toHaveBeenCalledWith("kv", expect.stringContaining("k2"), expect.anything())
	})

	it("drops an unparseable envelope: null, warns, never throws", async () => {
		const warnSpy = vi.spyOn(log, "warn").mockImplementation(() => undefined)

		fakeStore.set("k3", "{not valid json")

		await expect(kvGetJson("k3", type({ n: "bigint" }))).resolves.toBeNull()
		expect(warnSpy).toHaveBeenCalledWith("kv", expect.stringContaining("k3"))
	})

	it("returns null for a missing key without touching the schema", async () => {
		await expect(kvGetJson("missing", type({ n: "bigint" }))).resolves.toBeNull()
	})

	it("kvClear wipes every row regardless of prefix", async () => {
		await kvSetJson("rq.v1.a", { n: 1n })
		await kvSetJson("keymap.v1.overrides", { n: 2n })
		await kvSetJson("session", { n: 3n })

		await kvClear()

		expect(fakeStore.size).toBe(0)
	})

	it("kvClear resets the database, then sweeps once with a prefix delete, never a per-row enumeration", async () => {
		await kvSetJson("rq.v1.a", { n: 1n })
		await kvSetJson("session", { n: 2n })
		const wipe = vi.spyOn(fakeApi, "wipe")
		const deletePrefix = vi.spyOn(fakeApi, "kvDeletePrefix")
		const keys = vi.spyOn(fakeApi, "kvKeys")
		const del = vi.spyOn(fakeApi, "kvDelete")

		await kvClear()

		expect(wipe).toHaveBeenCalledTimes(1)
		expect(deletePrefix.mock.calls).toEqual([[""]])
		expect(keys).not.toHaveBeenCalled()
		expect(del).not.toHaveBeenCalled()
	})

	it("kvClear sweeps with two prefix deletes through a leader whose database cannot reset", async () => {
		await kvSetJson("session", { n: 1n })
		const unknownMethod = new Error("Cannot read properties of undefined (reading 'apply')")

		unknownMethod.name = "TypeError"
		vi.spyOn(fakeApi, "wipe").mockRejectedValue(unknownMethod)
		const deletePrefix = vi.spyOn(fakeApi, "kvDeletePrefix")

		await kvClear()

		expect(deletePrefix.mock.calls).toEqual([[""], [""]])
		expect(fakeStore.size).toBe(0)
	})

	it("kvClear leaves no row behind when a straggler write lands mid-wipe (logout resurrection guard)", async () => {
		await kvSetJson("session", { n: 1n })

		// Model a write that races the wipe: an in-flight persist (e.g. an outbox flush already past its own
		// abort gate) re-adds a decrypted row right after the first pass ran. A single sweep leaves that
		// row on disk to replay next boot; the second sweep must catch it.
		const origWipe = fakeApi.wipe

		vi.spyOn(fakeApi, "wipe").mockImplementation(async () => {
			await origWipe()
			fakeStore.set("inflightChatMessages", "straggler")
		})

		await kvClear()

		expect(fakeStore.has("inflightChatMessages")).toBe(false)
		expect(fakeStore.size).toBe(0)
	})

	it("kvClear still runs the second pass when the first rejects, then rejects itself", async () => {
		await kvSetJson("session", { n: 1n })
		const timeout = new Error("db rpc timeout: wipe")

		vi.spyOn(fakeApi, "wipe").mockRejectedValue(timeout)
		const deletePrefix = vi.spyOn(fakeApi, "kvDeletePrefix")

		await expect(kvClear()).rejects.toBe(timeout)
		expect(deletePrefix).toHaveBeenCalledTimes(1)
		expect(fakeStore.size).toBe(0)
	})

	it("kvDeleteIfAt reads and deletes row by row through a leader on an older build without it", async () => {
		await kvSetJson("r.a", { at: 1 })
		await kvSetJson("r.b", { at: 2 })
		await kvSetJson("r.c", { at: 3n })
		fakeStore.set("r.d", "not json")

		expect(
			await kvDeleteIfAt([
				["r.a", 1],
				["r.b", 5],
				["r.c", 3],
				["r.d", 1]
			])
		).toEqual(["r.a"])
		expect([...fakeStore.keys()].sort()).toEqual(["r.b", "r.c", "r.d"])
	})

	it("kvClear wipes through a leader on an older build that has no kvDeletePrefix", async () => {
		await kvSetJson("rq.v1.a", { n: 1n })
		await kvSetJson("session", { n: 2n })

		// What a follower receives from such a leader: Comlink's TypeError for the unknown method, rebuilt
		// by leader forwarding as a plain Error that keeps the name.
		const unknownMethod = new Error("Cannot read properties of undefined (reading 'apply')")

		unknownMethod.name = "TypeError"

		vi.spyOn(fakeApi, "wipe").mockRejectedValue(unknownMethod)
		const deletePrefix = vi.spyOn(fakeApi, "kvDeletePrefix").mockRejectedValue(unknownMethod)
		const keys = vi.spyOn(fakeApi, "kvKeys")

		await kvClear()

		expect(fakeStore.size).toBe(0)
		expect(deletePrefix).toHaveBeenCalledTimes(1)
		expect(keys).toHaveBeenCalledTimes(2)
	})

	it("kvClear does not fall back to a per-key wipe when the prefix delete times out", async () => {
		await kvSetJson("session", { n: 1n })
		const timeout = new Error("db rpc timeout: kvDeletePrefix")

		vi.spyOn(fakeApi, "wipe").mockRejectedValue(timeout)
		vi.spyOn(fakeApi, "kvDeletePrefix").mockRejectedValue(timeout)
		const keys = vi.spyOn(fakeApi, "kvKeys")

		await expect(kvClear()).rejects.toBe(timeout)
		expect(keys).not.toHaveBeenCalled()
	})

	it("kvHas reports existence independent of schema — even a mismatched value counts as present", async () => {
		await expect(kvHas("k4")).resolves.toBe(false)

		await kvSetJson("k4", { n: "not-a-bigint" }) // kvGetJson would drop this as invalid for a bigint schema

		await expect(kvHas("k4")).resolves.toBe(true)

		await kvClear()

		await expect(kvHas("k4")).resolves.toBe(false)
	})
})
