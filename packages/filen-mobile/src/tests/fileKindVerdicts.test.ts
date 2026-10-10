import { vi, describe, it, expect } from "vitest"
import { VERDICT_MAX_RECORDS, type StoredVerdict } from "@filen/shared"
import { FileKindVerdicts, type FileKindVerdictStorage } from "@/lib/fileKindVerdicts"

function memoryStorage(rows: (readonly [string, unknown])[] = []) {
	const saved = new Map<string, StoredVerdict>()
	const removed: string[] = []
	let release: (() => void) | null = null
	const gate = new Promise<void>(resolve => {
		release = resolve
	})

	const storage: FileKindVerdictStorage = {
		load: async () => {
			await gate

			return rows
		},
		save: vi.fn(async (uuid: string, value: StoredVerdict) => {
			saved.set(uuid, value)
		}),
		remove: vi.fn(async (uuid: string) => {
			removed.push(uuid)
		}),
		onError: vi.fn()
	}

	return {
		storage,
		saved,
		removed,
		loaded: async () => {
			release?.()

			await new Promise<void>(resolve => setTimeout(resolve, 0))
		}
	}
}

describe("FileKindVerdicts", () => {
	it("records, persists and reads back synchronously", () => {
		const store = new FileKindVerdicts()
		const { storage, saved } = memoryStorage()

		store.attach(storage)
		store.record("u1", "audio")

		expect(store.get("u1")).toBe("audio")
		expect(saved.get("u1")?.v).toBe("audio")
	})

	it("writes nothing for a verdict it already holds", () => {
		const store = new FileKindVerdicts()
		const { storage } = memoryStorage()

		store.attach(storage)
		store.record("u1", "audio")
		store.record("u1", "audio")

		expect(storage.save).toHaveBeenCalledTimes(1)
	})

	it("keeps the newest VERDICT_MAX_RECORDS, removing the oldest from the kv", () => {
		const store = new FileKindVerdicts()
		const { storage, removed } = memoryStorage()

		store.attach(storage)

		for (let index = 0; index <= VERDICT_MAX_RECORDS; index++) {
			store.record(`u${index}`, "audio")
		}

		expect(store.size()).toBe(VERDICT_MAX_RECORDS)
		expect(store.get("u0")).toBeUndefined()
		expect(store.get(`u${VERDICT_MAX_RECORDS}`)).toBe("audio")
		expect(removed).toEqual(["u0"])
	})

	it("re-recording a file makes it the newest", () => {
		const store = new FileKindVerdicts()

		store.record("first", "audio")

		for (let index = 1; index < VERDICT_MAX_RECORDS; index++) {
			store.record(`u${index}`, "audio")
		}

		store.record("first", "video")
		store.record("overflow", "audio")

		expect(store.get("first")).toBe("video")
		expect(store.get("u1")).toBeUndefined()
	})

	it("fills memory once from the kv, merged with this session's writes, dropping malformed and surplus rows", async () => {
		const store = new FileKindVerdicts()
		const rows: (readonly [string, unknown])[] = [
			["old", { v: "audio", at: 1 }],
			["kept", { v: "transportStream", at: 2 }],
			["bad", { v: 7, at: 3 }],
			["session", { v: "audio", at: 4 }]
		]

		for (let index = 0; index < VERDICT_MAX_RECORDS - 2; index++) {
			rows.push([`fill${index}`, { v: "video", at: 10 + index }])
		}

		const { storage, removed, loaded } = memoryStorage(rows)

		store.attach(storage)
		store.record("session", "notTransportStream")

		await loaded()

		expect(store.size()).toBe(VERDICT_MAX_RECORDS)
		expect(store.get("session")).toBe("notTransportStream")
		expect(store.get("kept")).toBe("transportStream")
		expect(store.get("old")).toBeUndefined()
		expect(removed).toEqual(expect.arrayContaining(["bad", "old"]))
	})

	it("notifies only the listeners of the file that changed, and every file on a load", async () => {
		const store = new FileKindVerdicts()
		const { storage, loaded } = memoryStorage([["a", { v: "audio", at: 1 }]])
		const onA = vi.fn()
		const onB = vi.fn()
		const onAny = vi.fn()

		store.subscribeFile("a", onA)

		const unsubscribeB = store.subscribeFile("b", onB)

		store.subscribeAny(onAny)
		store.record("b", "audio")

		expect(onA).not.toHaveBeenCalled()
		expect(onB).toHaveBeenCalledTimes(1)
		expect(onAny).toHaveBeenCalledTimes(1)

		unsubscribeB()
		store.attach(storage)

		await loaded()

		expect(onA).toHaveBeenCalledTimes(1)
		expect(onB).toHaveBeenCalledTimes(1)
		expect(onAny).toHaveBeenCalledTimes(2)
	})

	it("tells any-listeners which files a load changed, leaving out what this session already held", async () => {
		const store = new FileKindVerdicts()
		const { storage, loaded } = memoryStorage([
			["same", { v: "audio", at: 1 }],
			["fresh", { v: "video", at: 2 }]
		])
		const onAny = vi.fn()
		const onSame = vi.fn()

		store.attach(storage)
		store.record("same", "audio")
		store.subscribeAny(onAny)
		store.subscribeFile("same", onSame)

		await loaded()

		expect(onAny).toHaveBeenCalledTimes(1)
		expect(onAny).toHaveBeenCalledWith(["fresh"])
		expect(onSame).not.toHaveBeenCalled()
	})

	it("replaces a watch's reader only for a change it is affected by, or a reset", () => {
		const store = new FileKindVerdicts()
		const watch = store.watch(changed => changed.includes("mine"))
		const listener = vi.fn()
		const initial = watch.getSnapshot()

		watch.subscribe(listener)
		store.record("other", "audio")

		expect(listener).not.toHaveBeenCalled()
		expect(watch.getSnapshot()).toBe(initial)

		store.record("mine", "audio")

		expect(listener).toHaveBeenCalledTimes(1)
		expect(watch.getSnapshot()).not.toBe(initial)
		expect(watch.getSnapshot().get("mine")).toBe("audio")

		const afterMine = watch.getSnapshot()

		store.clearForLogout()

		expect(listener).toHaveBeenCalledTimes(2)
		expect(watch.getSnapshot()).not.toBe(afterMine)
	})

	it("takes up a change made between a watch's creation and its subscription", () => {
		const store = new FileKindVerdicts()
		const watch = store.watch(() => false)
		const created = watch.getSnapshot()

		store.record("mine", "audio")

		watch.subscribe(vi.fn())

		const subscribed = watch.getSnapshot()

		expect(subscribed).not.toBe(created)
		expect(subscribed).toBe(store.getReader())

		store.record("other", "audio")

		expect(watch.getSnapshot()).toBe(subscribed)
	})

	it("hands out a new reader after each change", () => {
		const store = new FileKindVerdicts()
		const before = store.getReader()

		expect(store.getReader()).toBe(before)

		store.record("a", "audio")

		expect(store.getReader()).not.toBe(before)
		expect(store.getReader().get("a")).toBe("audio")
	})

	it("forgets everything at logout and takes no write after it, a pending load included", async () => {
		const store = new FileKindVerdicts()
		const { storage, saved, loaded } = memoryStorage([["a", { v: "audio", at: 1 }]])
		const listener = vi.fn()

		store.attach(storage)
		store.record("b", "audio")
		store.subscribeFile("b", listener)
		store.clearForLogout()

		expect(store.get("b")).toBeUndefined()
		expect(listener).toHaveBeenCalledTimes(1)

		await loaded()
		store.record("c", "audio")
		store.attach(storage)

		expect(store.size()).toBe(0)
		expect(saved.has("c")).toBe(false)
	})

	it("forgets a verdict only where it holds the contradicted one", () => {
		const store = new FileKindVerdicts()
		const { storage, removed } = memoryStorage()
		const listener = vi.fn()

		store.attach(storage)
		store.record("a", "audio")
		store.record("b", "transportStream")
		store.subscribeFile("a", listener)
		store.forget("a", "audio")
		store.forget("b", "audio")

		expect(store.get("a")).toBeUndefined()
		expect(store.get("b")).toBe("transportStream")
		expect(removed).toEqual(["a"])
		expect(listener).toHaveBeenCalledTimes(1)
	})

	it("drops a loaded row a forget before the load contradicted, keeping what the session wrote since", async () => {
		const store = new FileKindVerdicts()
		const { storage, removed, loaded } = memoryStorage([
			["contradicted", { v: "audio", at: 1 }],
			["other", { v: "video", at: 2 }],
			["rewritten", { v: "audio", at: 3 }]
		])

		store.forget("contradicted", "audio")
		store.forget("other", "audio")
		store.attach(storage)
		store.forget("rewritten", "audio")
		store.record("rewritten", "audio")

		await loaded()

		expect(store.get("contradicted")).toBeUndefined()
		expect(store.get("other")).toBe("video")
		expect(store.get("rewritten")).toBe("audio")
		expect(removed).toEqual(["contradicted"])
	})

	it("keeps a row the session saved while the load was pending, even when the loaded one is malformed", async () => {
		const store = new FileKindVerdicts()
		const { storage, removed, loaded } = memoryStorage([["a", { v: 7, at: 1 }]])

		store.attach(storage)
		store.record("a", "audio")

		await loaded()

		expect(store.get("a")).toBe("audio")
		expect(removed).toEqual([])
	})

	it("reports a failed save through the storage", async () => {
		const store = new FileKindVerdicts()
		const { storage } = memoryStorage()

		storage.save = vi.fn(async () => {
			throw new Error("disk")
		})

		store.attach(storage)
		store.record("a", "audio")

		await new Promise<void>(resolve => setTimeout(resolve, 0))

		expect(storage.onError).toHaveBeenCalledWith("verdict not saved", expect.any(Error))
		expect(store.get("a")).toBe("audio")
	})
})
