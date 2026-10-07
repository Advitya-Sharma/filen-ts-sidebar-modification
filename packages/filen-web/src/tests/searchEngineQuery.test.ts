import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { Client, CacheSearchSnapshot, CacheStatusMessage, Dir } from "@filen/sdk-rs"
import { createSearchEngine, type SearchPush } from "@/workers/searchEngine"

function snapshot(overrides: Partial<CacheSearchSnapshot> = {}): CacheSearchSnapshot {
	return { results: [], total: 0n, live: true, ...overrides }
}

type SnapshotListener = (s: CacheSearchSnapshot) => void

function fakeHandle(initial: CacheSearchSnapshot) {
	let listener: SnapshotListener | undefined
	const searchWindow = { free: vi.fn(), initialSnapshot: vi.fn(() => initial) }
	const search = {
		close: vi.fn(() => Promise.resolve(undefined)),
		free: vi.fn(),
		setConfig: vi.fn(() => Promise.resolve(undefined)),
		getRange: vi.fn((_start: bigint, _end: bigint, next: SnapshotListener) => {
			listener = next

			return Promise.resolve(searchWindow)
		})
	}

	return { search, searchWindow, deliver: (s: CacheSearchSnapshot) => listener?.(s) }
}

function fakeClient(searches: ReturnType<typeof fakeHandle>["search"][]) {
	let status: ((messages: CacheStatusMessage[]) => void) | undefined
	const createSearch = vi.fn(() => {
		const next = searches.shift()

		return next === undefined ? Promise.reject(new Error("no search")) : Promise.resolve(next)
	})
	const client = {
		configureCache: vi.fn((_path: string, listener: (messages: CacheStatusMessage[]) => void) => {
			status = listener

			return Promise.resolve(undefined)
		}),
		createSearch,
		root: vi.fn(() => ({ uuid: "root" }))
	} as unknown as Client

	return { client, createSearch, status: (messages: CacheStatusMessage[]) => status?.(messages) }
}

const hit = { parentPath: "Docs", result: { type: "dir" as const, dir: { uuid: "d1" } as unknown as Dir } }
const OPTIONS = { settleMs: 100, capMs: 1_000 }

beforeEach(() => {
	vi.useFakeTimers()
})

afterEach(() => {
	vi.useRealTimers()
})

describe("createSearchEngine — query", () => {
	it("answers busy while the user's own search holds the engine, without touching it", async () => {
		const user = fakeHandle(snapshot())
		const { client, createSearch } = fakeClient([user.search])
		const engine = createSearchEngine()
		const pushes: SearchPush[] = []

		await engine.open(client, { rootUuid: null, name: "doc" }, p => pushes.push(p))

		await expect(engine.query(client, { rootUuid: null, name: "x" }, OPTIONS)).resolves.toEqual({ status: "busy" })
		expect(createSearch).toHaveBeenCalledTimes(1)
		expect(user.search.close).not.toHaveBeenCalled()
		await expect(engine.setName("docs")).resolves.toBe(true)
	})

	it("answers once the results have been still for settleMs, then closes its own handle", async () => {
		const own = fakeHandle(snapshot())
		const { client } = fakeClient([own.search])
		const engine = createSearchEngine()
		const result = engine.query(client, { rootUuid: null, name: "doc" }, OPTIONS)

		await vi.advanceTimersByTimeAsync(50)
		own.deliver(snapshot({ results: [hit], total: 1n }))
		await vi.advanceTimersByTimeAsync(99)
		expect(own.search.close).not.toHaveBeenCalled()
		await vi.advanceTimersByTimeAsync(1)

		await expect(result).resolves.toEqual({
			status: "done",
			snapshot: { hits: [{ parentPath: "Docs", item: hit.result.dir }], total: 1n, live: true },
			settled: true
		})
		expect(own.search.close).toHaveBeenCalledTimes(1)
		await expect(engine.setName("again")).resolves.toBe(false)
	})

	it("waits out a resync and gives up at capMs with settled false", async () => {
		const own = fakeHandle(snapshot())
		const { client, status } = fakeClient([own.search])
		const engine = createSearchEngine()
		const result = engine.query(client, { rootUuid: null, name: "doc" }, OPTIONS)

		await vi.advanceTimersByTimeAsync(0)
		status([{ type: "resyncProgress", progress: { type: "started", roots: ["root"] } } as unknown as CacheStatusMessage])
		await vi.advanceTimersByTimeAsync(1_000)

		await expect(result).resolves.toMatchObject({ status: "done", settled: false })
	})

	it("never refilters its handle for the user: setName answers false so the user's keystroke reopens", async () => {
		const own = fakeHandle(snapshot())
		const { client } = fakeClient([own.search])
		const engine = createSearchEngine()
		const result = engine.query(client, { rootUuid: null, name: "doc" }, OPTIONS)

		await vi.advanceTimersByTimeAsync(0)
		await expect(engine.setName("mine")).resolves.toBe(false)
		expect(own.search.setConfig).not.toHaveBeenCalled()
		await vi.advanceTimersByTimeAsync(100)
		await result
	})

	it("yields to the user's open: answers what it had and leaves the user's routing and handle alone", async () => {
		const own = fakeHandle(snapshot({ results: [hit], total: 1n }))
		const user = fakeHandle(snapshot())
		const { client } = fakeClient([own.search, user.search])
		const engine = createSearchEngine()
		const result = engine.query(client, { rootUuid: null, name: "doc" }, OPTIONS)

		await vi.advanceTimersByTimeAsync(0)

		const pushes: SearchPush[] = []
		const userOpen = engine.open(client, { rootUuid: null, name: "mine" }, p => pushes.push(p))

		await expect(result).resolves.toMatchObject({ status: "done", settled: false, snapshot: { total: 1n } })
		await userOpen
		expect(own.search.close).toHaveBeenCalledTimes(1)
		expect(user.search.close).not.toHaveBeenCalled()
		await expect(engine.setName("mine2")).resolves.toBe(true)
	})
})
