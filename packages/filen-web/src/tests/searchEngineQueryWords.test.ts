import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { Client, CacheSearchConfig, CacheSearchSnapshot, Dir } from "@filen/sdk-rs"
import { createSearchEngine } from "@/workers/searchEngine"

// Which word the engine settles on when the longest one matches no name: a query's descriptive extra word
// ("tax return" for "Tax 2023.pdf") must not hide the names the other words find.
function snapshot(total: bigint): CacheSearchSnapshot {
	return {
		results: total === 0n ? [] : [{ parentPath: "Docs", result: { type: "dir", dir: { uuid: "d1" } as unknown as Dir } }],
		total,
		live: true
	}
}

// An engine handle that answers each filter with `totals[name]`; one search handle for the user's open
// search, one for the query.
function fakeClient(totals: Record<string, bigint>) {
	const handle = () => {
		let current = snapshot(0n)

		return {
			close: vi.fn(() => Promise.resolve(undefined)),
			free: vi.fn(),
			total: vi.fn(() => Promise.resolve(current.total)),
			setConfig: vi.fn((config: CacheSearchConfig) => {
				current = snapshot(totals[config.name ?? ""] ?? 0n)

				return Promise.resolve(undefined)
			}),
			getRange: vi.fn(() => Promise.resolve({ free: vi.fn(), initialSnapshot: vi.fn(() => current) })),
			configure: (config: CacheSearchConfig) => {
				current = snapshot(totals[config.name ?? ""] ?? 0n)
			}
		}
	}
	const handles = [handle(), handle()]
	const query = handles[1]
	const client = {
		configureCache: vi.fn(() => Promise.resolve(undefined)),
		createSearch: vi.fn((_root: string, config: CacheSearchConfig) => {
			const next = handles.shift()

			next?.configure(config)

			return next === undefined ? Promise.reject(new Error("no search")) : Promise.resolve(next)
		}),
		root: vi.fn(() => ({ uuid: "root" }))
	} as unknown as Client

	return { client, query }
}

const OPTIONS = { settleMs: 100, stallMs: 2_000, capMs: 10_000, keepMs: 60_000 }

beforeEach(() => {
	vi.useFakeTimers()
})

afterEach(() => {
	vi.useRealTimers()
})

describe("createSearchEngine — words without hits", () => {
	it("goes on to the next word when one matches no name, and keeps a word with hits over one without", async () => {
		const { client, query } = fakeClient({ return: 0n, tax: 2n })
		const engine = createSearchEngine()

		await engine.open(client, { rootUuid: null, name: "" }, () => undefined)

		const result = engine.query(client, { rootUuid: null, names: ["return", "tax"] }, OPTIONS)

		await vi.advanceTimersByTimeAsync(300)

		await expect(result).resolves.toMatchObject({ settled: true, name: "tax", snapshot: { total: 2n } })
		expect(query?.setConfig).toHaveBeenCalledTimes(1)
	})

	it("still answers with nothing when no word matches", async () => {
		const { client } = fakeClient({})
		const engine = createSearchEngine()

		await engine.open(client, { rootUuid: null, name: "" }, () => undefined)

		const result = engine.query(client, { rootUuid: null, names: ["return", "tax"] }, OPTIONS)

		await vi.advanceTimersByTimeAsync(300)

		await expect(result).resolves.toMatchObject({ settled: true, snapshot: { total: 0n } })
	})
})
