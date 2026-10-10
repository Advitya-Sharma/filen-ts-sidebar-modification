import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

const mocks = vi.hoisted(() => ({
	catalogTitle: vi.fn(),
	listeners: new Set<() => void>(),
	version: { value: 1 },
	warn: vi.fn()
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi: { videoCatalogTitle: mocks.catalogTitle } }))
vi.mock("@/lib/log", () => ({ log: { warn: mocks.warn, info: vi.fn(), error: vi.fn(), debug: vi.fn() } }))
vi.mock("@/features/videos/lib/index", () => ({
	videoLibrary: {
		getSnapshot: () => ({ version: mocks.version.value }),
		subscribe: (listener: () => void) => {
			mocks.listeners.add(listener)

			return () => mocks.listeners.delete(listener)
		}
	}
}))

import { FOLLOW_TIMEOUT_MS, whenTitleKnown } from "@/features/videos/lib/followTitle"

function catalogChanged(): void {
	mocks.version.value++

	for (const listener of [...mocks.listeners]) {
		listener()
	}
}

async function flush(): Promise<void> {
	await vi.advanceTimersByTimeAsync(0)
}

beforeEach(() => {
	vi.useFakeTimers()
	mocks.listeners.clear()
	mocks.catalogTitle.mockReset()
	mocks.warn.mockReset()
})

afterEach(() => {
	vi.useRealTimers()
})

describe("following a corrected item to its title", () => {
	it("goes once the catalog has the title, asking again only when the catalog changes", async () => {
		const then = vi.fn()

		mocks.catalogTitle.mockResolvedValueOnce(null).mockResolvedValue({ id: "t2" })
		whenTitleKnown("t2", then)
		await flush()
		expect(then).not.toHaveBeenCalled()

		catalogChanged()
		await flush()
		expect(then).toHaveBeenCalledTimes(1)
		expect(mocks.catalogTitle).toHaveBeenCalledTimes(2)
		expect(mocks.listeners.size).toBe(0)
	})

	it("never goes after the page that asked cancelled (it was left)", async () => {
		const then = vi.fn()
		let answer: (title: unknown) => void = () => undefined

		mocks.catalogTitle.mockReturnValue(
			new Promise(resolve => {
				answer = resolve
			})
		)

		const cancel = whenTitleKnown("t2", then)

		cancel()
		answer({ id: "t2" })
		await flush()

		expect(then).not.toHaveBeenCalled()
		expect(mocks.listeners.size).toBe(0)
	})

	it("logs a failed lookup instead of leaving it unhandled, and gives up after the wait", async () => {
		const then = vi.fn()

		mocks.catalogTitle.mockRejectedValue(new Error("worker gone"))
		whenTitleKnown("t2", then)
		await flush()

		expect(mocks.warn).toHaveBeenCalledTimes(1)

		await vi.advanceTimersByTimeAsync(FOLLOW_TIMEOUT_MS)
		expect(mocks.listeners.size).toBe(0)
		expect(then).not.toHaveBeenCalled()
	})
})
