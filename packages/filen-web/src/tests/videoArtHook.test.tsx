// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, renderHook } from "@testing-library/react"
import type { ArtLoadRequest, ArtResult } from "@/features/videos/lib/artwork"

// A card's artwork lifecycle: a download waits for the dwell on screen, keeps following the card's visibility for its
// rank, is cancelled when the card leaves, and a revoked image never shows again.

const { loader, visibility, revoke } = vi.hoisted(() => ({
	revoke: { fn: (_url: string): void => undefined },
	loader: {
		loadLocal: vi.fn<(request: ArtLoadRequest) => Promise<ArtResult | "network" | null>>(),
		load: vi.fn<(request: ArtLoadRequest) => Promise<ArtResult | null>>()
	},
	visibility: { listener: null as ((inView: boolean) => void) | null, unobserved: 0 }
}))

vi.mock("@/features/videos/lib/artwork", () => ({ artLoader: loader }))
vi.mock("@/features/videos/lib/cardObserver", () => ({
	visibleCardObserver: () => ({
		observe: (_element: Element, listener: (inView: boolean) => void) => {
			visibility.listener = listener

			return () => {
				visibility.unobserved++
				visibility.listener = null
			}
		}
	})
}))
vi.mock("@/features/videos/lib/index", () => ({
	videoLibrary: {
		titleArt: () =>
			Promise.resolve({
				candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
				fileUuid: () => "poster",
				resolveFile: () => Promise.resolve(null)
			}),
		itemArt: () => Promise.resolve(null)
	}
}))

import { useVideoArt } from "@/features/videos/hooks/useVideoArt"
import { CARD_PROBE_DWELL_MS, PROBE_RANK } from "@/features/videos/lib/probeService.logic"

const element = document.createElement("div")

function params(near: boolean) {
	return { titleId: "t", stableUuid: null, slot: "p" as const, title: "Up", near, element }
}

async function flush(): Promise<void> {
	await act(async () => {
		await vi.advanceTimersByTimeAsync(0)
	})
}

beforeEach(() => {
	vi.useFakeTimers()
	visibility.listener = null
	visibility.unobserved = 0
	URL.createObjectURL = vi.fn(() => "blob:art")
	revoke.fn = vi.fn()
	URL.revokeObjectURL = revoke.fn
})

afterEach(() => {
	vi.useRealTimers()
})

describe("useVideoArt", () => {
	it("downloads after the dwell, re-ranks with visibility while loading, and cancels when the card leaves", async () => {
		loader.loadLocal.mockResolvedValue("network")
		loader.load.mockReturnValue(new Promise(() => undefined))

		const hook = renderHook(({ near }: { near: boolean }) => useVideoArt(params(near)), { initialProps: { near: true } })

		await flush()
		act(() => {
			visibility.listener?.(true)
		})
		await act(async () => {
			await vi.advanceTimersByTimeAsync(CARD_PROBE_DWELL_MS)
		})

		const request = loader.load.mock.calls[0]?.[0]

		expect(request?.rank?.()).toBe(PROBE_RANK.visible)

		act(() => {
			visibility.listener?.(false)
		})

		expect(request?.rank?.()).toBe(PROBE_RANK.nearViewport)
		expect(loader.load).toHaveBeenCalledTimes(1)
		expect(request?.signal?.aborted).toBe(false)

		hook.rerender({ near: false })

		expect(request?.signal?.aborted).toBe(true)
		expect(visibility.unobserved).toBe(1)
	})

	it("never shows a revoked image when the card comes near again", async () => {
		loader.loadLocal.mockResolvedValueOnce({
			type: "image",
			blob: new Blob([new Uint8Array([1])]),
			source: "sidecar",
			width: 1,
			height: 1,
			fit: "cover",
			dominant: null
		})

		const hook = renderHook(({ near }: { near: boolean }) => useVideoArt(params(near)), { initialProps: { near: true } })

		await flush()

		expect(hook.result.current).toMatchObject({ kind: "image", url: "blob:art" })

		hook.rerender({ near: false })

		expect(revoke.fn).toHaveBeenCalledWith("blob:art")

		loader.loadLocal.mockReturnValueOnce(new Promise(() => undefined))
		hook.rerender({ near: true })

		expect(hook.result.current).toEqual({ kind: "none" })
	})
})
