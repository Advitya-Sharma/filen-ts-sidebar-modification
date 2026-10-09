// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest"
import { renderHook } from "@testing-library/react"
import type { MediaSessionHandler } from "@/lib/media/mediaSessionOwner"

const handlers = vi.hoisted(() => new Map<string, MediaSessionHandler>())

vi.mock("@/lib/media/mediaSessionOwner", () => ({
	mediaSessionOwners: {
		create: () => ({
			acquire: () => undefined,
			release: () => undefined,
			setMetadata: () => undefined,
			setPlaybackState: () => undefined,
			setPositionState: () => undefined,
			setActionHandler: (action: string, handler: MediaSessionHandler) => {
				handlers.set(action, handler)
			}
		})
	}
}))

import { useStageMediaSession } from "@/components/media/video/useStageMediaSession"

describe("useStageMediaSession", () => {
	it("counts every OS media control as the user touching the player (still watching)", () => {
		const video = document.createElement("video")
		const onInteraction = vi.fn()
		const onNext = vi.fn()

		video.pause = vi.fn()
		renderHook(() => {
			useStageMediaSession({ video, info: { title: "T", artist: "", artworkUrl: null }, onNext, onPrevious: null, onInteraction })
		})

		handlers.get("pause")?.({ action: "pause" })
		handlers.get("seekforward")?.({ action: "seekforward" })
		handlers.get("nexttrack")?.({ action: "nexttrack" })

		expect(onInteraction).toHaveBeenCalledTimes(3)
		expect(onNext).toHaveBeenCalledTimes(1)
		expect(handlers.get("previoustrack")).toBeNull()
	})
})
