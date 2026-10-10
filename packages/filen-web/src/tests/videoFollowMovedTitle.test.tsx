// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { renderHook } from "@testing-library/react"
import type { ReactElement } from "react"

const mocks = vi.hoisted(() => ({
	navigate: vi.fn(() => Promise.resolve()),
	whenTitleKnown: vi.fn<(id: string, then: () => void) => () => void>()
}))

vi.mock("@tanstack/react-router", async importOriginal => ({
	...(await importOriginal<typeof import("@tanstack/react-router")>()),
	useNavigate: () => mocks.navigate
}))
vi.mock("@/features/videos/lib/followTitle", () => ({ whenTitleKnown: mocks.whenTitleKnown }))
vi.mock("@/features/videos/screens/videoTitle", () => ({ VideoTitleScreen: () => null }))

import { useFollowMovedTitle } from "@/features/videos/hooks/useFollowMovedTitle"
import { Route } from "@/routes/_app/videos.title.$uuid"

afterEach(() => {
	vi.clearAllMocks()
})

describe("useFollowMovedTitle", () => {
	it("follows the moved item to its title once the index has it", () => {
		const cancel = vi.fn()

		mocks.whenTitleKnown.mockReturnValue(cancel)

		const { result, unmount } = renderHook(() => useFollowMovedTitle("a"))

		result.current("b")
		mocks.whenTitleKnown.mock.calls[0]?.[1]()
		expect(mocks.navigate).toHaveBeenCalledWith({ to: "/videos/title/$uuid", params: { uuid: "b" }, replace: true })

		// Its own title is no move.
		result.current("a")
		expect(mocks.whenTitleKnown).toHaveBeenCalledTimes(1)
		unmount()
	})

	it("drops a pending follow when the page goes, and starts none for a save that resolves after it went", () => {
		const cancel = vi.fn()

		mocks.whenTitleKnown.mockReturnValue(cancel)

		const { result, unmount } = renderHook(() => useFollowMovedTitle("a"))
		const followLater = result.current

		result.current("b")
		unmount()
		expect(cancel).toHaveBeenCalledTimes(1)

		// The dialog closed while saving; the save resolves on another page.
		followLater("c")
		expect(mocks.whenTitleKnown).toHaveBeenCalledTimes(1)
		expect(mocks.navigate).not.toHaveBeenCalled()
	})
})

describe("the title route", () => {
	it("keys the screen by title, so another title never inherits this one's state", () => {
		vi.spyOn(Route, "useParams").mockReturnValue({ uuid: "b" })

		const Page = Route.options.component as () => ReactElement
		const element = Page()

		expect(element.key).toBe("b")
	})
})
