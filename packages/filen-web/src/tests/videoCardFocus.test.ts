// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, renderHook } from "@testing-library/react"
import { useCardFocus } from "@/features/videos/hooks/useCardFocus"

// A cursor move focuses its card when it mounts, which can be a page fetch later (a grid row out of range).

function card(): HTMLAnchorElement {
	const element = document.createElement("a")

	element.href = "#"
	document.body.append(element)

	return element
}

afterEach(() => {
	cleanup()
	document.body.replaceChildren()
	vi.useRealTimers()
})

describe("card focus", () => {
	it("focuses a mounted card at once, and one that mounts later when it registers", () => {
		const { result } = renderHook(() => useCardFocus<number>())
		const first = card()
		const later = card()

		result.current.register(1, first)
		result.current.focus(1, { scroll: false })

		expect(document.activeElement).toBe(first)

		result.current.focus(200, { scroll: false })
		result.current.register(200, later)

		expect(document.activeElement).toBe(later)
	})

	it("drops a request a newer one replaced, or one that waited too long", () => {
		vi.useFakeTimers({ toFake: ["performance"] })

		const { result } = renderHook(() => useCardFocus<number>())
		const first = card()
		const stale = card()
		const slow = card()

		result.current.register(1, first)
		result.current.focus(1, { scroll: false })
		result.current.focus(5, { scroll: false })
		result.current.focus(6, { scroll: false })
		result.current.register(5, stale)

		expect(document.activeElement).toBe(first)

		vi.advanceTimersByTime(5000)
		result.current.register(6, slow)

		expect(document.activeElement).toBe(first)
	})
})
