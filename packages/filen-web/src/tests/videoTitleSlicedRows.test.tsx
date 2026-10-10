// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { act, renderHook } from "@testing-library/react"
import { FIRST_SLICE_ROWS, SLICE_ROWS, useRowsInSlices } from "@/features/videos/hooks/useRowsInSlices"
import { restoreScroll } from "@/features/videos/lib/scrollMemory"

afterEach(() => {
	vi.useRealTimers()
	vi.unstubAllGlobals()
})

describe("a long season's rows", () => {
	it("mounts the first screen at once and the rest in idle slices, in order", () => {
		vi.useFakeTimers()

		const { result } = renderHook(() => useRowsInSlices(1000))

		expect(result.current).toBe(FIRST_SLICE_ROWS)

		act(() => {
			vi.advanceTimersByTime(20)
		})
		expect(result.current).toBe(FIRST_SLICE_ROWS + SLICE_ROWS)

		for (let i = 0; i < 20; i++) {
			act(() => {
				vi.advanceTimersByTime(20)
			})
		}

		expect(result.current).toBe(1000)
	})

	it("mounts a short season whole", () => {
		const { result } = renderHook(() => useRowsInSlices(12))

		expect(result.current).toBe(12)
	})
})

describe("restoring a title page's scroll position", () => {
	function scroller(height: { value: number }) {
		let top = 0
		const listeners = new Map<string, () => void>()
		const element = {
			children: [{}],
			get scrollTop() {
				return top
			},
			set scrollTop(value: number) {
				top = Math.min(value, Math.max(0, height.value - 500))
			},
			addEventListener: (type: string, listener: () => void) => listeners.set(type, listener),
			removeEventListener: (type: string) => listeners.delete(type)
		}

		return { element: element as unknown as HTMLElement, listeners }
	}

	function stubResizeObserver(): { grow: () => void } {
		const callbacks: (() => void)[] = []

		vi.stubGlobal(
			"ResizeObserver",
			class {
				public constructor(callback: () => void) {
					callbacks.push(callback)
				}

				public readonly observe = vi.fn()

				public disconnect(): void {
					callbacks.length = 0
				}
			}
		)

		return {
			grow: () => {
				for (const callback of [...callbacks]) {
					callback()
				}
			}
		}
	}

	it("follows the rows as they mount until the offset is reached", () => {
		const resize = stubResizeObserver()
		const height = { value: 2000 }
		const { element, listeners } = scroller(height)

		restoreScroll(element, 6000)
		expect(element.scrollTop).toBe(1500)

		height.value = 4000
		resize.grow()
		expect(element.scrollTop).toBe(3500)

		height.value = 9000
		resize.grow()
		expect(element.scrollTop).toBe(6000)
		expect(listeners.size).toBe(0)
	})

	it("stops when the user scrolls first", () => {
		const resize = stubResizeObserver()
		const height = { value: 2000 }
		const { element, listeners } = scroller(height)

		restoreScroll(element, 6000)
		listeners.get("wheel")?.()
		element.scrollTop = 100
		height.value = 9000
		resize.grow()

		expect(element.scrollTop).toBe(100)
	})
})
