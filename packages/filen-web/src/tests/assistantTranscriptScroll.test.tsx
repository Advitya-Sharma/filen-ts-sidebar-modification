// @vitest-environment jsdom

import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest"
import { act, render } from "@testing-library/react"
import type { TranscriptRow } from "@/features/assistant/components/assistant.logic"

// The transcript pinned to the newest message while rows append: every row the virtualizer mounts is measured
// from its ref callback, inside React's commit, and moves the scroll position to keep the end in view. That
// update must not reach flushSync while React is committing (a development warning, and no flush there).
// jsdom has no layout, so the scroller and the rows report faked sizes; rows measure above their estimate.

vi.mock("@/features/assistant/components/replyMarkdown", () => ({ AssistantReplyMarkdown: () => null }))

import "@/lib/i18n"
import { AssistantTranscript } from "@/features/assistant/components/assistantTranscript"

const CLIENT_HEIGHT = 300
const ROW_HEIGHT = 120
const geometry = { scrollHeight: CLIENT_HEIGHT }
const originalDescriptors = new Map<string, PropertyDescriptor | undefined>()
const FAKED = {
	scrollHeight: (el: HTMLElement) => (el.getAttribute("role") === "region" ? geometry.scrollHeight : 0),
	clientHeight: (el: HTMLElement) => (el.getAttribute("role") === "region" ? CLIENT_HEIGHT : 0),
	offsetHeight: (el: HTMLElement) =>
		el.getAttribute("role") === "region" ? CLIENT_HEIGHT : el.hasAttribute("data-index") ? ROW_HEIGHT : 0
}

beforeAll(() => {
	for (const [name, read] of Object.entries(FAKED)) {
		originalDescriptors.set(name, Object.getOwnPropertyDescriptor(HTMLElement.prototype, name))
		Object.defineProperty(HTMLElement.prototype, name, {
			configurable: true,
			get(this: HTMLElement) {
				return read(this)
			}
		})
	}
})

afterAll(() => {
	for (const [name, descriptor] of originalDescriptors) {
		if (descriptor === undefined) {
			Reflect.deleteProperty(HTMLElement.prototype, name)
		} else {
			Object.defineProperty(HTMLElement.prototype, name, descriptor)
		}
	}
})

afterEach(() => {
	vi.restoreAllMocks()
})

function rows(count: number): TranscriptRow[] {
	return Array.from({ length: count }, (_, i) => ({
		kind: "user",
		id: `u${String(i)}`,
		turnId: `t${String(i)}`,
		text: `message ${String(i)}`
	}))
}

describe("AssistantTranscript pinned to the end", () => {
	it("measures appended rows without calling flushSync inside React's commit", () => {
		const errors = vi.spyOn(console, "error").mockImplementation(() => undefined)
		const { rerender, container } = render(<AssistantTranscript rows={rows(1)} />)

		for (let count = 2; count <= 6; count++) {
			geometry.scrollHeight = Math.max(CLIENT_HEIGHT, count * ROW_HEIGHT)
			act(() => {
				rerender(<AssistantTranscript rows={rows(count)} />)
			})
		}

		expect(errors.mock.calls.flat().filter(arg => typeof arg === "string" && arg.includes("flushSync"))).toEqual([])

		// The update the measurements raised still committed with them, before the browser could paint: every
		// row sits at its measured offset (8px top padding), none at the estimate it mounted with.
		const placed = [...container.querySelectorAll<HTMLElement>("[data-index]")].map(row => [
			Number(row.getAttribute("data-index")),
			row.style.transform
		])

		expect(placed.length).toBeGreaterThan(1)

		for (const [index, transform] of placed) {
			expect(transform).toBe(`translateY(${String(8 + Number(index) * ROW_HEIGHT)}px)`)
		}
	})
})
