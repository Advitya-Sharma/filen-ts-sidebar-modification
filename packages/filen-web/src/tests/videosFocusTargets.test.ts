// @vitest-environment jsdom

import { afterEach, describe, expect, it } from "vitest"
import { VIDEOS_FIRST_RUN_ACTION_ID, VIDEOS_HEADING_ID, videosFocusFallback } from "@/features/videos/lib/focusTargets"

afterEach(() => {
	document.body.replaceChildren()
})

function element(tag: string, id: string): HTMLElement {
	const node = document.createElement(tag)

	node.id = id
	document.body.append(node)

	return node
}

describe("videosFocusFallback", () => {
	it("prefers the first-run button, then the heading, else leaves focus to the dialog", () => {
		expect(videosFocusFallback()).toBe(true)

		const heading = element("h1", VIDEOS_HEADING_ID)

		expect(videosFocusFallback()).toBe(heading)

		const button = element("button", VIDEOS_FIRST_RUN_ACTION_ID)

		expect(videosFocusFallback()).toBe(button)
	})
})
