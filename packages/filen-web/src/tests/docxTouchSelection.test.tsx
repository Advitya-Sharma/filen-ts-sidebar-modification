// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { render, screen } from "@testing-library/react"
import type { DocxEditorHandle } from "@filen/office/react/docx/editor"
import "@/lib/i18n"
import { DocxTouchSelection } from "@/features/preview/components/docxTouchSelection"

// The touch menu over a touch selection: a named menu (screen readers announce what it is for).

// Frames run on a timer, and nothing is resized.
beforeEach(() => {
	vi.stubGlobal("requestAnimationFrame", (run: FrameRequestCallback) =>
		setTimeout(() => {
			run(0)
		}, 0)
	)
	vi.stubGlobal("cancelAnimationFrame", (id: number) => {
		clearTimeout(id)
	})
	vi.stubGlobal(
		"ResizeObserver",
		class {
			observe = vi.fn()
			disconnect = vi.fn()
		}
	)
})

afterEach(() => {
	vi.unstubAllGlobals()
})

describe("DocxTouchSelection", () => {
	it("names its menu", async () => {
		const container = document.createElement("div")

		container.getBoundingClientRect = () =>
			({ left: 0, top: 0, width: 400, height: 400, right: 400, bottom: 400, x: 0, y: 0 }) as DOMRect
		document.body.appendChild(container)

		const handle = {
			selectionPoints: () => ({ start: { x: 100, y: 120, height: 16 }, end: { x: 160, y: 120, height: 16 } }),
			scrollElement: () => null
		} as unknown as DocxEditorHandle

		render(
			<DocxTouchSelection
				handleRef={{ current: handle }}
				container={container}
				shown
				selectionKey="1"
				actions={{ cut: () => undefined, copy: () => undefined, paste: null, selectAll: () => undefined }}
			/>
		)

		const menu = await screen.findByRole("menu", { name: "Selection actions" })

		expect(menu.querySelectorAll("[role='menuitem']")).toHaveLength(3)
		container.remove()
	})
})
