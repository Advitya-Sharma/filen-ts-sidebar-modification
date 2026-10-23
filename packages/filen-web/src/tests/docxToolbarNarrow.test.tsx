// @vitest-environment jsdom

import { describe, expect, it } from "vitest"
import { render, screen } from "@testing-library/react"
import { DOCX_VIEWER_CAPABILITIES, type DocxViewerState } from "@filen/office/dom/docx"
import "@/lib/i18n"
import { DocxToolbar } from "@/features/preview/components/docxToolbar"

// On a phone the toolbar row is too narrow for the page and zoom buttons beside the markup menu: the row is a
// size container and, by its width, gives up the fit-width button and then the page buttons, the markup and find
// buttons keeping the room they need (the right cell is never squeezed under its content).

const STATE: DocxViewerState = {
	status: "ready",
	capabilities: DOCX_VIEWER_CAPABILITIES,
	page: 0,
	pageCount: 5,
	pageCountFinal: true,
	zoom: 0.5,
	zoomMode: "fitWidth",
	find: { open: false, query: "", caseSensitive: false, count: 0, capped: false, active: -1 }
}

function toolbar(markup: boolean): HTMLElement {
	render(
		<DocxToolbar
			state={STATE}
			dispatch={() => undefined}
			markup={markup ? "simple" : null}
			onMarkupChange={() => undefined}
			viewOnly={[]}
			findFocus={0}
			onFindClosed={() => undefined}
		/>
	)

	const row = screen.getByRole("button", { name: /fit/i }).closest(".grid")

	if (!(row instanceof HTMLElement)) throw new Error("No toolbar row")

	return row
}

describe("DocxToolbar at phone widths", () => {
	it("is a size container whose right cell keeps the width of its buttons", () => {
		const row = toolbar(true)
		expect(row.className).toContain("@container/docx-toolbar")
		expect(row.lastElementChild?.className).not.toContain("min-w-0")
	})

	it("gives up fit-width, then the page buttons, by the row's width when the markup menu is shown", () => {
		toolbar(true)

		expect(screen.getByRole("button", { name: /fit/i }).className).toContain("@max-[24rem]/docx-toolbar:hidden")
		expect(screen.getByRole("button", { name: /previous page/i }).className).toContain("@max-[21.5rem]/docx-toolbar:hidden")
		expect(screen.getByRole("button", { name: /next page/i }).className).toContain("@max-[21.5rem]/docx-toolbar:hidden")
	})

	it("keeps the page buttons without the markup menu, giving up only fit-width below 360 px", () => {
		toolbar(false)

		expect(screen.getByRole("button", { name: /fit/i }).className).toContain("@max-[21.5rem]/docx-toolbar:hidden")
		expect(screen.getByRole("button", { name: /previous page/i }).className).not.toContain(":hidden")
	})
})
