// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen } from "@testing-library/react"
import "@/lib/i18n"
import type { SheetChartSource } from "@filen/office/react/sheet/charts"

// A chartsheet's chart: reached with the keyboard, and opened in the editor with Enter or Space as with a double-click.

vi.mock("@/lib/useElementSize", () => ({ useElementSize: () => ({ width: 800, height: 600 }) }))

const { ChartSheetView } = await import("@/features/spreadsheet/charts/chartSheetView")

afterEach(() => {
	cleanup()
})

const SOURCE: SheetChartSource = {
	render: () => Promise.resolve({ kind: "placeholder", reason: "missing", stamp: "" }),
	media: () => Promise.resolve(null)
}

describe("a chartsheet's chart", () => {
	it("takes focus and opens the editor on Enter, Space or a double-click", () => {
		const onEdit = vi.fn()

		render(
			<ChartSheetView
				part="xl/charts/chart1.xml"
				width={5_486_400}
				height={3_657_600}
				source={SOURCE}
				version={0}
				onEdit={onEdit}
			/>
		)

		const chart = screen.getByRole("button", { name: "Edit chart…" })

		chart.focus()
		expect(document.activeElement).toBe(chart)
		fireEvent.keyDown(chart, { key: "Enter" })
		fireEvent.keyDown(chart, { key: " " })
		fireEvent.doubleClick(chart)
		fireEvent.keyDown(chart, { key: "a" })
		expect(onEdit).toHaveBeenCalledTimes(3)
	})
})
