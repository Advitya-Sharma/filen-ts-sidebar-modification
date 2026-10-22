// @vitest-environment jsdom

import { afterEach, describe, expect, it } from "vitest"
import { act, useState, type RefObject } from "react"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import "@/lib/i18n"
import { writeXlsx } from "hucre/xlsx"
import { inThreadSheetHost } from "@filen/office/dom/sheet"
import type { SheetOpened } from "@filen/office/core/sheet"
import { createSheetGrid, useSheetGrid, type SheetGridHandle, type SheetGridModel } from "@filen/office/react/sheet"
import { requestChart } from "@/features/spreadsheet/next/charts/chartRequests"
import { useSheetCharts } from "@/features/spreadsheet/next/charts/useSheetCharts"
import type { SheetHandle } from "@/features/spreadsheet/next/hooks/useSheetSession"
import { selectedObjectKey } from "@/features/spreadsheet/next/objects/objectSelection"

// The charts feature in the sheet view over a real session on this thread: the Insert menu's request inserts a chart of
// the selected cells, selects it and opens its editor, and an edit made there reaches the chart and comes back into
// the panel; undoing the insert closes the editor.

afterEach(() => {
	cleanup()
})

async function opened(): Promise<SheetOpened> {
	const bytes = await writeXlsx({
		sheets: [
			{
				name: "Data",
				rows: [
					["Fruit", "Sold", "Kept"],
					["Apple", 3, 1],
					["Pear", 5, 2],
					["Plum", 7, 4]
				]
			}
		]
	})
	const result = await inThreadSheetHost().open(bytes, { kind: "xlsx" }, () => undefined)

	if (result.kind !== "opened") throw new Error(result.kind)

	await result.session.intent()

	return result
}

function Harness({ open, onModel }: { open: SheetOpened; onModel: (model: SheetGridModel) => void }) {
	const [model] = useState(() => {
		const created = createSheetGrid(open)

		onModel(created)

		return created
	})
	const grid = useSheetGrid(model)
	// The grid's handle: focusing it focuses the stand-in grid below.
	const [gridRef] = useState<RefObject<SheetGridHandle | null>>(() => ({
		current: {
			focus: () => {
				document.getElementById("grid")?.focus()
			}
		} as unknown as SheetGridHandle
	}))

	const charts = useSheetCharts({
		grid,
		gridRef,
		handle: { ...open, events: undefined } as unknown as SheetHandle,
		canEdit: true,
		text: false
	})

	return (
		<div>
			<div
				id="grid"
				tabIndex={0}
			/>
			{charts.panel}
		</div>
	)
}

describe("charts in the sheet view", () => {
	it("inserts a chart of the selection, opens its editor, edits it and closes when the insert is undone", async () => {
		const open = await opened()
		let model: SheetGridModel | null = null

		render(
			<Harness
				open={open}
				onModel={created => {
					model = created
				}}
			/>
		)

		const grid = model as unknown as SheetGridModel
		const sheet = open.info.activeSheet

		act(() => {
			grid.select({
				active: { row: 0, col: 0 },
				anchor: { row: 0, col: 0 },
				focus: { row: 3, col: 2 },
				ranges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 2 }]
			})
			requestChart(grid, "line")
		})

		const panel = await screen.findByRole("complementary", { name: "Chart" })

		await waitFor(() => {
			expect(screen.getAllByLabelText<HTMLInputElement>("Values").map(field => field.value)).toEqual([
				"Data!$B$2:$B$4",
				"Data!$C$2:$C$4"
			])
		})
		expect(selectedObjectKey(grid)).toEqual({ sheet, key: 0 })
		expect(panel.textContent).toContain("Line")

		const legend = grid.layout(sheet)?.extensions?.drawings?.objects[0]?.chartPart

		expect(legend).toBe("xl/charts/chart1.xml")

		const title = await screen.findByRole("switch", { name: "Show chart title" })

		fireEvent.click(title)
		await waitFor(() => {
			expect(screen.getByLabelText("Chart title")).toBeTruthy()
		})

		const field = screen.getByLabelText("Chart title")

		fireEvent.change(field, { target: { value: "Fruit" } })
		fireEvent.keyDown(field, { key: "Enter" })
		await waitFor(async () => {
			const got = await open.session.query("chart.get", { sheet, part: "xl/charts/chart1.xml" })

			expect(got?.info.title).toBe("Fruit")
		})

		// Three steps: the insert, the title shown, its text. Undone, the chart and its editor are gone.
		await act(async () => {
			await grid.undo()
			await grid.undo()
			await grid.undo()
		})
		await waitFor(() => {
			expect(screen.queryByRole("complementary", { name: "Chart" })).toBeNull()
		})
		open.session.close()
	})

	it("returns focus to the grid when the editor closes", async () => {
		const open = await opened()
		let model: SheetGridModel | null = null

		render(
			<Harness
				open={open}
				onModel={created => {
					model = created
				}}
			/>
		)

		const grid = model as unknown as SheetGridModel

		act(() => {
			grid.select({
				active: { row: 0, col: 0 },
				anchor: { row: 0, col: 0 },
				focus: { row: 3, col: 2 },
				ranges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 2 }]
			})
			requestChart(grid, "column")
		})

		const heading = await screen.findByRole("heading", { name: "Chart" })

		expect(document.activeElement).toBe(heading)
		fireEvent.click(screen.getByRole("button", { name: "Close chart editor" }))
		await waitFor(() => {
			expect(document.activeElement?.id).toBe("grid")
		})
		open.session.close()
	})
})
