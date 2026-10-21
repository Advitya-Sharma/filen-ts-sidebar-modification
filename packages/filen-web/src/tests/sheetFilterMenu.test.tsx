// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react"
import type { SheetCommand, SheetFilterRange, SheetFilterValues, SheetSessionApi } from "@filen/office/core/sheet"
import type { SheetGridHandle, SheetGridModel } from "@filen/office/react/sheet"
import "@/lib/i18n"
import { preview } from "@/locales/en/preview"
import { FilterMenu } from "@/features/spreadsheet/next/data/filterMenu"
import { setSortFilterOpen, useSortFilterOpen } from "@/features/spreadsheet/next/data/sortFilterState"

// A header's filter menu: its values listed once each (searched, checked, the blanks too), OK filtering by what is
// checked, the sort commands ordering the filter's range by the column, and the conditions opening their dialog.

afterEach(() => {
	cleanup()
})

const RANGE: SheetFilterRange = { table: null, range: { startRow: 0, startCol: 0, endRow: 9, endCol: 2 }, buttons: Uint8Array.of(1, 3, 1) }

const VALUES: SheetFilterValues = {
	table: null,
	range: RANGE.range,
	header: "Fruit",
	items: [
		{ text: "apple", checked: true },
		{ text: "fig", checked: true },
		{ text: "pear", checked: false }
	],
	blank: { checked: true },
	more: false,
	criteria: { kind: "values", values: ["apple", "fig"], blank: true, dates: [] },
	unknown: false,
	kind: "text",
	sorted: null
}

const HANDLE = {
	cellRect: () => ({ x: 10, y: 20, width: 64, height: 20 }),
	onFrame: () => () => undefined,
	focus: () => undefined
} as unknown as SheetGridHandle

function show(col = 1, values: SheetFilterValues = VALUES) {
	const run = vi.fn<(command: SheetCommand) => void>()
	const query = vi.fn().mockResolvedValue(values)
	const session = { query } as unknown as SheetSessionApi
	const model = { getSnapshot: () => ({ sheet: 2, layout: null }), scale: () => 1 } as unknown as SheetGridModel
	let opened: ReturnType<typeof useSortFilterOpen> = null

	function Probe() {
		opened = useSortFilterOpen(model)

		return null
	}

	setSortFilterOpen(model, { type: "menu", row: 0, col })
	render(
		<>
			<Probe />
			<FilterMenu
				model={model}
				gridRef={{ current: HANDLE }}
				session={session}
				range={RANGE}
				row={0}
				col={col}
				rightToLeft={false}
				canEdit
				run={run}
			/>
		</>
	)

	return { run, query, opened: () => opened }
}

describe("the filter menu", () => {
	it("lists the column's values with the blanks and filters by what is checked", async () => {
		const { run, query } = show()

		await act(async () => {
			await Promise.resolve()
		})

		expect(query).toHaveBeenCalledWith("filter.values", { sheet: 2, row: 0, col: 1, limit: 10_000 })
		expect(screen.getByLabelText("pear")).toHaveProperty("checked", false)
		expect(screen.getByLabelText(preview.previewSheetFilterBlanks)).toHaveProperty("checked", true)

		fireEvent.click(screen.getByLabelText("fig"))
		fireEvent.click(screen.getByRole("button", { name: preview.previewSheetFilterApply }))

		expect(run).toHaveBeenCalledWith({
			type: "filter.column",
			sheet: 2,
			row: 0,
			col: 1,
			criteria: { kind: "values", values: ["apple"], blank: true, dates: [] }
		})
	})

	it("checks every shown value at once, and leaves out what a search hides", async () => {
		const { run } = show()

		await act(async () => {
			await Promise.resolve()
		})

		fireEvent.click(screen.getByLabelText(preview.previewSheetFilterSelectAll))
		fireEvent.click(screen.getByRole("button", { name: preview.previewSheetFilterApply }))
		// Everything checked: the column's filter goes.
		expect(run).toHaveBeenLastCalledWith({ type: "filter.column", sheet: 2, row: 0, col: 1, criteria: null })
	})

	it("searches the values and applies to the ones it shows", async () => {
		const { run } = show()

		await act(async () => {
			await Promise.resolve()
		})

		fireEvent.change(screen.getByRole("textbox", { name: preview.previewSheetFilterSearch }), { target: { value: "PE" } })
		expect(screen.queryByLabelText("apple")).toBeNull()
		fireEvent.click(screen.getByLabelText("pear"))
		fireEvent.click(screen.getByRole("button", { name: preview.previewSheetFilterApply }))
		expect(run).toHaveBeenLastCalledWith({
			type: "filter.column",
			sheet: 2,
			row: 0,
			col: 1,
			criteria: { kind: "values", values: ["pear"], blank: false, dates: [] }
		})
	})

	it("shows a search's matches from the top of the list, wherever it was scrolled", async () => {
		show()

		await act(async () => {
			await Promise.resolve()
		})

		const list = screen.getByRole("group", { name: preview.previewSheetFilterValues })

		list.scrollTop = 120
		fireEvent.scroll(list)
		fireEvent.change(screen.getByRole("textbox", { name: preview.previewSheetFilterSearch }), { target: { value: "PE" } })

		expect(list.scrollTop).toBe(0)
		expect(screen.getByLabelText("pear")).toBeTruthy()
	})

	it("sorts the filter's range by the column and clears the column's filter", async () => {
		const { run } = show()

		await act(async () => {
			await Promise.resolve()
		})

		fireEvent.click(screen.getByRole("button", { name: preview.previewSheetSortDescending }))
		expect(run).toHaveBeenLastCalledWith({
			type: "sort.apply",
			sheet: 2,
			range: RANGE.range,
			orientation: "rows",
			header: true,
			caseSensitive: false,
			keys: [{ index: 1, descending: true }]
		})

		cleanup()

		const again = show()

		await act(async () => {
			await Promise.resolve()
		})

		fireEvent.click(screen.getByRole("button", { name: preview.previewSheetFilterClearColumn.replace("{{name}}", "Fruit") }))
		expect(again.run).toHaveBeenLastCalledWith({ type: "filter.column", sheet: 2, row: 0, col: 1, criteria: null })
		// The menu closes once a command is sent.
		expect(again.opened()).toBeNull()
	})

	it("groups dates in a tree of years, months and days, a group checking all it holds", async () => {
		const dates: SheetFilterValues = {
			...VALUES,
			kind: "date",
			blank: null,
			criteria: null,
			items: [
				{ text: "1/2/2024", checked: true, date: { year: 2024, month: 1, day: 2 } },
				{ text: "1/5/2024", checked: true, date: { year: 2024, month: 1, day: 5 } },
				{ text: "3/1/2024", checked: true, date: { year: 2024, month: 3, day: 1 } },
				{ text: "12/31/2023", checked: true, date: { year: 2023, month: 12, day: 31 } }
			]
		}
		const { run } = show(1, dates)

		await act(async () => {
			await Promise.resolve()
		})

		// Collapsed: the years alone.
		expect(screen.queryByLabelText("January")).toBeNull()
		fireEvent.click(screen.getByRole("button", { name: `${preview.previewSheetFilterExpand} 2024` }))
		fireEvent.click(screen.getByRole("button", { name: `${preview.previewSheetFilterExpand} January` }))
		expect(screen.getByLabelText("02")).toHaveProperty("checked", true)

		// A month unchecked takes its days; its year is then partly checked.
		fireEvent.click(screen.getByLabelText("January"))
		expect(screen.getByLabelText("05")).toHaveProperty("checked", false)
		expect(screen.getByLabelText("2024")).toHaveProperty("indeterminate", true)

		fireEvent.click(screen.getByRole("button", { name: preview.previewSheetFilterApply }))
		expect(run).toHaveBeenLastCalledWith({
			type: "filter.column",
			sheet: 2,
			row: 0,
			col: 1,
			criteria: {
				kind: "values",
				values: [],
				blank: false,
				dates: [
					{ grouping: "month", year: 2024, month: 3 },
					{ grouping: "year", year: 2023 }
				]
			}
		})
	})

	it("disables Clear on a column without a filter", async () => {
		show(0)

		await act(async () => {
			await Promise.resolve()
		})

		expect(screen.getByRole("button", { name: preview.previewSheetFilterClearColumn.replace("{{name}}", "Fruit") })).toHaveProperty(
			"disabled",
			true
		)
	})
})
