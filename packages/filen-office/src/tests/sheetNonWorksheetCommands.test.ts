import { describe, expect, it } from "vitest"
import type { SheetCommand, SheetCommandType } from "../core/sheet/protocol"
import { changeOf } from "./sheetNotesFixtures"
import { buildFixture } from "./sheetSessionFixtures"
import { KEPT, SHEET_CHART_PART, SOLD, barChart, chartFixture, openCharts, withChartsheet } from "./sheetChartFixtures"

// A chartsheet has no cells: every command on its cells, rows, columns, merges, sorts, filters, panes or formats is
// refused there by the session, never applied to nothing and then dropped by the save. Sheet operations and the
// chartsheet's own chart still run.

const A1 = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }

const CELL_COMMANDS: readonly ((sheet: number) => SheetCommand)[] = [
	sheet => ({ type: "cells.set", sheet, cells: { rows: Uint32Array.of(0), cols: Uint16Array.of(0), inputs: ["x"] } }),
	sheet => ({ type: "cells.clear", sheet, ranges: [A1], what: "all" }),
	sheet => ({ type: "cells.shift", sheet, range: A1, edit: "insert", axis: "rows" }),
	sheet => ({ type: "format.apply", sheet, ranges: [A1], patch: { bold: true } }),
	sheet => ({ type: "format.borders", sheet, ranges: [A1], edges: { top: null } }),
	sheet => ({ type: "format.clear", sheet, ranges: [A1] }),
	sheet => ({ type: "axis.insert", sheet, axis: "rows", at: 0, count: 1 }),
	sheet => ({ type: "axis.delete", sheet, axis: "cols", at: 0, count: 1 }),
	sheet => ({ type: "axis.resize", sheet, axis: "cols", indices: Uint32Array.of(0), sizes: Float32Array.of(200) }),
	sheet => ({ type: "axis.hide", sheet, axis: "rows", spans: [{ start: 0, end: 0 }], hidden: true }),
	sheet => ({ type: "merge.cells", sheet, ranges: [A1], mode: "center" }),
	sheet => ({ type: "merge.unmerge", sheet, ranges: [A1] }),
	sheet => ({
		type: "sort.apply",
		sheet,
		range: A1,
		orientation: "rows",
		header: false,
		caseSensitive: false,
		keys: [{ index: 0, descending: false }]
	}),
	sheet => ({ type: "filter.toggle", sheet, range: A1, active: { row: 0, col: 0 } }),
	sheet => ({ type: "view.freeze", sheet, rows: 1, cols: 1 }),
	sheet => ({ type: "clipboard.paste", sheet, at: { row: 0, col: 0 }, text: "x" }),
	sheet => ({ type: "fill.copy", sheet, ranges: [A1], direction: "down" }),
	sheet => ({ type: "validation.set", sheet, ranges: [A1], rule: null }),
	sheet => ({ type: "note.set", sheet, row: 0, col: 0, text: "hello" }),
	sheet => ({ type: "link.set", sheet, ranges: [A1], target: { kind: "url", url: "https://example.com" }, row: 0, col: 0, text: "x" })
]

async function chartsheet() {
	const opened = await openCharts(withChartsheet(buildFixture(chartFixture(barChart([SOLD]))), barChart([SOLD, KEPT])))

	return { ...opened, graph: opened.sheets[2] ?? 0 }
}

describe("commands on a sheet that is not a worksheet", () => {
	for (const make of CELL_COMMANDS) {
		const type = make(0).type

		it(`refuses ${type} on a chartsheet and leaves the workbook clean`, async () => {
			const { session, graph } = await chartsheet()
			const result = await session.command(make(graph))

			const state = result.ok ? result.change.state : result.state

			expect(result.ok ? null : result.refusal).toEqual({ reason: "unsupported", detail: "sheet" })
			expect(state.dirty).toBe(false)
			expect(state.canUndo).toBe(false)
			session.close()
		})
	}

	it("reports every cell command as refused there for menus and toolbars", async () => {
		const { session, graph, sheets } = await chartsheet()
		const types = CELL_COMMANDS.map(make => make(0).type) as SheetCommandType[]
		const onGraph = await session.query("commandState", { sheet: graph, ranges: [A1], types })
		const onCells = await session.query("commandState", { sheet: sheets[0] ?? 0, ranges: [A1], types: ["format.apply", "cells.set"] })

		expect(onGraph.enabled.every(enabled => !enabled)).toBe(true)
		expect(new Set(onGraph.reasons)).toEqual(new Set(["unsupported"]))
		expect(onCells.enabled).toEqual([true, true])
		session.close()
	})

	it("still renames the chartsheet and edits its chart", async () => {
		const { session, graph } = await chartsheet()

		changeOf(await session.command({ type: "sheet.rename", sheet: graph, name: "Plot" }))
		changeOf(
			await session.command({ type: "chart.edit", sheet: graph, part: SHEET_CHART_PART, edit: { kind: "title", text: "Own sheet" } })
		)
		session.close()
	})
})
