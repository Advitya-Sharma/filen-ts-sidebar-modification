// @vitest-environment happy-dom
import { act } from "react"
import { createRoot, type Root } from "react-dom/client"
import { afterEach, describe, expect, it } from "vitest"
import type { ChartRender } from "../core/xlsx/charts/render"
import { SheetChart, type SheetChartSource } from "../react/sheet/charts/index"
import { cells } from "./sheetSessionFixtures"
import { CHART_PART, KEPT, SOLD, barChart, chartFixture, openCharts } from "./sheetChartFixtures"

// A chart frame's drawing in the page: painted as SVG once, asked again on every change of the session's version and
// repainted only when the session says the chart changed; a chart that cannot be drawn shows a placeholder.

Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true)

let root: Root | null = null
let host: HTMLDivElement | null = null

afterEach(() => {
	act(() => {
		root?.unmount()
	})
	host?.remove()
	root = null
	host = null
})

async function flush(): Promise<void> {
	for (let k = 0; k < 10; k++)
		await act(async () => {
			await new Promise(resolve => setTimeout(resolve, 0))
		})
}

function mount(source: SheetChartSource, version: number): HTMLDivElement {
	host ??= document.body.appendChild(document.createElement("div"))
	root ??= createRoot(host)

	act(() => {
		root?.render(
			<SheetChart
				part={CHART_PART}
				width={480}
				height={288}
				source={source}
				version={version}
			/>
		)
	})

	return host
}

describe("chart frames in the page", () => {
	it("paints a chart, keeps it while nothing it shows changed, and repaints after its data changed", async () => {
		const { session, sheets } = await openCharts(chartFixture(barChart([SOLD, KEPT], { title: "Fruit" })))
		const answers: ChartRender["kind"][] = []
		const source: SheetChartSource = {
			render: async (part, width, height, since) => {
				const answer = await session.query("chart.render", { part, width, height, ...(since === undefined ? {} : { since }) })

				answers.push(answer.kind)

				return answer
			},
			media: partName => session.query("drawing.media", { partName })
		}
		const page = mount(source, 1)

		await flush()

		const first = page.querySelector("svg")

		expect(first).not.toBeNull()
		expect(page.querySelector(".fo-sheet-chart")?.getAttribute("data-state")).toBe("drawn")
		// The title is drawn as text.
		expect(page.textContent).toContain("Fruit")

		mount(source, 2)
		await flush()
		expect(page.querySelector("svg")).toBe(first)

		await session.command({ type: "cells.set", sheet: sheets[0] ?? 0, cells: cells([[1, 1, "40"]]) })
		mount(source, 3)
		await flush()
		expect(page.querySelector("svg")).not.toBe(first)
		expect(answers).toEqual(["drawing", "same", "drawing"])
		session.close()
	})

	it("shows a placeholder for a chart part that is not there", async () => {
		const source: SheetChartSource = {
			render: () => Promise.resolve({ kind: "placeholder", reason: "missing", stamp: "" }),
			media: () => Promise.resolve(null)
		}
		const page = mount(source, 1)

		await flush()
		expect(page.querySelector("svg")).toBeNull()
		expect(page.querySelector(".fo-sheet-chart")?.getAttribute("data-state")).toBe("placeholder")
	})
})
