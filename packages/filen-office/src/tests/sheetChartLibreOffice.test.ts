import { spawn } from "node:child_process"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { createInterface } from "node:readline"
import { unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { readPlotArea } from "../core/drawingml/chart/model"
import { D, drawingReader } from "../core/drawingml/reader"
import type { SheetCommand, SheetSessionApi } from "../core/sheet/protocol"
import { parseXml } from "../core/xml/index"
import { officeEnv } from "./interop/soffice"
import { sheetOffices, sheetProfile, unoEnv, unoOffice } from "./oracles/sheet/libreOffice"
import { cells } from "./sheetSessionFixtures"
import { CHART_PART, KEPT, SOLD, barChart, chartFixture, openCharts } from "./sheetChartFixtures"

// Our charts read by LibreOffice (over UNO, uno/chart.py): each saved after an edit (data, type, title, legend, axis
// bounds, series, a chart inserted), it opens, exports as PDF, and LibreOffice draws the values our refreshed caches
// hold, with the types, title, legend and bounds we wrote. Opt-in: FILEN_OFFICE_SOFFICE=<soffice>;
// FILEN_OFFICE_LO_PROFILE names where LibreOffice's profile goes.

const UNO = unoOffice(sheetOffices())
const SCRIPT = join(import.meta.dirname, "oracles", "sheet", "uno", "chart.py")

interface LoSeries {
	readonly label: string | null
	readonly values: readonly (number | string | null)[] | null
	readonly xValues: readonly (number | string | null)[] | null
	readonly ranges: readonly string[]
}

interface LoChart {
	readonly sheet: string
	readonly types: readonly string[]
	readonly title: string | null
	readonly legend: string | null
	readonly series: readonly LoSeries[]
	readonly axes: readonly {
		readonly dimension: number
		readonly index: number
		readonly min: number | null
		readonly max: number | null
	}[]
}

interface LoResult {
	readonly ok: boolean
	readonly error?: string
	readonly charts?: readonly LoChart[]
	readonly pdf?: number
}

interface Scenario {
	readonly name: string
	readonly commands: (sheets: readonly number[]) => SheetCommand[]
	readonly check: (chart: readonly LoChart[], parts: Record<string, Uint8Array>) => void
}

// Each series' values as our saved caches hold them, by chart part, in drawing order.
function cachedValues(bytes: Uint8Array): (number | null)[][] {
	const reader = drawingReader(parseXml(bytes))
	const chart = reader.child(reader.document.documentElement(), D.C, "chart")
	const area = readPlotArea(reader, reader.child(chart, D.C, "plotArea"))

	return area.plots.flatMap(plot =>
		plot.series.map(series => [...(series.values.numbers ?? [])].map(value => (Number.isNaN(value) ? null : value)))
	)
}

const edit = (sheet: number, change: Extract<SheetCommand, { type: "chart.edit" }>["edit"]): SheetCommand => ({
	type: "chart.edit",
	sheet,
	part: CHART_PART,
	edit: change
})

const SCENARIOS: readonly Scenario[] = [
	{
		name: "data edited",
		commands: ([data = 0]) => [
			{
				type: "cells.set",
				sheet: data,
				cells: cells([
					[2, 1, "50"],
					[3, 2, "-2.5"]
				])
			}
		],
		check: (charts, parts) => {
			const ours = cachedValues(parts[CHART_PART] ?? new Uint8Array(0))

			expect(ours).toEqual([
				[3, 50, 7],
				[1, 2, -2.5]
			])
			expect(charts[0]?.series.map(series => series.values)).toEqual(ours)
		}
	},
	{
		name: "type, title, legend and bounds",
		commands: ([data = 0]) => [
			edit(data, { kind: "type", plot: 0, type: "line" }),
			edit(data, { kind: "title", text: "Fruit sales" }),
			edit(data, { kind: "legend", position: "b" }),
			edit(data, { kind: "axis", axis: 222, patch: { min: 0, max: 20 } })
		],
		check: charts => {
			const chart = charts[0]

			expect(chart?.types).toEqual(["com.sun.star.chart2.LineChartType"])
			expect(chart?.title).toBe("Fruit sales")
			// LegendPosition PAGE_END: below the diagram.
			expect(chart?.legend).toBe("PAGE_END")
			expect(chart?.axes.find(axis => axis.dimension === 1)).toMatchObject({ min: 0, max: 20 })
		}
	},
	{
		name: "series added and moved",
		commands: ([data = 0]) => [
			edit(data, { kind: "series.add", plot: 0, name: "=Other!A1", categories: null, values: "Data!C2:C4" }),
			edit(data, { kind: "series.move", series: 2, to: 0 })
		],
		check: (charts, parts) => {
			expect(charts[0]?.series.map(series => series.values)).toEqual(cachedValues(parts[CHART_PART] ?? new Uint8Array(0)))
			expect(charts[0]?.series.length).toBe(3)
		}
	},
	...(["column", "bar", "line", "area", "pie", "doughnut", "scatter"] as const).map((type): Scenario => ({
		name: `${type} inserted`,
		commands: ([data = 0]) => [
			{
				type: "chart.insert",
				sheet: data,
				range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 },
				chart: type,
				placement: { from: { row: 6, col: 0, rowOff: 0, colOff: 0 } }
			}
		],
		check: (charts, parts) => {
			const inserted = charts[1]
			const ours = cachedValues(parts["xl/charts/chart2.xml"] ?? new Uint8Array(0))

			expect(inserted?.types.length).toBe(1)
			expect(inserted?.series.map(series => series.values)).toEqual(ours)
			expect(inserted?.types[0]).toBe(
				type === "column" || type === "bar"
					? "com.sun.star.chart2.ColumnChartType"
					: type === "line"
						? "com.sun.star.chart2.LineChartType"
						: type === "area"
							? "com.sun.star.chart2.AreaChartType"
							: type === "scatter"
								? "com.sun.star.chart2.ScatterChartType"
								: "com.sun.star.chart2.PieChartType"
			)
		}
	}))
]

async function saved(session: SheetSessionApi): Promise<Uint8Array> {
	const result = await session.save()

	if (!result.ok) throw new Error(result.message)

	return result.bytes
}

describe.skipIf(UNO === undefined)("charts against LibreOffice", () => {
	it("opens, exports and draws what we wrote", { timeout: 600_000 }, async () => {
		const office = UNO

		if (office?.resources === undefined) return

		const scratch = mkdtempSync(join(tmpdir(), "filen-office-chart-lo-"))
		const profile = sheetProfile(process.env["FILEN_OFFICE_LO_PROFILE"] ?? scratch, "chart")
		const pipe = `filen-office-chart-${String(process.pid)}-${String(Date.now())}`
		const soffice = spawn(
			office.path,
			[
				`-env:UserInstallation=file://${profile}`,
				"--headless",
				"--norestore",
				"--nologo",
				"--nodefault",
				"--nolockcheck",
				`--accept=pipe,name=${pipe};urp;`
			],
			{ detached: true, stdio: ["ignore", "ignore", "ignore"], env: officeEnv() }
		)
		const python = spawn(join(office.resources, "python"), ["-I", SCRIPT, office.resources, pipe], {
			stdio: ["pipe", "pipe", "inherit"],
			env: unoEnv(office.resources)
		})
		const lines = createInterface({ input: python.stdout })[Symbol.asyncIterator]()
		const read = async (): Promise<string> => {
			const line = await lines.next()

			if (line.done === true) throw new Error("the UNO runner exited")

			return line.value
		}

		try {
			expect(JSON.parse(await read())).toEqual({ ready: true })

			for (const [index, scenario] of SCENARIOS.entries()) {
				const { session, sheets } = await openCharts(chartFixture(barChart([SOLD, KEPT])))

				for (const command of scenario.commands(sheets)) {
					const result = await session.command(command)

					expect(result.ok, `${scenario.name}: ${JSON.stringify(result.ok ? null : result.refusal)}`).toBe(true)
				}

				const bytes = await saved(session)
				const input = join(scratch, `input-${String(index)}.xlsx`)
				const pdf = join(scratch, `output-${String(index)}.pdf`)

				session.close()
				writeFileSync(input, bytes)
				python.stdin.write(`${JSON.stringify({ id: index, input, pdf })}\n`)

				const theirs = JSON.parse(await read()) as LoResult

				expect(theirs.ok, `${scenario.name}: ${theirs.error ?? ""}`).toBe(true)
				expect(theirs.pdf ?? 0, scenario.name).toBeGreaterThan(0)
				scenario.check(theirs.charts ?? [], unzipSync(bytes))
			}
		} finally {
			python.stdin.end()

			try {
				if (soffice.pid !== undefined) process.kill(-soffice.pid, "SIGKILL")
			} catch {
				// Already gone.
			}

			rmSync(scratch, { recursive: true, force: true })
		}
	})
})
