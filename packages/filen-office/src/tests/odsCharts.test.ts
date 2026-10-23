import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { odsChartsRefusal } from "../core/ods/charts/write"
import { openOds } from "../core/ods/open"
import { odsSource } from "../core/ods/source"
import { odsWriter } from "../core/ods/write/writer"
import { XLSX_FEATURES } from "../core/sheet/commands"
import type { SheetSessionApi } from "../core/sheet/protocol"
import { sheetRegistry } from "../core/sheet/registry"
import { RenderStyleTable, SheetSession, sessionFeature } from "../core/sheet/session"
import type { SheetDrawings } from "../core/xlsx/drawings/feature"
import { XlsxDocument } from "../core/xlsx/edit/document"
import { CORPUS_DIR } from "./corpusDir"
import { odsBytes, table } from "./odsFixtures"
import { IMMEDIATE } from "./sheetSessionFixtures"
import { settle } from "./sheetNotesFixtures"

// ODF charts in an ODS session: each chart:chart object shows as a chart of its sheet, the chart editor's edits are
// written into the object (the rest of the package as it was), a frame on the page moves in table:shapes, and a
// chart the translation does not carry whole is refused rather than written otherwise.

const OBJECT_NS = [
	'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
	'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
	'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
	'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
	'xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"',
	'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
	'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
	'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
	'xmlns:xlink="http://www.w3.org/1999/xlink"'
].join(" ")

function chartObject(options: { readonly threeD?: boolean; readonly title?: string } = {}): string {
	return (
		`<?xml version="1.0" encoding="UTF-8"?><office:document-content ${OBJECT_NS} office:version="1.3"><office:automatic-styles>` +
		`<style:style style:name="ch3" style:family="chart"><style:chart-properties chart:stacked="false"${options.threeD === true ? ' chart:three-dimensional="true"' : ""}/></style:style>` +
		'<style:style style:name="ch5" style:family="chart"><style:chart-properties chart:display-label="true"/></style:style>' +
		'<style:style style:name="ch7" style:family="chart"><style:graphic-properties draw:fill-color="#004586"/></style:style>' +
		'</office:automatic-styles><office:body><office:chart><chart:chart svg:width="10cm" svg:height="6cm" chart:class="chart:bar">' +
		`<chart:title><text:p>${options.title ?? "Sales"}</text:p></chart:title><chart:legend chart:legend-position="end"/>` +
		'<chart:plot-area chart:style-name="ch3" table:cell-range-address="Data.A1:Data.B4">' +
		'<chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="Data.A2:Data.A4"/></chart:axis>' +
		'<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch5"><chart:grid chart:class="major"/></chart:axis>' +
		'<chart:series chart:style-name="ch7" chart:values-cell-range-address="Data.B2:Data.B4" chart:label-cell-address="Data.B1" chart:class="chart:bar"/>' +
		"</chart:plot-area></chart:chart></office:chart></office:body></office:document-content>"
	)
}

const ROWS = ["Fruit,Sold", "Apple,3", "Pear,5", "Plum,7"]
	.map(line => {
		const [label, value] = line.split(",")
		const number = Number(value)
		const second = Number.isNaN(number)
			? `<table:table-cell office:value-type="string"><text:p>${value}</text:p></table:table-cell>`
			: `<table:table-cell office:value-type="float" office:value="${number}"><text:p>${number}</text:p></table:table-cell>`

		return `<table:table-row><table:table-cell office:value-type="string"><text:p>${label}</text:p></table:table-cell>${second}</table:table-row>`
	})
	.join("")

const FRAME =
	'<table:shapes><draw:frame draw:z-index="0" draw:name="Chart 1" svg:width="10cm" svg:height="6cm" svg:x="1cm" svg:y="2cm">' +
	'<draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></table:shapes>'

function workbook(object: string): Uint8Array {
	return odsBytes({
		body: table("Data", ROWS, `${FRAME}<table:table-column table:number-columns-repeated="2"/>`),
		extra: [{ name: "Object 1/content.xml", data: new TextEncoder().encode(object) }],
		manifestEntries:
			'<manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>'
	})
}

// A session whose writer asks the charts, after each edit, whether the file can hold it (what the session's own
// edit check asks of them).
async function open(bytes: Uint8Array): Promise<{ session: SheetSessionApi; document: XlsxDocument; sheet: number }> {
	const opened = await openOds(bytesSource(bytes.slice()))

	if (!opened.ok) throw new Error(opened.message)

	const writer = odsWriter(opened.workbook, undefined)
	const styles = new RenderStyleTable()
	const document = new XlsxDocument(opened.workbook, styles, {
		kind: "ods",
		settings: { author: "Test" },
		writer: {
			...writer,
			check: candidate => {
				const why = odsChartsRefusal(candidate)

				return why === null ? null : { reason: "chartUnsupported", detail: why }
			}
		}
	})
	const session = new SheetSession(
		document,
		styles,
		{ features: own => sheetRegistry([...XLSX_FEATURES, sessionFeature(own)]) },
		{ kind: "ods", slicing: IMMEDIATE },
		() => undefined
	)
	const info = await session.start()

	await session.intent()

	const sheet = info.sheets[0]?.id ?? 0

	await session.layout(sheet)
	await settle()

	return { session, document, sheet }
}

async function saved(session: SheetSessionApi): Promise<Uint8Array> {
	const result = await session.save()

	if (!result.ok) throw new Error(`${result.reason}: ${result.message}`)

	return result.bytes
}

function entry(bytes: Uint8Array, name: string): string {
	return strFromU8(unzipSync(bytes)[name] ?? new Uint8Array(0))
}

const PART = "xl/charts/chart1.xml"

describe("ODF charts in an ODS session", () => {
	it("shows a chart:chart as a chart of its sheet, reading its cells", async () => {
		const { session, sheet } = await open(workbook(chartObject()))
		const chart = await session.query("chart.get", { sheet, part: PART })

		expect(chart?.refusal).toBeNull()
		expect(chart?.info.title).toBe("Sales")
		expect(chart?.info.legend).toBe("r")
		expect(chart?.info.plots[0]?.type).toBe("column")
		expect(chart?.info.plots[0]?.series[0]).toMatchObject({
			nameRef: "Data!$B$1",
			categories: "Data!$A$2:$A$4",
			values: "Data!$B$2:$B$4"
		})
		expect(chart?.info.axes.map(axis => axis.majorGridlines)).toEqual([false, true])

		const drawn = await session.query("chart.render", { part: PART, width: 3_600_000, height: 2_160_000 })

		expect(drawn.kind).toBe("drawing")
		session.close()
	})

	it("saves untouched as the source's bytes", async () => {
		const bytes = workbook(chartObject())
		const { session } = await open(bytes)

		expect(await saved(session)).toEqual(bytes)
		session.close()
	})

	it("writes a title, legend and axis edit into the chart object, the rest of the package as it was", async () => {
		const bytes = workbook(chartObject())
		const { session, sheet } = await open(bytes)

		for (const edit of [
			{ kind: "title", text: "Fruit sold" },
			{ kind: "legend", position: "b" },
			{ kind: "axis", axis: 2, patch: { max: 10, gridlines: false, title: "Count" } }
		] as const)
			expect((await session.command({ type: "chart.edit", sheet, part: PART, edit })).ok).toBe(true)

		const written = await saved(session)
		const object = entry(written, "Object 1/content.xml")

		expect(object).toContain("<chart:title><text:p>Fruit sold</text:p></chart:title>")
		expect(object).toContain('chart:legend-position="bottom"')
		expect(object).toContain('chart:maximum="10"')
		expect(object).not.toContain('<chart:grid chart:class="major"/>')
		const body = (text: string): string => text.slice(text.indexOf("<office:body>"))

		expect(body(entry(written, "content.xml"))).toBe(body(entry(bytes, "content.xml")))
		session.close()

		const again = await open(written)
		const chart = await again.session.query("chart.get", { sheet: again.sheet, part: PART })

		expect(chart?.info.title).toBe("Fruit sold")
		expect(chart?.info.legend).toBe("b")
		expect(chart?.info.axes[1]).toMatchObject({ max: 10, majorGridlines: false, title: "Count" })
		again.session.close()
	})

	it("writes a series' cells and a new series", async () => {
		const { session, sheet } = await open(workbook(chartObject()))

		expect(
			(
				await session.command({
					type: "chart.edit",
					sheet,
					part: PART,
					edit: { kind: "series.set", series: 0, values: "Data!$B$2:$B$3" }
				})
			).ok
		).toBe(true)
		expect(
			(
				await session.command({
					type: "chart.edit",
					sheet,
					part: PART,
					edit: { kind: "series.add", plot: 0, name: "=Data!$A$1", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" }
				})
			).ok
		).toBe(true)

		const written = await saved(session)
		const object = entry(written, "Object 1/content.xml")

		expect(object).toContain('chart:values-cell-range-address="Data.B2:Data.B3"')
		expect(object).toContain('chart:label-cell-address="Data.A1"')
		session.close()
	})

	for (const [label, axis, plotArea, values] of [
		["a row inserted above", "rows", "Data.A2:Data.B5", "Data.B3:Data.B5"],
		["a column inserted before", "cols", "Data.B1:Data.C4", "Data.C2:Data.C4"]
	] as const)
		it(`moves the plot area's range with its series' cells, ${label}`, async () => {
			const { session, sheet } = await open(workbook(chartObject()))
			const inserted = await session.command({ type: "axis.insert", sheet, axis, at: 0, count: 1 })

			expect(inserted.ok).toBe(true)

			const object = entry(await saved(session), "Object 1/content.xml")

			expect(object).toContain(`chart:values-cell-range-address="${values}"`)
			expect(object).toContain(`<chart:plot-area chart:style-name="ch3" table:cell-range-address="${plotArea}">`)
			session.close()
		})

	it("refuses an edit of a chart it cannot write whole when it is made, keeping the chart as it was", async () => {
		const bytes = workbook(chartObject({ threeD: true }))
		const { session, document, sheet } = await open(bytes)

		expect(odsSource(document.workbook)?.charts[0]?.refusal).toBe("three-dimensional")

		const result = await session.command({ type: "chart.edit", sheet, part: PART, edit: { kind: "title", text: "New" } })

		expect(result.ok).toBe(false)
		expect(odsChartsRefusal(document)).toBeNull()
		expect((await session.query("chart.get", { sheet, part: PART }))?.info.title).toBe("Sales")
		session.close()
	})

	it("refuses an edit it has no ODF form for (an automatic title)", async () => {
		const { session, sheet } = await open(workbook(chartObject()))

		expect((await session.command({ type: "chart.edit", sheet, part: PART, edit: { kind: "title", text: "" } })).ok).toBe(false)
		session.close()
	})

	it("moves a frame on the page in table:shapes", async () => {
		const { session, sheet } = await open(workbook(chartObject()))
		const drawings = (await session.layout(sheet)).extensions?.drawings as SheetDrawings | undefined
		const frame = drawings?.objects[0]

		if (frame === undefined) throw new Error("no frame")

		const moved = await session.command({
			type: "drawing.move",
			sheet,
			key: frame.key,
			stamp: frame.stamp,
			anchor: { mode: "absolute", x: 12_700 * 72, y: 12_700 * 36, cx: 12_700 * 200, cy: 12_700 * 100 }
		})

		expect(moved.ok).toBe(true)

		const content = entry(await saved(session), "content.xml")

		expect(content).toContain('svg:x="72pt"')
		expect(content).toContain('svg:width="200pt"')
		session.close()
	})
})

const DIR = join(CORPUS_DIR, "ods")
const CORPUS = [
	"lo__chart-import-basic.ods",
	"lo__tdf31231.ods",
	"lo__fdo67682-2.ods",
	"odftk__chart_positioning_import_xml_MSO15.ods",
	"govuk__FC2521_Percentage_of_owner-occupiers_aged_65__still_repaying_mortgage.ods"
].filter(name => existsSync(join(DIR, name)))

describe.skipIf(CORPUS.length === 0)("ODF charts of the corpus", () => {
	it("draws every chart object, and saves each workbook untouched as its bytes", { timeout: 300_000 }, async () => {
		let charts = 0
		const placeholders: string[] = []

		for (const name of CORPUS) {
			const bytes = new Uint8Array(readFileSync(join(DIR, name)))
			const { session, document } = await open(bytes)
			const source = odsSource(document.workbook)

			for (const chart of source?.charts ?? []) {
				charts++

				const drawn = await session.query("chart.render", { part: chart.part, width: 3_600_000, height: 2_160_000 })

				if (drawn.kind !== "drawing") placeholders.push(`${name} ${chart.object}`)
			}

			expect(await saved(session)).toEqual(bytes)
			session.close()
		}

		expect(charts).toBe(99)
		expect(placeholders).toEqual([])
	})
})
