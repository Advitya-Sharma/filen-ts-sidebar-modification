import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { openOds } from "../core/ods/open"
import { SHEET_MAX_ROWS } from "../core/sheet/protocol"
import type { SheetStore } from "../core/xlsx/model/store"
import { RowFlag } from "../core/xlsx/model/types"
import { inThreadSheetHost } from "../dom/sheet/host"
import { createGridAxis } from "../react/sheet/axis"
import { odsBytes, table } from "./odsFixtures"
import { IMMEDIATE } from "./sheetSessionFixtures"
import { budget, cpuMs } from "./timing"

// A few kilobytes of OpenDocument can repeat a row or a cell millions of times: what each copy costs is charged to
// the workbook's limits, the copies of one note or link share it, and past a limit the file is refused as too large.

const HIDDEN_SHEET_ROWS =
	'<table:table-row table:visibility="collapse" table:number-rows-repeated="1048576"><table:table-cell/></table:table-row>'

describe("an ODS file's repeats", () => {
	it("opens forty sheets that each hide their rows to the grid's end, holding each run as one entry", async () => {
		const content =
			'<table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>'
		const sheets = Array.from({ length: 40 }, (_, k) => table(`S${k}`, content + HIDDEN_SHEET_ROWS)).join("")
		const started = cpuMs()
		const result = await openOds(bytesSource(odsBytes({ body: sheets })))

		if (!result.ok) throw new Error(`${result.reason}: ${result.message}`)

		expect(cpuMs() - started).toBeLessThan(budget(5_000))
		expect(result.workbook.sheets).toHaveLength(40)

		for (const entry of result.workbook.sheets) {
			const model = entry.model

			if (model === null) throw new Error("no model")

			let stored = 0

			for (const block of (model.store as SheetStore).blocks) stored += block.rowCount

			// The content row alone is in the store; the hidden rows are one run.
			expect(stored).toBe(1)
			expect(model.rowRuns).toEqual([{ first: 1, last: SHEET_MAX_ROWS - 1, flags: RowFlag.HIDDEN, heightPt: Number.NaN }])
		}
	}, 60_000)

	it("shows those rows hidden in the sheet's layout, to the grid's end", async () => {
		const host = inThreadSheetHost()
		const opened = await host.open(
			odsBytes({
				body: table(
					"S",
					`<table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>${HIDDEN_SHEET_ROWS}`
				)
			}),
			{ kind: "ods", slicing: IMMEDIATE },
			() => undefined
		)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const layout = await opened.session.layout(opened.info.sheets[0]?.id ?? 0)
		const axis = createGridAxis(layout.rows, layout.rowCount, 1)

		expect(layout.rowCount).toBe(SHEET_MAX_ROWS)
		expect(layout.rows.indices.length).toBe(1)
		expect(axis.hidden(1)).toBe(true)
		expect(axis.hidden(SHEET_MAX_ROWS - 1)).toBe(true)
		expect(axis.lastShown).toBe(0)
		opened.session.close()
	}, 60_000)

	it("charges each run of rows to the workbook's cell limit, refusing past it", async () => {
		// Hidden and shown rows in turn: no two runs join.
		const rows =
			'<table:table-row table:visibility="collapse"><table:table-cell/></table:table-row><table:table-row table:style-name="ro2"><table:table-cell/></table:table-row>'.repeat(
				600
			)
		const bytes = odsBytes({
			automatic:
				'<style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="1in"/></style:style>',
			body: table("S", rows)
		})

		expect((await openOds(bytesSource(bytes))).ok).toBe(true)
		expect(await openOds(bytesSource(bytes), { readLimits: { maxCells: 1_000 } })).toMatchObject({ ok: false, reason: "tooLarge" })
	}, 60_000)

	it("refuses more sheets than the limit", async () => {
		const body = Array.from({ length: 5_000 }, (_, k) => `<table:table table:name="S${k}"/>`).join("")

		expect(await openOds(bytesSource(odsBytes({ body })))).toMatchObject({ ok: false, reason: "tooLarge" })
	}, 60_000)

	it("shows one link over a block of repeated linked cells, with one relationship for its target", async () => {
		const host = inThreadSheetHost()
		const opened = await host.open(
			odsBytes({
				body: table(
					"S",
					'<table:table-row table:number-rows-repeated="300"><table:table-cell table:number-columns-repeated="16384" office:value-type="string"><text:p><text:a xlink:href="https://example.com/">x</text:a></text:p></table:table-cell></table:table-row>'
				)
			}),
			{ kind: "ods", slicing: IMMEDIATE },
			() => undefined
		)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const sheet = opened.info.sheets[0]?.id ?? 0

		expect(await opened.session.query("link.get", { sheet, row: 0, col: 0 })).toMatchObject({ url: "https://example.com/" })
		expect(await opened.session.query("link.get", { sheet, row: 0, col: 16_383 })).toMatchObject({ url: "https://example.com/" })
		opened.session.close()
	}, 60_000)

	it("refuses a repeated note whose copies pass the limit, quickly", async () => {
		const started = cpuMs()
		const result = await openOds(
			bytesSource(
				odsBytes({
					body: table(
						"S",
						'<table:table-row table:number-rows-repeated="60"><table:table-cell table:number-columns-repeated="16384" office:value-type="float" office:value="1"><office:annotation><dc:creator>x</dc:creator><text:p>note</text:p></office:annotation><text:p>1</text:p></table:table-cell></table:table-row>'
					)
				})
			)
		)

		expect(result).toMatchObject({ ok: false, reason: "tooLarge" })
		expect(cpuMs() - started).toBeLessThan(budget(5_000))
	}, 60_000)

	it("shows every copy of a repeated note", async () => {
		const host = inThreadSheetHost()
		const opened = await host.open(
			odsBytes({
				body: table(
					"S",
					'<table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="2" office:value-type="float" office:value="1"><office:annotation><dc:creator>Ann</dc:creator><text:p>note</text:p></office:annotation><text:p>1</text:p></table:table-cell></table:table-row>'
				)
			}),
			{ kind: "ods", slicing: IMMEDIATE },
			() => undefined
		)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const sheet = opened.info.sheets[0]?.id ?? 0

		for (const [row, col] of [
			[0, 0],
			[2, 1]
		] as const)
			expect(await opened.session.query("note.get", { sheet, row, col })).toMatchObject({ text: "note", author: "Ann" })
		opened.session.close()
	}, 60_000)

	it("keeps a cell whose internal link is malformed, without that link", async () => {
		const result = await openOds(
			bytesSource(
				odsBytes({
					body: table(
						"S",
						'<table:table-row><table:table-cell office:value-type="string"><text:p><text:a xlink:href="#%E0%A4%A">x</text:a></text:p></table:table-cell><table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell></table:table-row>'
					)
				})
			)
		)

		expect(result.ok).toBe(true)
		if (result.ok) expect(result.issues.some(issue => issue.kind === "link")).toBe(true)
	}, 60_000)
})
