import { afterEach, describe, expect, it, vi } from "vitest"
import { TILE_ROWS, type SheetRenderStyle, type SheetStyleAppend, type SheetTile } from "../core/sheet/protocol"
import type { Slicing } from "../core/slicing"
import { DocumentFormulaHost } from "../core/sheet/cf/host"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE, STYLES_WITH, cells, tileEntry } from "./sheetCfFixtures"
import { buildFixture } from "./xlsxWriteFixtures"

// What conditional formats compute over their ranges (top and bottom ranks, averages, duplicates, the scales' ends),
// kept current through edits: an edit reads again only the rows around it, one made while the statistics are
// computed ahead is not lost, a range of many parts costs what it covers, and a rule that cannot be calculated
// leaves what it would decide undecided.

const RED = '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
const GREEN = '<dxf><font><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>'
const FILL = '<dxf><fill><patternFill><bgColor rgb="FF0000FF"/></patternFill></fill></dxf>'

function numbers(count: number, value: (i: number) => number): string {
	const rows: string[] = []

	for (let i = 0; i < count; i++) rows.push(`<row r="${String(i + 1)}"><c r="A${String(i + 1)}"><v>${String(value(i))}</v></c></row>`)

	return rows.join("")
}

async function open(body: string, slicing: Slicing = IMMEDIATE) {
	const host = inThreadSheetHost()
	const result = await host.open(
		buildFixture({
			sheets: [{ name: "One", body }],
			styles: STYLES_WITH([RED, GREEN, FILL])
		}),
		{ kind: "xlsx", slicing },
		() => undefined
	)

	if (result.kind !== "opened") throw new Error(result.kind)

	const styles: SheetRenderStyle[] = []
	const add = (append: SheetStyleAppend | undefined): void => {
		append?.styles.forEach((style, i) => {
			styles[append.from + i] = style
		})
	}

	add(result.info.styles)

	const sheet = result.info.sheets[0]?.id ?? 0

	await result.session.layout(sheet)

	const look = async (row: number, col = 0): Promise<SheetRenderStyle> => {
		const { tiles, styles: appended } = await result.session.tiles(sheet, Uint32Array.of(Math.floor(row / TILE_ROWS), 0))

		add(appended)

		return styles[tileEntry(tiles[0] as SheetTile, row, col)?.style ?? 0] ?? {}
	}
	const set = async (row: number, input: string): Promise<void> => {
		const done = await result.session.command({
			type: "cells.set",
			sheet,
			cells: cells([[row, 0, input]])
		})

		if (!done.ok) throw new Error(done.refusal.reason)

		// Looks first interned by the change's patches.
		add(done.change.styles)
	}

	return { session: result.session, sheet, look, set }
}

const settle = (): Promise<void> => new Promise(resolve => setTimeout(resolve, 0))

// Cells the rules' host reads while `run` runs.
async function cellsRead(run: () => Promise<void>): Promise<number> {
	let read = 0
	const original = DocumentFormulaHost.prototype.forEach
	const spy = vi.spyOn(DocumentFormulaHost.prototype, "forEach").mockImplementation(function (this: DocumentFormulaHost, ref, visit) {
		original.call(this, ref, (row, col, value) => {
			read++
			visit(row, col, value)
		})
	})

	try {
		await run()
	} finally {
		spy.mockRestore()
	}

	return read
}

afterEach(() => {
	vi.restoreAllMocks()
})

describe("conditional format statistics", () => {
	it("ranks again after an edit from the rows around it, as a fresh read ranks", async () => {
		const count = 60_000
		const { session, look, set } = await open(
			`<sheetData>${numbers(count, i => i + 1)}</sheetData>` +
				`<conditionalFormatting sqref="A1:A${String(count)}"><cfRule type="top10" dxfId="0" priority="1" rank="1"/>` +
				'<cfRule type="duplicateValues" dxfId="1" priority="2"/><cfRule type="aboveAverage" dxfId="2" priority="3" aboveAverage="0" stdDev="1"/></conditionalFormatting>'
		)

		await settle()
		expect((await look(count - 1)).fill).toBe(0xffc7ce)
		expect((await look(4)).fill).toBe(0x0000ff)

		const read = await cellsRead(async () => {
			await set(0, "1000000000")
		})

		// One band of rows, not the whole range.
		expect(read).toBeLessThanOrEqual(8_192)
		expect((await look(0)).fill).toBe(0xffc7ce)
		expect((await look(count - 1)).fill).toBeUndefined()

		// A1 now repeats A5; the mean and deviation moved with A1.
		await set(0, "5")
		expect((await look(0)).fill).toBe(0xc6efce)
		expect((await look(4)).fill).toBe(0xc6efce)
		expect((await look(count - 1)).fill).toBe(0xffc7ce)
		// Below the average by one standard deviation: (n+1)/2 - sqrt((n^2-1)/12) for 1..n, A1 nearly unchanged.
		const limit = (count + 1) / 2 - Math.sqrt((count * count - 1) / 12)

		expect((await look(Math.floor(limit) - 2)).fill).toBe(0x0000ff)
		expect((await look(Math.ceil(limit) + 1)).fill).toBeUndefined()

		// Cleared: no longer a duplicate, the largest stays.
		await set(4, "")
		expect((await look(0)).fill).toBe(0x0000ff)
		session.close()
	}, 30_000)

	it("keeps an edit made while a rule's statistics are computed ahead", async () => {
		const count = 100_000
		// Every turn ends after a band of rows; the first turn's pause once the sheet is read is held.
		let hold = false
		let release: (() => void) | null = null
		const slicing: Slicing = {
			sliceMs: 0,
			now: () => performance.now(),
			pause: () => {
				if (!hold) return settle()

				hold = false

				return new Promise<void>(resolve => {
					release = resolve
				})
			}
		}
		const { session, look, set } = await open(
			`<sheetData>${numbers(count, i => i + 1)}</sheetData>` +
				`<conditionalFormatting sqref="A1:A${String(count)}"><cfRule type="top10" dxfId="0" priority="1" rank="1"/></conditionalFormatting>`,
			slicing
		)

		hold = true

		while (release === null) await settle()

		// While the top-one statistics are computed ahead, past the first cell, it becomes the largest.
		await set(0, "1000000000")
		;(release as () => void)()

		for (let k = 0; k < 100; k++) await settle()

		expect((await look(count - 1)).fill).toBeUndefined()
		expect((await look(0)).fill).toBe(0xffc7ce)
		session.close()
	}, 60_000)

	it("reads a range of many parts for what it covers, and edits it from the parts around the edit", async () => {
		// Every other row of A1:A40000 as its own part.
		const parts = Array.from({ length: 20_000 }, (_, i) => `A${String(2 * i + 1)}`).join(" ")
		const { session, look, set } = await open(
			`<sheetData>${numbers(40_000, i => i + 1)}</sheetData>` +
				`<conditionalFormatting sqref="${parts}"><cfRule type="top10" dxfId="0" priority="1" rank="1"/>` +
				'<cfRule type="cellIs" dxfId="2" priority="2" operator="lessThan"><formula>4</formula></cfRule></conditionalFormatting>'
		)

		await settle()
		// A39999 is the largest the parts cover (A40000 is outside them).
		expect((await look(39_998)).fill).toBe(0xffc7ce)
		expect((await look(39_999)).fill).toBeUndefined()
		expect((await look(0)).fill).toBe(0x0000ff)
		expect((await look(1)).fill).toBeUndefined()
		expect((await look(2)).fill).toBe(0x0000ff)

		// An edit outside the parts changes nothing they rank.
		await set(39_999, "100000")
		expect((await look(39_998)).fill).toBe(0xffc7ce)

		const read = await cellsRead(async () => {
			await set(10, "100000")
		})

		// The parts in the edit's band of rows, not all 20,000.
		expect(read).toBeLessThanOrEqual(4_096)
		expect((await look(10)).fill).toBe(0xffc7ce)
		expect((await look(39_998)).fill).toBeUndefined()
		session.close()
	})

	it("reads a large edit's bands again in turns, its tiles meanwhile showing the last known ranks", async () => {
		const count = 120_000
		let pauses = 0
		const slicing: Slicing = {
			sliceMs: 0,
			now: () => performance.now(),
			pause: () => {
				pauses++

				return settle()
			}
		}
		const { session, look, sheet } = await open(
			`<sheetData>${numbers(count, i => i + 1)}</sheetData>` +
				`<conditionalFormatting sqref="A1:A${String(count)}"><cfRule type="top10" dxfId="0" priority="1" rank="1"/><cfRule type="duplicateValues" dxfId="1" priority="2"/></conditionalFormatting>`,
			slicing
		)

		for (let k = 0; k < 200; k++) await settle()
		expect((await look(count - 1)).fill).toBe(0xffc7ce)

		// A paste over most of the range: the command and the next tile read no more than an edit's bands.
		const pasted = 100_000
		pauses = 0
		const rows = Uint32Array.from({ length: pasted }, (_, i) => i)
		const read = await cellsRead(async () => {
			const done = await session.command({
				type: "cells.set",
				sheet,
				cells: {
					rows,
					cols: new Uint16Array(pasted),
					inputs: Array.from({ length: pasted }, (_, i) => (i === 500 ? "1000000000" : "1"))
				}
			})

			if (!done.ok) throw new Error(done.refusal.reason)

			// Shown as last known until the bands are read: the old largest value, and the 1s not yet repeated.
			expect((await look(count - 1)).fill).toBe(0xffc7ce)
			expect((await look(1)).fill).toBeUndefined()
		})

		expect(read).toBeLessThanOrEqual(2 * 4_096)

		for (let k = 0; k < 400 && (await look(count - 1)).fill !== undefined; k++) await settle()

		expect(pauses).toBeGreaterThan(10)
		expect((await look(500)).fill).toBe(0xffc7ce)
		expect((await look(count - 1)).fill).toBeUndefined()
		expect((await look(1)).fill).toBe(0xc6efce)

		// Rows inserted above the range: the rewritten formats show the ranks they had, then their own.
		const inserted = await cellsRead(async () => {
			const done = await session.command({
				type: "axis.insert",
				sheet,
				axis: "rows",
				at: 0,
				count: 1
			})

			if (!done.ok) throw new Error(done.refusal.reason)

			expect((await look(501)).fill).toBe(0xffc7ce)
		})

		expect(inserted).toBeLessThanOrEqual(2 * 4_096)

		for (let k = 0; k < 400; k++) await settle()

		expect((await look(501)).fill).toBe(0xffc7ce)
		expect((await look(500)).fill).toBe(0xc6efce)
		expect((await look(count)).fill).toBeUndefined()
		session.close()
	}, 60_000)

	it("leaves undecided what a rule it cannot calculate would decide", async () => {
		const { session, look } = await open(
			`<sheetData>${numbers(2, i => 10 * (i + 1))}</sheetData>` +
				// The first rule's function is not calculated here: below it, A1 takes no fill and stops; A2 keeps the
				// second rule's font but not its fill, which the first rule's fill would decide.
				'<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1"><formula>XLOOKUP(1,$A$1:$A$2,$A$1:$A$2)=1</formula></cfRule>' +
				'<cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>' +
				'<conditionalFormatting sqref="A2"><cfRule type="expression" dxfId="2" priority="3"><formula>XLOOKUP(1,$A$1:$A$2,$A$1:$A$2)=1</formula></cfRule>' +
				'<cfRule type="cellIs" dxfId="1" priority="4" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>'
		)

		await settle()
		expect(await look(0)).not.toHaveProperty("fill")
		expect(await look(0)).not.toHaveProperty("color", 0x006100)

		const second = await look(1)

		expect(second.color).toBe(0x006100)
		expect(second.fill).toBeUndefined()
		session.close()
	})
})
