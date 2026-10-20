import { describe, expect, it, vi } from "vitest"
import {
	SHEET_MAX_COLS,
	SHEET_MAX_ROWS,
	type SheetCommandResult,
	type SheetRenderStyle,
	type SheetSessionApi
} from "../core/sheet/protocol"
import { borderPresetEdges } from "../core/sheet/style/borders"
import { SheetStore } from "../core/xlsx/model/store"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE, buildFixture, cells, savedBytes, tileEntry } from "./sheetSessionFixtures"
import { STYLES, equalBytes, parts, type Fixture } from "./xlsxWriteFixtures"

// Formatting whole columns and rows as Excel writes it: the column's <col style> and the row's s with
// customFormat, so cells typed later, past the used area too, show the format; no empty cell is written for it.

const LAST_ROW = SHEET_MAX_ROWS - 1
const LAST_COL = SHEET_MAX_COLS - 1

async function open(
	fixture: Fixture | Uint8Array
): Promise<{ session: SheetSessionApi; sheet: number; styles: Map<number, SheetRenderStyle> }> {
	const bytes = fixture instanceof Uint8Array ? fixture : buildFixture(fixture)
	const result = await inThreadSheetHost().open(bytes.slice(), { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

	if (result.kind !== "opened") throw new Error(result.kind)

	const styles = new Map<number, SheetRenderStyle>()

	result.info.styles.styles.forEach((style, i) => styles.set(result.info.styles.from + i, style))
	held = styles

	return { session: result.session, sheet: result.info.activeSheet, styles }
}

// Styles a change brings are kept, as a client keeps them.
let held: Map<number, SheetRenderStyle> | null = null

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason}`)

	result.change.styles?.styles.forEach((style, i) => held?.set((result.change.styles?.from ?? 0) + i, style))
}

const FIXTURE: Fixture = {
	sheets: [
		{
			name: "One",
			body:
				'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' +
				'<row r="3" s="1" customFormat="1"><c r="A3"><v>3</v></c></row></sheetData>'
		}
	],
	styles: STYLES
}

// What a cell shows its look as, through a tile (styles sent with the answer).
async function look(
	open: { session: SheetSessionApi; sheet: number; styles: Map<number, SheetRenderStyle> },
	row: number,
	col: number
): Promise<SheetRenderStyle> {
	const answer = await open.session.tiles(open.sheet, Uint32Array.of(row >>> 6, col >>> 5))

	answer.styles?.styles.forEach((style, i) => open.styles.set((answer.styles?.from ?? 0) + i, style))

	const tile = answer.tiles[0]
	const entry = tile === undefined ? null : tileEntry(tile, row, col)

	if (entry !== null) return open.styles.get(entry.style) ?? {}

	const layout = await open.session.layout(open.sheet)
	const rowAt = layout.rows.indices.indexOf(row)
	const colAt = layout.cols.indices.indexOf(col)
	const style =
		rowAt >= 0 && (layout.rows.flags[rowAt] ?? 0) & 64
			? (layout.rows.styles[rowAt] ?? 0)
			: colAt >= 0
				? (layout.cols.styles[colAt] ?? 0)
				: 0

	return open.styles.get(style) ?? {}
}

describe("formatting whole columns and rows", () => {
	it("gives whole columns a column style, patching the cells they hold and adding none", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 0, startCol: 1, endRow: LAST_ROW, endCol: 2 }],
				patch: { bold: true }
			})
		)

		const sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml).toMatch(/<cols><col min="2" max="3" width="[\d.]+" style="(\d+)"\/><\/cols>/u)
		// B1 patched; no cell for the empty B2:C2 or C1.
		expect(sheetXml).toMatch(/<c r="B1" s="\d+"><v>2<\/v><\/c><\/row>/u)
		expect(sheetXml).not.toContain('r="C1"')
		// The styled row 3 shows its own style in empty cells: its B3 and C3 get the format as cells.
		expect(sheetXml).toMatch(/<c r="B3" s="\d+"\/><c r="C3" s="\d+"\/>/u)

		// A value typed far below the used area shows bold.
		ok(await session.command({ type: "cells.set", sheet, cells: cells([[5000, 1, "x"]]) }))
		expect((await look(opened, 5000, 1)).bold).toBe(true)
		expect((await look(opened, 5000, 0)).bold).toBeUndefined()
		session.close()
	})

	// ECMA-376 Part 1 §18.3.1.13: the default column's 64 px are stored as Truncate(64 / 7 * 256) / 256.
	it("keeps a formatted column at the default width, stored as Excel stores it", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened
		const width = async (): Promise<number | undefined> => {
			const layout = await session.layout(sheet)
			const at = layout.cols.indices.indexOf(2)

			return at < 0 ? layout.cols.defaultSize : layout.cols.sizes[at]
		}
		const before = await width()

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 0, startCol: 2, endRow: LAST_ROW, endCol: 2 }],
				patch: { bold: true }
			})
		)

		const saved = await savedBytes(session)
		const sheetXml = parts(saved)["xl/worksheets/sheet1.xml"] ?? ""

		expect(/<col [^>]*width="([\d.]+)"/u.exec(sheetXml)?.[1]).toBe("9.140625")
		expect(before).toBe(64)
		expect(await width()).toBe(64)
		session.close()

		const reopened = await open(saved)
		const layout = await reopened.session.layout(reopened.sheet)

		expect(layout.cols.sizes[layout.cols.indices.indexOf(2)]).toBe(64)
		reopened.session.close()
	})

	it("gives whole rows the row style Excel writes, customFormat set", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 1, startCol: 0, endRow: 1, endCol: LAST_COL }],
				patch: { italic: true }
			})
		)

		const sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml).toMatch(/<row r="2" s="\d+" customFormat="1"\/>/u)
		ok(await session.command({ type: "cells.set", sheet, cells: cells([[1, 9000, "y"]]) }))
		expect((await look(opened, 1, 9000)).italic).toBe(true)
		session.close()
	})

	it("formats every column of a select-all, and the styled rows' own styles", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 0, startCol: 0, endRow: LAST_ROW, endCol: LAST_COL }],
				patch: { fill: 0xffff00 }
			})
		)

		const sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml).toMatch(/<cols><col min="1" max="16384" width="[\d.]+" style="\d+"\/><\/cols>/u)
		expect((await look(opened, 2, 7)).fill).toBe(0xffff00)
		expect((await look(opened, 900, 900)).fill).toBe(0xffff00)
		session.close()
	})

	it("undoes to the file as it was", async () => {
		const bytes = buildFixture(FIXTURE)
		const { session, sheet } = await open(bytes)

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 0, startCol: 1, endRow: LAST_ROW, endCol: 1 }],
				patch: { bold: true }
			})
		)
		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 1, startCol: 0, endRow: 1, endCol: LAST_COL }],
				patch: { italic: true }
			})
		)
		ok(await session.undo())
		ok(await session.undo())

		expect(equalBytes(await savedBytes(session), bytes)).toBe(true)
		ok(await session.redo())
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toContain('<col min="2" max="2"')
		session.close()
	})

	it("borders whole columns through the column style, the top row's empty cells getting the top line", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 0, startCol: 3, endRow: LAST_ROW, endCol: 4 }],
				edges: { top: { style: "thin" }, left: { style: "thin" }, right: { style: "medium" }, insideVertical: { style: "dashed" } }
			})
		)

		const d1 = await look(opened, 0, 3)
		const d2000 = await look(opened, 2000, 3)
		const e2000 = await look(opened, 2000, 4)

		expect([d1.top?.style, d1.left?.style, d1.right?.style]).toEqual(["thin", "thin", "dashed"])
		expect([d2000.top, d2000.left?.style, d2000.right?.style]).toEqual([undefined, "thin", "dashed"])
		expect([e2000.left?.style, e2000.right?.style]).toEqual(["dashed", "medium"])
		session.close()
	})

	it("puts Outside on whole columns and rows, the far edge on a cell of the grid's last row or column", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 0, startCol: 3, endRow: LAST_ROW, endCol: 3 }],
				edges: borderPresetEdges("outside")
			})
		)
		ok(
			await session.command({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 4, startCol: 0, endRow: 4, endCol: LAST_COL }],
				edges: borderPresetEdges("outside")
			})
		)

		const sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""
		const d1 = await look(opened, 0, 3)
		const d3000 = await look(opened, 3000, 3)
		const dLast = await look(opened, LAST_ROW, 3)

		expect([d1.top?.style, d1.left?.style, d1.right?.style, d1.bottom]).toEqual(["thin", "thin", "thin", undefined])
		expect([d3000.top, d3000.left?.style, d3000.bottom]).toEqual([undefined, "thin", undefined])
		expect(dLast.bottom?.style).toBe("thin")
		expect((await look(opened, 4, 9000)).bottom?.style).toBe("thin")
		expect((await look(opened, 4, LAST_COL)).right?.style).toBe("thin")
		expect((await look(opened, 4, 0)).left?.style).toBe("thin")
		// One cell in the last row, none per row of the column.
		expect(sheetXml).toMatch(/<row r="1048576"[^>]*><c r="D1048576" s="\d+"\/><\/row>/u)
		session.close()
	})

	it("formats a column from its second row down through the column's style, the first row keeping its look", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 1, startCol: 2, endRow: LAST_ROW, endCol: 2 }],
				patch: { bold: true }
			})
		)
		ok(
			await session.command({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 1, startCol: 4, endRow: LAST_ROW, endCol: 4 }],
				edges: borderPresetEdges("outside")
			})
		)

		expect((await look(opened, 0, 2)).bold).toBeUndefined()
		expect((await look(opened, 1, 2)).bold).toBe(true)
		expect((await look(opened, 90_000, 2)).bold).toBe(true)
		expect((await look(opened, 0, 4)).left).toBeUndefined()
		expect((await look(opened, 1, 4)).top?.style).toBe("thin")
		expect((await look(opened, 5000, 4)).top).toBeUndefined()
		expect((await look(opened, 5000, 4)).left?.style).toBe("thin")
		session.close()
	})

	it("puts a select-all's top line on the first row's own style, not on a cell per column", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 0, startCol: 0, endRow: LAST_ROW, endCol: LAST_COL }],
				edges: borderPresetEdges("top")
			})
		)

		const sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""
		const csv = await session.export("csv", { sheet })

		expect((/<row r="1"[^>]*>(.*?)<\/row>/su.exec(sheetXml)?.[1]?.match(/<c /gu) ?? []).length).toBe(2)
		expect(sheetXml).toMatch(/<row r="1" s="\d+" customFormat="1"/u)
		expect(csv.ok ? new TextDecoder().decode(csv.bytes).split("\r\n")[0] : "").toBe("1,2")
		expect((await look(opened, 0, 0)).top?.style).toBe("thin")
		expect((await look(opened, 0, 9000)).top?.style).toBe("thin")
		expect((await look(opened, 1, 9000)).top).toBeUndefined()
		session.close()
	})

	it("draws a select-all's left line on a styled row's empty first cell", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 0, startCol: 0, endRow: LAST_ROW, endCol: LAST_COL }],
				edges: borderPresetEdges("left")
			})
		)
		ok(
			await session.command({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 0, startCol: 0, endRow: LAST_ROW, endCol: LAST_COL }],
				edges: borderPresetEdges("thickBottom")
			})
		)

		// Row 3 has a style of its own, which beats column A's.
		expect((await look(opened, 2, 0)).left?.style).toBe("thin")
		expect((await look(opened, 2, 1)).left).toBeUndefined()
		expect((await look(opened, 7, 0)).left?.style).toBe("thin")
		expect((await look(opened, LAST_ROW, 500)).bottom?.style).toBe("thick")
		expect((await look(opened, LAST_ROW - 1, 500)).bottom).toBeUndefined()
		session.close()
	})

	it("keeps a styled column's format in a whole row's empty cells, as a row's style beats a column's", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 0, startCol: 4, endRow: LAST_ROW, endCol: 4 }],
				patch: { fill: 0xff0000 }
			})
		)
		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 1, startCol: 0, endRow: 1, endCol: LAST_COL }],
				patch: { bold: true }
			})
		)

		const e2 = await look(opened, 1, 4)

		expect([e2.bold, e2.fill]).toEqual([true, 0xff0000])
		expect((await look(opened, 1, 7)).bold).toBe(true)
		expect((await look(opened, 1, 7)).fill).toBeUndefined()
		session.close()
	})

	it("gives the host a turn while it formats many cells", async () => {
		let turns = 0
		const counting = {
			sliceMs: 0,
			now: () => 0,
			pause: () => {
				turns++

				return Promise.resolve()
			}
		}
		const rows = Array.from(
			{ length: 4000 },
			(_, r) => `<row r="${String(r + 1)}"><c r="A${String(r + 1)}"><v>${String(r)}</v></c></row>`
		).join("")
		const result = await inThreadSheetHost().open(
			buildFixture({ sheets: [{ name: "One", body: `<sheetData>${rows}</sheetData>` }] }),
			{ kind: "xlsx", slicing: counting },
			() => undefined
		)

		if (result.kind !== "opened") throw new Error(result.kind)

		const { session } = result
		const sheet = result.info.activeSheet
		const ranges = [{ startRow: 0, startCol: 0, endRow: 3999, endCol: 9 }]

		turns = 0
		ok(await session.command({ type: "format.borders", sheet, ranges, edges: borderPresetEdges("all") }))
		expect(turns).toBeGreaterThan(40000 / 256 - 1)

		turns = 0
		ok(await session.command({ type: "format.apply", sheet, ranges, patch: { bold: true } }))
		expect(turns).toBeGreaterThan(40000 / 256 - 1)
		session.close()
	})

	it("adds cells to a long empty column a block of rows at a time, not one cell at a time", async () => {
		const rows = Array.from(
			{ length: 40_000 },
			(_, r) => `<row r="${String(r + 1)}"><c r="A${String(r + 1)}"><v>${String(r)}</v></c></row>`
		).join("")
		const { session, sheet } = await open(buildFixture({ sheets: [{ name: "One", body: `<sheetData>${rows}</sheetData>` }] }))
		// A cell added through setCell copies its block's arrays (and measures every block): 40,000 of them made
		// this edit take seconds on a long sheet.
		const single = vi.spyOn(SheetStore.prototype, "setCell")

		try {
			ok(
				await session.command({
					type: "format.apply",
					sheet,
					ranges: [{ startRow: 0, startCol: 1, endRow: 39_999, endCol: 1 }],
					patch: { bold: true }
				})
			)
			ok(await session.undo())
			ok(await session.redo())
			expect(single.mock.calls.length).toBeLessThan(100)
		} finally {
			single.mockRestore()
		}

		const sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml.match(/<c r="B\d+" s="\d+"\/>/gu)?.length).toBe(40_000)
		session.close()
	})

	it("gives the host turns while it weighs whole columns and rows, before it writes anything", async () => {
		let turns = 0
		const counting = {
			sliceMs: 0,
			now: () => 0,
			pause: () => {
				turns++

				return Promise.resolve()
			}
		}
		// Rows with a format of their own and cells in column A: what a select-all and a whole column look through.
		const rows = Array.from(
			{ length: 20_000 },
			(_, r) => `<row r="${String(r + 1)}" s="1" customFormat="1"><c r="A${String(r + 1)}" s="1"><v>${String(r)}</v></c></row>`
		).join("")
		const result = await inThreadSheetHost().open(
			buildFixture({ sheets: [{ name: "One", body: `<sheetData>${rows}</sheetData>` }], styles: STYLES }),
			{ kind: "xlsx", slicing: counting },
			() => undefined
		)

		if (result.kind !== "opened") throw new Error(result.kind)

		const { session } = result
		const sheet = result.info.activeSheet

		turns = 0
		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 0, startCol: 0, endRow: LAST_ROW, endCol: LAST_COL }],
				patch: { italic: true }
			})
		)
		// No cell changes (each row's style takes the format), so every turn is the weighing's.
		// A turn every few hundred cells or rows weighed: each well under the 8 ms slice.
		expect(turns).toBeGreaterThan(60)
		session.close()
	})

	it("clears formats back to the default, whole columns' and rows' own too", async () => {
		const bytes = buildFixture(FIXTURE)
		const opened = await open(bytes)
		const { session, sheet } = opened
		const column = { startRow: 0, startCol: 1, endRow: LAST_ROW, endCol: 1 }

		ok(await session.command({ type: "format.apply", sheet, ranges: [column], patch: { bold: true, fill: 0xff0000 } }))
		ok(await session.command({ type: "format.clear", sheet, ranges: [column] }))

		let sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheetXml).not.toContain("<cols>")
		expect((await look(opened, 4000, 1)).bold).toBeUndefined()

		// Row 3's own style goes with a clear of the whole row; A1 (a cell of its own) shows the default.
		ok(await session.command({ type: "format.clear", sheet, ranges: [{ startRow: 2, startCol: 0, endRow: 2, endCol: LAST_COL }] }))
		sheetXml = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""
		expect(sheetXml).toMatch(/<row r="3"><c r="A3"><v>3<\/v><\/c><\/row>/u)
		session.close()
	})

	it("clears a cell that shows its column's format to a cell of the default format", async () => {
		const opened = await open(FIXTURE)
		const { session, sheet } = opened

		ok(
			await session.command({
				type: "format.apply",
				sheet,
				ranges: [{ startRow: 0, startCol: 3, endRow: LAST_ROW, endCol: 3 }],
				patch: { italic: true }
			})
		)
		ok(await session.command({ type: "format.clear", sheet, ranges: [{ startRow: 5, startCol: 3, endRow: 5, endCol: 3 }] }))

		expect((await look(opened, 5, 3)).italic).toBeUndefined()
		expect((await look(opened, 6, 3)).italic).toBe(true)
		// A cell of a copy of the default format: the column's format stops there.
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toMatch(/<c r="D6" s="\d+"\/>/u)
		session.close()
	})
})
