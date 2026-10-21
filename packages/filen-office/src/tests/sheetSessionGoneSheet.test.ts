import { describe, expect, it } from "vitest"
import { openFixture } from "./sheetSessionFixtures"
import type { Fixture } from "./xlsxWriteFixtures"

// Requests that reach the session after an edit took their sheet away (a delete sent in the same burst, an undo of
// the sheet's add): answered as for an empty sheet, never with a read error.

const TWO: Fixture = {
	sheets: [
		{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' },
		{ name: "Two", body: '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>' }
	]
}

const FILTERED: Fixture = {
	sheets: [
		{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' },
		{
			name: "Two",
			body:
				'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Fruit</t></is></c><c r="B1" t="inlineStr"><is><t>Count</t></is></c></row>' +
				'<row r="2"><c r="A2" t="inlineStr"><is><t>apple</t></is></c><c r="B2"><v>3</v></c></row>' +
				'<row r="3"><c r="A3" t="inlineStr"><is><t>pear</t></is></c><c r="B3"><v>7</v></c></row>' +
				'<row r="4"><c r="A4" t="inlineStr"><is><t>apple</t></is></c><c r="B4"><v>1</v></c></row></sheetData><autoFilter ref="A1:B4"/>'
		}
	]
}

const SORT = {
	type: "sort.apply",
	range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 },
	orientation: "rows",
	header: true,
	caseSensitive: false,
	keys: [{ index: 1, descending: false }]
} as const

describe("a sheet taken away while requests were on their way", () => {
	it("answers its tiles, layout and cells as an empty sheet's", async () => {
		const { session, info } = await openFixture(TWO)
		const two = info.sheets[1]!.id
		const deleted = session.command({ type: "sheet.delete", sheet: two })
		const [tiles, layout, cell] = await Promise.all([
			session.tiles(two, Uint32Array.of(0, 0)),
			session.layout(two),
			session.query("cell", { sheet: two, row: 0, col: 0 })
		])

		expect((await deleted).ok).toBe(true)
		expect(tiles.tiles).toEqual([])
		expect(layout).toMatchObject({ sheet: two, rowCount: 0, colCount: 0, pane: null })
		expect(cell.input).toBe("")
	})

	it("answers the same for a sheet whose add was undone", async () => {
		const { session } = await openFixture(TWO)
		const added = await session.command({ type: "sheet.add" })
		const sheet = added.ok ? (added.change.sheets?.at(-1)?.id ?? -1) : -1

		await session.undo()

		expect((await session.tiles(sheet, Uint32Array.of(0, 0))).tiles).toEqual([])
		expect((await session.layout(sheet)).rowCount).toBe(0)
	})

	it("answers a sheet's feature queries with null while it is gone, and its edits work once an undo brings it back", async () => {
		const { session, info } = await openFixture(FILTERED)
		const two = info.sheets[1]!.id

		expect((await session.command({ type: "sheet.delete", sheet: two })).ok).toBe(true)

		const cell = { sheet: two, row: 1, col: 0 }
		const answers = await Promise.all([
			session.query("filter.values", { ...cell, row: 0, limit: 100 }),
			session.query("note.get", cell),
			session.query("thread.get", cell),
			session.query("link.get", cell),
			session.query("validation.list", cell),
			session.query("validation.check", { ...cell, input: "x" }),
			session.query("validation.rule", cell)
		])

		expect(answers).toEqual([null, null, null, null, null, null, null])
		expect((await session.undo()).ok).toBe(true)
		expect((await session.query("filter.values", { sheet: two, row: 0, col: 0, limit: 100 }))?.items.length).toBe(2)
		expect((await session.command({ ...SORT, sheet: two })).ok).toBe(true)
	})

	it("edits a sheet whose tiles were asked for while an undo had taken it away, once a redo brings it back", async () => {
		const { session, info } = await openFixture(FILTERED)
		const duplicated = await session.command({ type: "sheet.duplicate", sheet: info.sheets[1]!.id })
		const copy = duplicated.ok ? (duplicated.change.sheets?.at(-1)?.id ?? -1) : -1

		await session.undo()
		expect((await session.tiles(copy, Uint32Array.of(0, 0))).tiles).toEqual([])
		expect((await session.layout(copy)).rowCount).toBe(0)
		await session.redo()

		expect((await session.command({ ...SORT, sheet: copy })).ok).toBe(true)
		expect(
			(
				await session.command({
					type: "cells.set",
					sheet: copy,
					cells: { rows: Uint32Array.of(5), cols: Uint16Array.of(0), inputs: ["x"] }
				})
			).ok
		).toBe(true)
		expect((await session.save()).ok).toBe(true)
	})
})
