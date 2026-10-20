import { describe, expect, it } from "vitest"
import v8 from "node:v8"
import { CellFlag, SHEET_MAX_CELL_TEXT, SHEET_MAX_FORMULA_TEXT, SHEET_TILE_TEXT } from "../core/sheet/protocol"
import { openTextSession } from "../core/sheet/text"
import { IMMEDIATE, SML, cells, firstTile, openFixture, savedBytes, tileEntry } from "./sheetSessionFixtures"

// What one cell holds, per "Excel specifications and limits" (support.microsoft.com): a typed entry past it is
// refused by the command, never by a later save.
describe("sheet session: entries past what a cell holds", () => {
	it.each(["1e400", "-1e400", "1E+309"])("refuses a number %s that overflows a double, and every save still succeeds", async input => {
		const { session, info } = await openFixture({ sheets: [{ name: "One", body: "<sheetData/>" }] })
		const one = info.sheets[0]?.id ?? 0

		await firstTile(session, one)
		await session.intent()

		const result = await session.command({ type: "cells.set", sheet: one, cells: cells([[0, 0, input]]) })

		expect(result.ok ? null : result.refusal).toEqual({ reason: "tooLarge", detail: "number" })
		await expect(savedBytes(session)).resolves.toBeInstanceOf(Uint8Array)
		session.close()
	})

	it("takes text up to 32,767 characters and refuses more", async () => {
		const { session, info } = await openFixture({ sheets: [{ name: "One", body: "<sheetData/>" }] })
		const one = info.sheets[0]?.id ?? 0

		await firstTile(session, one)
		await session.intent()

		const fits = await session.command({ type: "cells.set", sheet: one, cells: cells([[0, 0, "x".repeat(SHEET_MAX_CELL_TEXT)]]) })
		const over = await session.command({ type: "cells.set", sheet: one, cells: cells([[1, 0, "x".repeat(SHEET_MAX_CELL_TEXT + 1)]]) })
		// The apostrophe that keeps an entry text is not part of it.
		const quoted = await session.command({
			type: "cells.set",
			sheet: one,
			cells: cells([[2, 0, `'${"x".repeat(SHEET_MAX_CELL_TEXT)}`]])
		})

		expect([fits.ok, over.ok ? null : over.refusal, quoted.ok]).toEqual([true, { reason: "tooLarge", detail: "text" }, true])
		session.close()
	})

	it("takes a formula up to 8,192 characters and refuses more", async () => {
		const { session, info } = await openFixture({ sheets: [{ name: "One", body: "<sheetData/>" }] })
		const one = info.sheets[0]?.id ?? 0
		const formula = (length: number): string => `=1${"+1".repeat((length - 1) / 2)}`

		await firstTile(session, one)
		await session.intent()

		const fits = await session.command({ type: "cells.set", sheet: one, cells: cells([[0, 0, formula(SHEET_MAX_FORMULA_TEXT - 1)]]) })
		const over = await session.command({ type: "cells.set", sheet: one, cells: cells([[1, 0, formula(SHEET_MAX_FORMULA_TEXT + 1)]]) })

		expect([fits.ok, over.ok ? null : over.refusal]).toEqual([true, { reason: "tooLarge", detail: "formula" }])
		session.close()
	})
})

describe("sheet session: a file's text longer than a cell paints", () => {
	it("is cut in its tile, never inside a surrogate pair, and the cell query gives all of it", async () => {
		const long = "x".repeat(40_000)
		const paired = `${"y".repeat(SHEET_TILE_TEXT - 1)}\u{1F600}tail`
		const { session, info } = await openFixture({
			sheets: [
				{
					name: "One",
					body: `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${long}</t></is></c><c r="B1" t="inlineStr"><is><t>${paired}</t></is></c><c r="C1" t="inlineStr"><is><t>short</t></is></c></row></sheetData>`
				}
			]
		})
		const one = info.sheets[0]?.id ?? 0
		const tile = await firstTile(session, one)
		const [a1, b1, c1] = [tileEntry(tile, 0, 0), tileEntry(tile, 0, 1), tileEntry(tile, 0, 2)]

		expect([a1?.text.length, b1?.text.length, c1?.text]).toEqual([SHEET_TILE_TEXT, SHEET_TILE_TEXT - 1, "short"])
		expect([(a1?.flags ?? 0) & CellFlag.CLIPPED, (b1?.flags ?? 0) & CellFlag.CLIPPED, (c1?.flags ?? 0) & CellFlag.CLIPPED]).toEqual([
			CellFlag.CLIPPED,
			CellFlag.CLIPPED,
			0
		])
		expect((await session.query("cell", { sheet: one, row: 0, col: 0 })).input).toBe(long)
		session.close()
	})

	it("keeps a tile of the longest texts a cell holds to a few MB", { timeout: 60_000 }, async () => {
		// 2,048 distinct shared strings of 32,760 characters: a zip of about 100 KB.
		const filler = "a".repeat(32_755)
		const strings: string[] = [`<sst xmlns="${SML}" count="2048" uniqueCount="2048">`]
		let rows = ""

		for (let i = 0; i < 2048; i++) strings.push(`<si><t>${String(i).padStart(5, "0")}${filler}</t></si>`)

		strings.push("</sst>")

		for (let r = 0; r < 64; r++) {
			rows += `<row r="${String(r + 1)}">`

			for (let c = 0; c < 32; c++) rows += `<c t="s"><v>${String(r * 32 + c)}</v></c>`

			rows += "</row>"
		}

		const { session } = await openFixture({
			sheets: [{ name: "One", body: `<sheetData>${rows}</sheetData>` }],
			sharedStrings: strings.join("")
		})
		const { tiles } = await session.tiles(0, Uint32Array.of(0, 0))

		session.close()
		expect(v8.serialize(tiles).byteLength).toBeLessThan(8 * 1024 * 1024)
	})

	it("is cut in a text file's tile too", async () => {
		const field = "x".repeat(200_000)
		const opened = await openTextSession(new TextEncoder().encode(`a,${field}\n`), { kind: "csv", slicing: IMMEDIATE }, () => undefined)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const tile = (await opened.session.tiles(0, Uint32Array.of(0, 0))).tiles[0]

		if (tile === undefined) throw new Error("no tile")

		const b1 = tileEntry(tile, 0, 1)

		expect([b1?.text.length, (b1?.flags ?? 0) & CellFlag.CLIPPED]).toEqual([SHEET_TILE_TEXT, CellFlag.CLIPPED])
		expect((await opened.session.query("cell", { sheet: 0, row: 0, col: 1 })).input).toBe(field)
		opened.session.close()
	})
})
