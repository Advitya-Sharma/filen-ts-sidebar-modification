import { describe, expect, it } from "vitest"
import type { SheetSessionApi, SheetTile } from "../core/sheet/protocol"
import { defineSheetFeature } from "../core/sheet/registry"
import type { XlsxDocument } from "../core/xlsx/edit/document"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE, buildFixture, cells } from "./sheetSessionFixtures"
import { STYLES, type Fixture } from "./xlsxWriteFixtures"

async function open(fixture: Fixture): Promise<{ session: SheetSessionApi; sheet: number }> {
	const result = await inThreadSheetHost().open(buildFixture(fixture), { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

	if (result.kind !== "opened") throw new Error(`expected an open, got ${result.kind}`)

	return { session: result.session, sheet: result.info.activeSheet }
}

// The General values of a tile, by "row,col".
function generalValues(tile: SheetTile): Map<string, number> {
	const out = new Map<string, number>()
	const layer = tile.layers?.general

	layer?.entries.forEach((entry, i) => {
		const position = tile.positions[entry] ?? 0

		out.set(`${String(position >> 5)},${String(position & 31)}`, layer.values[i] ?? Number.NaN)
	})

	return out
}

describe("General values in tiles", () => {
	it("carries the stored value of each number shown in plain General, formula results included", async () => {
		const { session, sheet } = await open({
			sheets: [
				{
					name: "One",
					body:
						"<sheetData>" +
						'<row r="1"><c r="A1"><v>2.44999999999</v></c><c r="B1" s="1"><v>45000</v></c><c r="C1" t="s"><v>0</v></c></row>' +
						'<row r="2"><c r="A2"><f>A1/3</f><v>0.81666666666333</v></c><c r="B2" t="b"><v>1</v></c></row>' +
						"</sheetData>"
				}
			],
			styles: STYLES,
			sharedStrings:
				'<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>x</t></si></sst>'
		})
		const tile = (await session.tiles(sheet, Uint32Array.of(0, 0))).tiles[0] as SheetTile
		const values = generalValues(tile)

		expect(values.get("0,0")).toBe(2.44999999999)
		expect(values.get("1,0")).toBe(0.81666666666333)
		// A date, text and a boolean are not General numbers.
		expect(values.has("0,1")).toBe(false)
		expect(values.has("0,2")).toBe(false)
		expect(values.has("1,1")).toBe(false)
		expect(tile.layers?.general?.decimal).toBe(".")
		session.close()
	})

	it("follows an edit and the session's locale", async () => {
		const { session, sheet } = await open({ sheets: [{ name: "One", body: "<sheetData/>" }] })

		session.configure({ locale: { decimal: ",", group: ".", argument: ";", dateOrder: "dmy" } })

		const result = await session.command({ type: "cells.set", sheet, cells: cells([[0, 0, "3.25"]]) })

		if (!result.ok) throw new Error(result.refusal.reason)

		const tile =
			result.change.patches.find(patch => patch.tileRow === 0 && patch.tileCol === 0) ??
			(await session.tiles(sheet, Uint32Array.of(0, 0))).tiles[0]

		expect(generalValues(tile as SheetTile).get("0,0")).toBe(3.25)
		expect(tile?.layers?.general?.decimal).toBe(",")
		session.close()
	})

	it("keeps a feature's layer on its entries when a feature adds an entry before others", async () => {
		const marks = defineSheetFeature<XlsxDocument>({
			id: "marks",
			tiles: (_context, tile) => {
				const late = tile.find(1, 1)
				// A blank cell before every stored one: entries are put in position order once the tile is built.
				const early = tile.entry(0, 0)

				tile.layer("general", { entries: Uint16Array.of(early, late), values: Float64Array.of(1, 2), decimal: "." })
			}
		})
		const result = await inThreadSheetHost({ features: [marks] }).open(
			buildFixture({ sheets: [{ name: "One", body: '<sheetData><row r="2"><c r="B2"><v>5</v></c></row></sheetData>' }] }),
			{ kind: "xlsx", slicing: IMMEDIATE },
			() => undefined
		)

		if (result.kind !== "opened") throw new Error(result.kind)

		const tile = (await result.session.tiles(result.info.activeSheet, Uint32Array.of(0, 0))).tiles[0] as SheetTile

		expect([...generalValues(tile)]).toEqual([
			["0,0", 1],
			["1,1", 2]
		])
		result.session.close()
	})
})
