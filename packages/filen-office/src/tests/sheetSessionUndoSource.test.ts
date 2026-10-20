import { describe, expect, it } from "vitest"
import type { SheetCommandResult } from "../core/sheet/protocol"
import { changedEntries } from "./xlsxWriteFixtures"
import { cells, firstTile, openFixture, savedBytes } from "./sheetSessionFixtures"

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason} ${result.refusal.detail ?? ""}`)
}

// An undo puts each cell back as the source cell it was: an edit and its undo save as the source.
describe("an undo", () => {
	it.each([
		[
			"keeps an empty cell the file holds through a format",
			'<sheetData><row r="3"><c r="B3"><v>1</v></c><c r="C3"/></row></sheetData>',
			"format"
		],
		[
			"gives a cleared cell back its value metadata",
			'<sheetData><row r="3"><c r="B3" t="e" vm="1"><v>#VALUE!</v></c><c r="C3" t="e" vm="2"><v>#VALUE!</v></c></row></sheetData>',
			"clear"
		]
	])("%s", async (_, body, edit) => {
		const { session, info, bytes } = await openFixture({ sheets: [{ name: "One", body }] })
		const one = info.sheets[0]?.id ?? 0

		await firstTile(session, one)
		await session.intent()

		if (edit === "format")
			ok(
				await session.command({
					type: "format.apply",
					sheet: one,
					ranges: [{ startRow: 2, startCol: 2, endRow: 2, endCol: 2 }],
					patch: { bold: true }
				})
			)
		else ok(await session.command({ type: "cells.set", sheet: one, cells: cells([[2, 1, ""]]) }))

		ok(await session.undo())

		expect(await changedEntries(bytes, await savedBytes(session))).toEqual({ changed: [], added: [], removed: [] })
		session.close()
	})
})
