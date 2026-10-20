import { describe, expect, it } from "vitest"
import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetCommand, type SheetCommandResult, type SheetSessionApi } from "../core/sheet/protocol"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE, buildFixture, savedBytes } from "./sheetSessionFixtures"
import { STYLES, parts, type Fixture } from "./xlsxWriteFixtures"

// cells.clear "formats" and "all" clear formats as format.clear does (Excel's Clear Formats, Clear All): whole
// columns' and rows' own formats go too, and a cell showing its column's format gets the default one.

const FIXTURE: Fixture = {
	sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>' }],
	styles: STYLES
}

async function open(): Promise<{ session: SheetSessionApi; sheet: number }> {
	const result = await inThreadSheetHost().open(buildFixture(FIXTURE), { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

	if (result.kind !== "opened") throw new Error(result.kind)

	return { session: result.session, sheet: result.info.activeSheet }
}

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason}`)
}

async function sheetXml(session: SheetSessionApi): Promise<string> {
	return parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""
}

const COLUMN = { startRow: 0, startCol: 1, endRow: SHEET_MAX_ROWS - 1, endCol: 1 }
const ROW = { startRow: 3, startCol: 0, endRow: 3, endCol: SHEET_MAX_COLS - 1 }

describe("clearing cells' formats", () => {
	for (const what of ["formats", "all"] as const) {
		it(`clears whole columns' and rows' formats with "${what}" as format.clear does`, async () => {
			const both: { session: SheetSessionApi; sheet: number }[] = [await open(), await open()]

			for (const { session, sheet } of both) {
				ok(await session.command({ type: "format.apply", sheet, ranges: [COLUMN, ROW], patch: { bold: true } }))
			}

			const [cells, format] = both as [{ session: SheetSessionApi; sheet: number }, { session: SheetSessionApi; sheet: number }]
			const clear: SheetCommand = { type: "cells.clear", sheet: cells.sheet, ranges: [COLUMN, ROW], what }

			ok(await cells.session.command(clear))
			ok(await format.session.command({ type: "format.clear", sheet: format.sheet, ranges: [COLUMN, ROW] }))

			const cleared = await sheetXml(cells.session)

			expect(cleared).not.toContain("<cols>")
			expect(cleared).not.toContain("customFormat")

			if (what === "formats") {
				expect(cleared).toBe(await sheetXml(format.session))
				expect(cleared).toContain("<v>2</v>")
			} else {
				expect(cleared).not.toContain("<v>2</v>")
			}

			// One step: undo puts the formats back.
			ok(await cells.session.undo())
			expect(await sheetXml(cells.session)).toContain("<cols>")
			cells.session.close()
			format.session.close()
		})
	}
})
