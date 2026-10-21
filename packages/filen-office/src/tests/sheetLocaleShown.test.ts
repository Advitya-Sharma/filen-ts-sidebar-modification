import { describe, expect, it } from "vitest"
import { firstTile, openFixture, tileEntry } from "./sheetSessionFixtures"

// What a cell shows to edit (the formula bar, the in-cell editor) in the session's locale: a commit of it unchanged
// reads back the same cell.

const DE = { decimal: ",", group: ".", argument: ";", dateOrder: "dmy" } as const

describe("inputs shown in the session's locale", () => {
	it("shows a number with the locale's decimal and a formula with its separators", async () => {
		const { session, info } = await openFixture(
			{
				sheets: [
					{
						name: "One",
						body: '<sheetData><row r="1"><c r="A1"><v>1.5</v></c><c r="B1"><f>ROUND(A1,2)+0.5</f><v>2</v></c></row></sheetData>'
					}
				]
			},
			{ settings: { locale: DE } }
		)
		const sheet = info.sheets[0]!.id
		const input = async (col: number): Promise<string> => (await session.query("cell", { sheet, row: 0, col })).input

		expect(await input(0)).toBe("1,5")
		expect(await input(1)).toBe("=ROUND(A1;2)+0,5")
		expect(tileEntry(await firstTile(session, sheet), 0, 1)?.input).toBe("=ROUND(A1;2)+0,5")

		await session.command({
			type: "cells.set",
			sheet,
			cells: { rows: Uint32Array.of(0, 0), cols: Uint16Array.of(0, 1), inputs: ["1,5", "=ROUND(A1;2)+0,5"] }
		})

		expect(await input(0)).toBe("1,5")
		expect(await input(1)).toBe("=ROUND(A1;2)+0,5")
	})
})
