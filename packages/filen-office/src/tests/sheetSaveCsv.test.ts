import { describe, expect, it } from "vitest"
import { SHEET_MAX_ROWS } from "../core/sheet/protocol"
import { borderPresetEdges } from "../core/sheet/style/borders"
import { openFixture } from "./sheetSessionFixtures"
import { STYLES } from "./xlsxWriteFixtures"

// Download as CSV covers the data area, as LibreOffice's CSV export does (measured, 26.8): from A1 to the last row
// and column holding a value or a formula, each row as wide as that area. Cells holding only a format are not data.

async function csvOf(
	body: string,
	edit?: (sheet: number) => Parameters<Awaited<ReturnType<typeof openFixture>>["session"]["command"]>[0]
): Promise<string> {
	const { session, info } = await openFixture({ sheets: [{ name: "One", body }], styles: STYLES })
	const sheet = info.activeSheet

	if (edit !== undefined) {
		const result = await session.command(edit(sheet))

		expect(result.ok).toBe(true)
	}

	const csv = await session.export("csv", { sheet })

	session.close()

	return csv.ok ? new TextDecoder("utf-8", { ignoreBOM: true }).decode(csv.bytes) : ""
}

describe("Download as CSV", () => {
	it("stops at the data area when a whole column's border styles cells down to the last row", async () => {
		const csv = await csvOf(
			'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>' +
				'<row r="5" s="1" customFormat="1"/></sheetData>',
			sheet => ({
				type: "format.borders",
				sheet,
				ranges: [{ startRow: 0, startCol: 1, endRow: SHEET_MAX_ROWS - 1, endCol: 1 }],
				edges: borderPresetEdges("outside")
			})
		)

		expect(csv).toBe("﻿1,2\r\n,\r\n3,\r\n")
	})

	it("starts at A1 and makes every row as wide as the data area", async () => {
		const csv = await csvOf(
			'<sheetData><row r="3"><c r="C3" t="inlineStr"><is><t>x</t></is></c><c r="E3" s="1"/></row>' +
				'<row r="4"><c r="B4" t="inlineStr"><is><t>y</t></is></c></row><row r="9"><c r="A9" s="1"/></row></sheetData>'
		)

		expect(csv).toBe("﻿,,\r\n,,\r\n,,x\r\n,y,\r\n")
	})

	it("counts a formula whose result is empty as data", async () => {
		const csv = await csvOf('<sheetData><row r="2"><c r="B2" t="str"><f>""</f><v></v></c></row></sheetData>')

		expect(csv).toBe("﻿,\r\n,\r\n")
	})
})
