import { strFromU8, unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { inThreadSheetHost } from "../dom/sheet/host"
import { odsBytes, table } from "./odsFixtures"
import { buildXlsx } from "./oracles/sheet/synthetic"
import { IMMEDIATE } from "./sheetSessionFixtures"

// The results a save stores for formulas typed in the editor, in both formats: Excel's (and LibreOffice's) values,
// not the formula engine's own reading of INT, MOD and DATEDIF.
const FORMULAS = ["=INT(-2.7)", "=MOD(-3,2)", '=DATEDIF(DATE(2020,1,1),DATE(2021,3,1),"m")']
const EXPECTED = ["-3", "1", "14"]

async function savedPart(kind: "xlsx" | "ods"): Promise<string> {
	const bytes =
		kind === "xlsx"
			? buildXlsx({ sheets: [{ name: "S", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }] })
			: odsBytes({
					body: table("S", '<table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row>')
				})
	const opened = await inThreadSheetHost().open(bytes, { kind, slicing: IMMEDIATE, settings: { author: "T" } }, () => undefined)

	if (opened.kind !== "opened") throw new Error(opened.kind)

	const session = opened.session

	try {
		await session.intent()

		const sheet = opened.info.sheets[0]?.id ?? 0

		await session.command({
			type: "cells.set",
			sheet,
			cells: { rows: Uint32Array.of(1, 2, 3), cols: Uint16Array.of(0, 0, 0), inputs: FORMULAS }
		})

		const saved = await session.save()

		if (!saved.ok) throw new Error(saved.message)

		return strFromU8(unzipSync(saved.bytes)[kind === "xlsx" ? "xl/worksheets/sheet1.xml" : "content.xml"] ?? new Uint8Array())
	} finally {
		session.close()
	}
}

describe("results saved for typed formulas", () => {
	it("stores Excel's INT, MOD and DATEDIF results in a workbook", async () => {
		const xml = await savedPart("xlsx")

		expect(["A2", "A3", "A4"].map(ref => new RegExp(`<c r="${ref}"[^>]*>.*?<v>([^<]*)</v>`).exec(xml)?.[1])).toEqual(EXPECTED)
	})

	it("stores the same results in a spreadsheet document", async () => {
		const xml = await savedPart("ods")
		const values = [...xml.matchAll(/table:formula="[^"]*"[^>]*office:value="([^"]*)"/g)].map(match => match[1])

		expect(values).toEqual(EXPECTED)
	})
})
