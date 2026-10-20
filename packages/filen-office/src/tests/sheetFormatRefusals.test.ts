import { describe, expect, it } from "vitest"
import type { SheetFormatPatch } from "../core/sheet/protocol"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE, buildFixture } from "./sheetSessionFixtures"
import { STYLES } from "./xlsxWriteFixtures"

// A format the styles part cannot hold is refused by the command, as the Format Cells dialog refuses it
// (numFmtProblem), so no save is ever refused for it.

const FIXTURE = buildFixture({
	sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }],
	styles: STYLES
})

describe("formats the file cannot hold", () => {
	it.each<[string, SheetFormatPatch]>([
		["a code with a control character", { numFmt: "0.0\u0001" }],
		["a code with a lone surrogate", { numFmt: "0\ud800" }],
		["a code longer than Excel takes", { numFmt: "0".repeat(300) }],
		["a code of five sections", { numFmt: "0;0;0;@;0" }],
		["an empty code", { numFmt: "" }],
		["a code with an unclosed quote", { numFmt: '0" kg' }],
		["a font name with a control character", { font: "Cal\u0002ibri" }],
		["an empty font name", { font: " " }]
	])("refuses %s, and the workbook still saves", async (_, patch) => {
		const result = await inThreadSheetHost().open(FIXTURE.slice(), { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

		if (result.kind !== "opened") throw new Error(result.kind)

		const { session } = result
		const sheet = result.info.activeSheet
		const applied = await session.command({
			type: "format.apply",
			sheet,
			ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
			patch
		})

		expect(applied.ok ? null : applied.refusal.reason).toBe("unsupported")
		expect((await session.save()).ok).toBe(true)
		session.close()
	})

	it("takes a code Excel takes", async () => {
		const result = await inThreadSheetHost().open(FIXTURE.slice(), { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

		if (result.kind !== "opened") throw new Error(result.kind)

		const applied = await result.session.command({
			type: "format.apply",
			sheet: result.info.activeSheet,
			ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
			patch: { numFmt: '#,##0.00 "kg";[Red]-0;"-";@' }
		})

		expect(applied.ok).toBe(true)
		expect((await result.session.save()).ok).toBe(true)
		result.session.close()
	})
})
