import { describe, expect, it } from "vitest"
import type { SheetCommandResult, SheetSessionApi } from "../core/sheet/protocol"
import { inThreadSheetHost } from "../dom/sheet/host"
import { IMMEDIATE, buildFixture, savedBytes } from "./sheetSessionFixtures"
import { STYLES, parts } from "./xlsxWriteFixtures"

// A row without a height of its own fits its tallest font when a format changes a font size, as Excel's row
// autofit does; the file keeps the fitted height as ht without customHeight (ECMA-376 Part 1 §18.3.1.73).

const BODY =
	'<sheetFormatPr defaultRowHeight="15"/><sheetData>' +
	'<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' +
	'<row r="2" ht="30" customHeight="1"><c r="A2"><v>3</v></c></row>' +
	'<row r="3" ht="45"><c r="A3" s="2"><v>4</v></c></row>' +
	'<row r="4" ht="30"><c r="A4"><v>5</v></c></row></sheetData>'
const WRAP_STYLES = STYLES.replace('<cellXfs count="2">', '<cellXfs count="3">').replace(
	"</cellXfs>",
	'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>'
)

async function open(): Promise<{ session: SheetSessionApi; sheet: number }> {
	const result = await inThreadSheetHost().open(
		buildFixture({ sheets: [{ name: "One", body: BODY }], styles: WRAP_STYLES }),
		{ kind: "xlsx", slicing: IMMEDIATE },
		() => undefined
	)

	if (result.kind !== "opened") throw new Error(result.kind)

	return { session: result.session, sheet: result.info.activeSheet }
}

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason}`)
}

async function height(session: SheetSessionApi, sheet: number, row: number): Promise<number> {
	const layout = await session.layout(sheet)
	const at = layout.rows.indices.indexOf(row)

	return at < 0 ? layout.rows.defaultSize : (layout.rows.sizes[at] ?? -1)
}

const cell = (row: number, col: number) => [{ startRow: row, startCol: col, endRow: row, endCol: col }]

describe("row heights after a font size changes", () => {
	it("grows a row to a larger font and back, writing ht without customHeight", async () => {
		const { session, sheet } = await open()

		expect(await height(session, sheet, 0)).toBe(20)
		ok(await session.command({ type: "format.apply", sheet, ranges: cell(0, 1), patch: { size: 20 } }))
		// A 20 pt line (26.67 px at 1.2) where an 11 pt line is: 20 + 32 - 18.
		expect(await height(session, sheet, 0)).toBe(34)
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toContain('<row r="1" ht="25.5">')

		ok(await session.command({ type: "format.apply", sheet, ranges: cell(0, 1), patch: { size: 11 } }))
		expect(await height(session, sheet, 0)).toBe(20)
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toContain('<row r="1">')
		session.close()
	})

	it("keeps a height set by hand, and gives the fitted one back on undo", async () => {
		const { session, sheet } = await open()

		ok(await session.command({ type: "format.apply", sheet, ranges: cell(1, 0), patch: { size: 30 } }))
		expect(await height(session, sheet, 1)).toBe(40)

		ok(await session.command({ type: "format.apply", sheet, ranges: cell(0, 0), patch: { size: 30 } }))
		expect(await height(session, sheet, 0)).toBeGreaterThan(40)
		ok(await session.undo())
		expect(await height(session, sheet, 0)).toBe(20)
		session.close()
	})

	it("never shrinks a row with wrapped text, whose lines the core cannot count", async () => {
		const { session, sheet } = await open()

		ok(await session.command({ type: "format.apply", sheet, ranges: cell(2, 1), patch: { size: 12 } }))
		expect(await height(session, sheet, 2)).toBe(60)
		session.close()
	})

	it("refits only for a font size: other formats leave a row's height from the file as it is", async () => {
		const { session, sheet } = await open()

		ok(await session.command({ type: "format.apply", sheet, ranges: cell(3, 0), patch: { bold: true } }))
		expect(await height(session, sheet, 3)).toBe(40)
		ok(await session.command({ type: "format.clear", sheet, ranges: cell(3, 0) }))
		expect(await height(session, sheet, 3)).toBe(20)
		session.close()
	})
})

describe("rows fitted to the wrapped text the grid measured (axis.fit)", () => {
	const fit = (sheet: number, rows: number[], heights: number[], firstCol = 0, lastCol = 31) =>
		({ type: "axis.fit", sheet, rows: Uint32Array.from(rows), heights: Float32Array.from(heights), firstCol, lastCol }) as const

	it("grows a row to its wrapped text and shrinks it back to its font, as one step each, ht without customHeight", async () => {
		const { session, sheet } = await open()

		ok(await session.command(fit(sheet, [0], [60])))
		expect(await height(session, sheet, 0)).toBe(60)
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toContain('<row r="1" ht="45">')

		ok(await session.command(fit(sheet, [0], [Number.NaN])))
		expect(await height(session, sheet, 0)).toBe(20)
		ok(await session.undo())
		expect(await height(session, sheet, 0)).toBe(60)
		session.close()
	})

	it("keeps a height set by hand, and never shrinks below wrapped text the grid did not measure", async () => {
		const { session, sheet } = await open()

		ok(await session.command(fit(sheet, [1], [90])))
		expect(await height(session, sheet, 1)).toBe(40)
		// Row 3's wrapped A3 lies outside the measured columns: its 60 px stay.
		const result = await session.command(fit(sheet, [2], [Number.NaN], 5, 9))

		expect(result.ok).toBe(true)
		expect(await height(session, sheet, 2)).toBe(60)
		ok(await session.command(fit(sheet, [2], [Number.NaN], 0, 9)))
		expect(await height(session, sheet, 2)).toBe(20)
		session.close()
	})
})
