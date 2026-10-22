import { describe, expect, it } from "vitest"
import { DEFAULT_DIGIT_PX, colWidthToPx, defaultColPx, defaultColWidth, pxToColWidth, workbookDigitPx } from "../core/sheet/sizes/sizes"
import { readFileSync } from "node:fs"
import { readSfnt } from "../core/fonts/sfnt"
import { MAX_DIGIT_PX, validDigit } from "../core/xlsx/edit/digitWidth"
import { sfntFromWoff2 } from "../scripts/woff2Tables"
import { openFixture, savedBytes } from "./sheetSessionFixtures"
import { parts } from "./xlsxWriteFixtures"

// Column widths are in units of the widest digit of the workbook's default font (the Normal style's), ECMA-376 Part 1
// §18.3.1.13: Calibri 11 at 96 dpi is 7 px, another default font another width, worked out from the font alone.

describe("column widths in another default font's digits", () => {
	it("keeps Calibri 11's 7 px as the default and round-trips pixels at any digit width", () => {
		expect(DEFAULT_DIGIT_PX).toBe(7)
		expect(colWidthToPx(9.140625)).toBe(64)

		for (const digit of [5, 6, 7, 8, 9, 11])
			for (let px = 8; px <= 400; px++) expect(colWidthToPx(pxToColWidth(px, digit), digit)).toBe(px)
	})

	it("draws the same stored width narrower in a narrower font, as §18.3.1.13's formula gives", () => {
		// Truncate(((256 × 8.43 + Truncate(128 / 6)) / 256) × 6): Calibri 9's 6 px digits.
		expect(colWidthToPx(8.43, 6)).toBe(Math.trunc(((256 * 8.43 + Math.trunc(128 / 6)) / 256) * 6))
		expect(colWidthToPx(8.43, 6)).toBeLessThan(colWidthToPx(8.43))
	})

	it("sizes the default column from baseColWidth in the workbook's digits", () => {
		const base = { defaultColWidth: Number.NaN, baseColWidth: 8 }

		expect(defaultColPx(base)).toBe(64)
		// 8 × 6 + 5 = 53, rounded up to whole 8 px (MS-OI29500 2.1.598 f).
		expect(defaultColPx(base, 6)).toBe(56)
		expect(colWidthToPx(defaultColWidth(base, 6), 6)).toBe(56)
	})
})

// LibreOffice 7.5's column widths for a workbook whose Normal font is each of these (UNO Width of two columns 10 and
// 30 characters wide, in pixels per character): it lays them out by the fonts' unhinted digit advances. Measured with
// the fonts this machine has; Cambria (no face of equal widths here) is left out.
// Calibri 8 is left to Excel's own widths (below), which are a pixel wider.
const LIBREOFFICE_DIGITS: readonly (readonly [string, number, number])[] = [
	["Calibri", 9, 6.066],
	["Calibri", 10, 6.733],
	["Calibri", 11, 7.4],
	["Calibri", 12, 8.134],
	["Calibri", 14, 9.466],
	["Calibri", 16, 10.8],
	["Calibri", 20, 13.533],
	["Calibri", 26, 17.599],
	["Arial", 8, 5.934],
	["Arial", 10, 7.4],
	["Arial", 11, 8.134],
	["Arial", 12, 8.867],
	["Arial", 14, 10.399],
	["Times New Roman", 11, 7.334],
	["Times New Roman", 12, 7.999],
	["Courier New", 10, 7.999],
	["Courier New", 11, 8.801],
	["Georgia", 11, 9.001]
]

describe("the workbook's maximum digit width", () => {
	it("is Excel's 7 px for Calibri 11 and for Arial 10", () => {
		expect(workbookDigitPx("Calibri", 11)).toBe(7)
		expect(workbookDigitPx(undefined, undefined)).toBe(7)
		expect(workbookDigitPx("Arial", 10)).toBe(7)
	})

	it.each(LIBREOFFICE_DIGITS)("rounds what LibreOffice lays %s %d out by (%d px)", (font, size, libreOffice) => {
		expect(workbookDigitPx(font, size)).toBe(Math.round(libreOffice))
	})

	it("reads its digit advances from the bundled faces of equal widths", () => {
		const faces: readonly (readonly [string, string])[] = [
			["Calibri", "carlito/Carlito-Regular.woff2"],
			["Arial", "arimo/Arimo-VF.woff2"],
			["Times New Roman", "tinos/Tinos-Regular.woff2"],
			["Courier New", "cousine/Cousine-Regular.woff2"],
			["Georgia", "gelasio/Gelasio-VF.woff2"]
		]

		for (const [name, file] of faces) {
			const face = readSfnt(sfntFromWoff2(readFileSync(new URL(`../../assets/fonts/${file}`, import.meta.url))))
			const widest = Math.max(...Array.from({ length: 10 }, (_, digit) => face.advanceOf(0x30 + digit)))

			// At 1,000 points the rounding shows any difference in the advance.
			expect(workbookDigitPx(name, 1000 * 0.409), name).toBe(Math.round((widest / face.unitsPerEm) * 409 * (4 / 3)))
		}
	})

	it("takes the unit Excel's own files show for the fonts and sizes they name", () => {
		// The fractions of the widths Excel wrote: 75 px at 8 px digits is 9.375; at 7 px it would be 10.7109375.
		expect(workbookDigitPx("Verdana", 10)).toBe(8)
		expect(workbookDigitPx("verdana", 10.0)).toBe(8)
		expect(workbookDigitPx("Calibri", 8)).toBe(6)
		expect(workbookDigitPx("Arial", 12)).toBe(9)
		expect(workbookDigitPx("Aptos Narrow", 11)).toBe(7)
	})

	it("counts a font with neither Excel's evidence nor a face of equal widths in Calibri's digits at its own size", () => {
		expect(workbookDigitPx("Cambria", 11)).toBe(7)
		// One source only says 7 for Calibri 9: its advances give 6.
		expect(workbookDigitPx("Calibri", 9)).toBe(6)
		expect(workbookDigitPx("Verdana", 12)).toBe(8)
		expect(workbookDigitPx("Aptos Narrow", 20)).toBe(14)
	})

	it("holds a digit wider than 64 px, up to Excel's 409-point fonts", () => {
		expect(workbookDigitPx("Calibri", 100)).toBe(68)
		expect(workbookDigitPx("Georgia", 409)).toBe(MAX_DIGIT_PX)
		expect(workbookDigitPx("Calibri", 5000)).toBe(workbookDigitPx("Calibri", 409))
		expect(validDigit(68)).toBe(68)
		expect(validDigit(10_000)).toBe(MAX_DIGIT_PX)
	})
})

describe("a session's column widths in the digit width it is handed", () => {
	const FIXTURE = {
		sheets: [
			{
				name: "One",
				body: '<cols><col min="2" max="2" width="10" customWidth="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
			}
		]
	}

	it("lays out the file's widths and the default column in it, and writes a resize in it", async () => {
		const { session, info } = await openFixture(FIXTURE)
		const one = info.sheets[0]?.id ?? 0

		expect(info.activeLayout.cols.defaultSize).toBe(64)

		expect(await session.query("digitWidth", { px: 6 })).toMatchObject({ px: 6 })

		const layout = await session.layout(one)

		expect(layout.cols.defaultSize).toBe(56)
		expect(Array.from(layout.cols.sizes)).toEqual([colWidthToPx(10, 6)])

		await session.command({ type: "axis.resize", sheet: one, axis: "cols", indices: Uint32Array.of(2), sizes: Float32Array.of(60) })

		const sheet = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheet).toContain(`<col min="3" max="3" width="${String(pxToColWidth(60, 6))}" customWidth="1"/>`)
		session.close()
	})

	it("holds a width handed in to whole pixels the unit can be", async () => {
		const { session } = await openFixture(FIXTURE)

		expect(await session.query("digitWidth", { px: 0 })).toMatchObject({ px: 1 })
		expect(await session.query("digitWidth", { px: 6.4 })).toMatchObject({ px: 6 })
		expect(await session.query("digitWidth", { px: 80 })).toMatchObject({ px: 80 })
		expect(await session.query("digitWidth", { px: Number.NaN })).toMatchObject({ px: 7 })
		session.close()
	})

	it("works the unit out from the default font when asked, and lays out in it", async () => {
		const { session, info } = await openFixture({
			...FIXTURE,
			styles: '<fonts count="1"><font><sz val="12"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
		})

		expect(await session.query("digitWidth", {})).toEqual({ px: 8, font: { font: "Calibri", size: 12 } })
		expect((await session.layout(info.sheets[0]?.id ?? 0)).cols.sizes[0]).toBe(colWidthToPx(10, 8))
		session.close()
	})

	it("names the default font, the Normal style's", async () => {
		const { session } = await openFixture({
			...FIXTURE,
			styles: '<fonts count="1"><font><sz val="10"/><name val="Verdana"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
		})

		expect(await session.query("digitWidth", {})).toEqual({ px: 8, font: { font: "Verdana", size: 10 } })
		session.close()
	})
})
