import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { parseClipboardHtml } from "../core/sheet/clipboard/html"
import { excelCode, libreOfficeNumber } from "../core/sheet/clipboard/libreOffice"
import { SHEET_MAX_EDIT_CELLS, type SheetLocale } from "../core/sheet/protocol"
import { firstTile, tileEntry } from "./sheetSessionFixtures"
import { openKeys, sheetBody } from "./sheetKeysFixtures"

// LibreOffice 7.5's "HTML (StarCalc)" export of fixtures/sheet-clipboard/libreoffice-7.5-locales.fods under each
// locale setting: sdval and sdnum in the locale's syntax, read as the workbook holds them.

const FIXTURES = join(import.meta.dirname, "fixtures", "sheet-clipboard")
const LOCALES = ["de-DE", "de-CH", "fr-FR", "fi-FI", "it-IT", "es-ES", "nl-NL", "sv-SE", "ru-RU", "pt-BR", "en-US"]

// The day of the week then the locale's separator, as NNNN writes it, where its separator is ", ".
const WEEKDAY_NNNN = 'dddd", """d mmmm yyyy'
const WEEKDAY = "dddd, d mmmm yyyy"
const NNNN = new Set(["de-DE", "de-CH", "es-ES", "ru-RU", "pt-BR", "en-US"])

const VALUES = [
	1234567.891,
	0.125,
	45306,
	0.753544560185185,
	45306,
	123456,
	-1234.5,
	true,
	1234.5,
	1.26041666666667,
	0.753472222222222,
	45306.5208333333,
	1.23456789012346e17,
	0.5
]

function formats(locale: string): (string | undefined)[] {
	const space = ["fr-FR", "fi-FI", "sv-SE", "ru-RU"].includes(locale)

	return [
		"#,##0.00",
		"0.00%",
		"dd.mm.yyyy",
		"hh:mm:ss.00",
		NNNN.has(locale) ? WEEKDAY_NNNN : WEEKDAY,
		"0.00E+00",
		`#,##0.00;[${["de-DE", "de-CH"].includes(locale) ? "Red" : "RED"}]-#,##0.00`,
		undefined,
		space ? '#,##0.00" "[$€-407]' : "#,##0.00 [$€-407]",
		"[hh]:mm",
		"h:mm AM/PM",
		"yyyy-mm-dd hh:mm",
		undefined,
		undefined
	]
}

function html(locale: string): string {
	return readFileSync(join(FIXTURES, `libreoffice-7.5-${locale}.html`), "utf8")
}

describe("LibreOffice HTML from any locale", () => {
	for (const locale of LOCALES)
		it(`reads ${locale}'s values and number formats`, () => {
			const table = parseClipboardHtml(html(locale), SHEET_MAX_EDIT_CELLS)

			if (table === null || table === "tooLarge") throw new Error("no table")

			expect(
				table.cells.map(cell => (cell.value.kind === "number" || cell.value.kind === "boolean" ? cell.value.value : cell.value))
			).toEqual(VALUES)
			expect(table.cells.map(cell => cell.format?.numFmt)).toEqual(formats(locale))
		})

	it("reads the keyword letters of each language LibreOffice localises", () => {
		expect(excelCode("TT.MM.JJJJ HH:MM", 0x0407)).toBe("dd.mm.yyyy hh:mm")
		expect(excelCode("PP.KK.VVVV TT:MM", 0x040b)).toBe("dd.mm.yyyy hh:mm")
		expect(excelCode("GG/MM/AA OOOO", 0x0410)).toBe("dd/mm/yy dddd")
		expect(excelCode("JJ/MM/AAAA", 0x040c)).toBe("dd/mm/yyyy")
		expect(excelCode("DD-MM-JJJJ UU:MM", 0x0413)).toBe("dd-mm-yyyy hh:mm")
		expect(excelCode("TT:MM", 0x041d)).toBe("hh:mm")
		expect(excelCode("DD/MM/AAAA", 0x0c0a)).toBe("dd/mm/yyyy")
		expect(excelCode("Standard", 0x0407)).toBe("General")
		expect(excelCode("[FARBE3]Standard;[ROT]-Standard", 0x0407)).toBe("[Color3]General;[Red]-General")
		expect(excelCode("[>1000,5]#.##0,0;0,00", 0x0407)).toBe("[>1000.5]#,##0.0;0.00")
		expect(excelCode("#.##0,00.", 0x0407)).toBe("#,##0.00,")
		expect(excelCode("0.0", 0x040c)).toBe("0\\.0")
	})

	it("drops a code a workbook cannot take, and keeps the value", () => {
		expect(excelCode("QQ JJJJ", 0x0407)).toBeUndefined()
		expect(excelCode("[NatNum1]0", 0x0409)).toBeUndefined()
		expect(excelCode("[BRAUN]0", 0x0407)).toBeUndefined()
		expect(excelCode("0 EUR", 0x0409)).toBeUndefined()
		expect(libreOfficeNumber("1,5", "1031;0;LOGISCH")).toEqual({ value: 1.5, boolean: true, format: undefined })
		// Longer than Excel's 255 characters: no code a workbook holds.
		expect(libreOfficeNumber("1,5", `1031;0;${"0".repeat(256)}`)).toEqual({ value: 1.5, boolean: false, format: undefined })
		expect(libreOfficeNumber("1,5", `1031;0;${"0".repeat(255)}`).format).toBe("0".repeat(255))
		// A system locale whose decimal point the language does not have: the value as written, no format.
		expect(libreOfficeNumber("1.5", "1031;0;0,00")).toEqual({ value: 1.5, boolean: false, format: undefined })
		expect(libreOfficeNumber("1,5", undefined)).toEqual({ value: 1.5, boolean: false, format: undefined })
	})

	const GERMAN: SheetLocale = { decimal: ",", group: ".", argument: ";", dateOrder: "dmy" }

	for (const locale of [undefined, GERMAN])
		it(`pastes German HTML into a ${locale === undefined ? "en-US" : "German"} session as shown there`, async () => {
			const { session } = await openKeys({ sheets: [{ name: "S", body: sheetBody([["old"]]) }] })

			if (locale !== undefined) session.configure({ locale })

			const pasted = await session.command({
				type: "clipboard.pasteHtml",
				sheet: 0,
				at: { row: 0, col: 0 },
				html: html("de-DE"),
				text: ""
			})

			expect(pasted.ok).toBe(true)

			const tile = await firstTile(session, 0)
			const shown = (row: number) => tileEntry(tile, row, 0)?.text

			expect(shown(0)).toBe(locale === undefined ? "1,234,567.89" : "1.234.567,89")
			expect(shown(1)).toBe(locale === undefined ? "12.50%" : "12,50%")
			expect(shown(2)).toBe("15.01.2024")
			expect(shown(8)).toBe(locale === undefined ? "1,234.50 €" : "1.234,50 €")
			expect(shown(9)).toBe("30:15")
			session.close()
		})
})
