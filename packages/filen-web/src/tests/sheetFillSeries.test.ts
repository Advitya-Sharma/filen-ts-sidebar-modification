import { describe, expect, it } from "vitest"
import { EN_US_LOCALE, type SheetLocale } from "@filen/office/core/sheet"
import {
	seriesCommand,
	seriesDefaults,
	seriesStep,
	seriesStopValid,
	seriesTrendOffered
} from "@/features/spreadsheet/fill/seriesDialog.logic"

// The Series dialog's defaults and the command it sends.

const COLUMN = { startRow: 0, startCol: 2, endRow: 9, endCol: 2 }
const ROW = { startRow: 3, startCol: 0, endRow: 3, endCol: 7 }
const EN = EN_US_LOCALE
const GERMAN: SheetLocale = { decimal: ",", group: ".", argument: ";", dateOrder: "dmy" }

describe("the Series dialog", () => {
	it("starts along the selection's longer side, linear, stepping 1", () => {
		expect(seriesDefaults(COLUMN)).toEqual({ orientation: "cols", type: "linear", dateUnit: "day", step: "1", stop: "", trend: false })
		expect(seriesDefaults(ROW).orientation).toBe("rows")
	})

	it("reads the step as a number and the stop as a number or an ISO date", () => {
		expect(seriesStep(" -2.5 ", EN)).toBe(-2.5)
		expect(seriesStep("1e3", EN)).toBe(1000)
		expect(seriesStep("two", EN)).toBeNull()
		expect(seriesStep("", EN)).toBeNull()
		expect(seriesStopValid("", EN)).toBe(true)
		expect(seriesStopValid("100", EN)).toBe(true)
		expect(seriesStopValid("2024-12-31", EN)).toBe(true)
		expect(seriesStopValid("2024-12-31 18:30", EN)).toBe(true)
		expect(seriesStopValid("18:30", EN)).toBe(true)
		expect(seriesStopValid("soon", EN)).toBe(false)
	})

	it("reads both in the session's locale, as every other entry", () => {
		expect(seriesStep("0,5", GERMAN)).toBe(0.5)
		expect(seriesStep("1.000", GERMAN)).toBe(1000)
		expect(seriesStep("0,5", EN)).toBeNull()
		expect(seriesStopValid("2,5", GERMAN)).toBe(true)
		expect(seriesStopValid("15.01.2024", GERMAN)).toBe(true)
		expect(seriesStopValid("1/15/2024", EN)).toBe(true)
		expect(seriesStopValid("15.01.2024", EN)).toBe(false)

		const range = { startRow: 0, startCol: 0, endRow: 9, endCol: 0 }

		expect(seriesCommand({ ...seriesDefaults(range), step: "0,5", stop: "2,5" }, 0, range, GERMAN)).toMatchObject({
			options: { step: 0.5, stop: "2,5" }
		})
	})

	it("sends fill.series over the selection, the trend only for linear and growth", () => {
		const form = {
			...seriesDefaults(COLUMN),
			type: "date" as const,
			dateUnit: "month" as const,
			step: "2",
			stop: " 2025-06-30 ",
			trend: true
		}

		expect(seriesTrendOffered(form)).toBe(false)
		expect(seriesCommand(form, 1, COLUMN, EN)).toEqual({
			type: "fill.series",
			sheet: 1,
			range: COLUMN,
			orientation: "cols",
			options: { type: "date", dateUnit: "month", step: 2, stop: "2025-06-30", trend: false }
		})
		expect(seriesCommand({ ...form, type: "growth" }, 1, COLUMN, EN)).toMatchObject({ options: { type: "growth", trend: true } })
		expect(seriesCommand({ ...form, step: "x" }, 1, COLUMN, EN)).toBeNull()
		expect(seriesCommand({ ...form, stop: "never" }, 1, COLUMN, EN)).toBeNull()
	})
})
