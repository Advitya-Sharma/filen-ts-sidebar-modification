import { describe, expect, it } from "vitest"
import { parseCellInput } from "../core/sheet/input/cellInput"
import { delocalizeFormula, localizedDate, localizedInput, localizedNumber, localizeFormula, textNumber } from "../core/sheet/input/locale"
import type { SheetLocale } from "../core/sheet/protocol"

const DE: SheetLocale = { decimal: ",", group: ".", argument: ";", dateOrder: "dmy" }
const FR: SheetLocale = { decimal: ",", group: " ", argument: ";", dateOrder: "dmy" }
const US: SheetLocale = { decimal: ".", group: ",", argument: ",", dateOrder: "mdy" }
const GB: SheetLocale = { decimal: ".", group: ",", argument: ",", dateOrder: "dmy" }

describe("formulas typed in a locale", () => {
	it("store en-US separators, leaving strings, sheet names and structured references as typed", () => {
		expect(delocalizeFormula('SUMME(A1;1,5;"a;b,c")', DE)).toBe('SUMME(A1,1.5,"a;b,c")')
		expect(delocalizeFormula("IF('Q1;2'!A1>0,5;{1,5\\2;3\\4};Table1[[#This Row];[a,b]])", DE)).toBe(
			"IF('Q1;2'!A1>0.5,{1.5,2;3,4},Table1[[#This Row];[a,b]])"
		)
		expect(delocalizeFormula("LOG10(2)+STDEV.S(A1:A3;,5)", DE)).toBe("LOG10(2)+STDEV.S(A1:A3,.5)")
	})

	it("show back as typed", () => {
		for (const formula of ['SUM(A1,1.5,"a;b,c")', "IF(A1>0.5,{1.5,2;3,4},0)", "ROUND(LOG10(B2),2)"])
			expect(delocalizeFormula(localizeFormula(formula, DE), DE)).toBe(formula)

		expect(localizeFormula("IF(A1>0.5,{1.5,2;3,4},0)", DE)).toBe("IF(A1>0,5;{1,5\\2;3\\4};0)")
	})

	it("change nothing in en-US", () => {
		expect(delocalizeFormula("SUM(A1,1.5)", US)).toBe("SUM(A1,1.5)")
		expect(localizeFormula("SUM(A1,1.5)", GB)).toBe("SUM(A1,1.5)")
	})
})

describe("values typed in a locale", () => {
	it("read numbers with the locale's point and grouping", () => {
		expect(localizedNumber("1,5", DE)).toEqual({ value: 1.5 })
		expect(localizedNumber("1.234,5", DE)).toEqual({ value: 1234.5, impliedFormat: "#,##0.00" })
		expect(localizedNumber("12,5%", DE)).toEqual({ value: 0.125, impliedFormat: "0.00%" })
		expect(localizedNumber("1 234", FR)).toEqual({ value: 1234, impliedFormat: "#,##0" })
		expect(localizedNumber("1.5", DE)).toBeNull()
		expect(localizedNumber("1,234.5", US)).toEqual({ value: 1234.5, impliedFormat: "#,##0.00" })
	})

	it("read dates in the locale's order, two-digit years as Excel does", () => {
		expect(localizedDate("3/5/2024", US)).toEqual({ year: 2024, month: 3, day: 5 })
		expect(localizedDate("3/5/2024", GB)).toEqual({ year: 2024, month: 5, day: 3 })
		expect(localizedDate("31.12.29", DE)).toEqual({ year: 2029, month: 12, day: 31 })
		expect(localizedDate("1.1.30", DE)).toEqual({ year: 1930, month: 1, day: 1 })
		expect(localizedDate("2024-02-30", DE)).toBeNull()
		expect(localizedDate("2024-02-29", US)).toEqual({ year: 2024, month: 2, day: 29 })
		expect(localizedDate("1.5.2024", US)).toBeNull()
	})

	it("parse an entry", () => {
		expect(parseCellInput("=SUMME(A1;2,5)", DE)).toEqual({ type: "formula", formula: "SUMME(A1,2.5)" })
		expect(parseCellInput("2,5", DE)).toEqual({ type: "value", value: 2.5 })
		expect(parseCellInput("31.12.2024", DE)).toEqual({
			type: "value",
			value: new Date(Date.UTC(2024, 11, 31)),
			impliedFormat: "dd/mm/yyyy"
		})
		expect(parseCellInput("12/31/2024", US)).toEqual({
			type: "value",
			value: new Date(Date.UTC(2024, 11, 31)),
			impliedFormat: "m/d/yyyy"
		})
		expect(parseCellInput("'2,5", DE)).toEqual({ type: "value", value: "2,5" })
		expect(parseCellInput("wahr", DE)).toEqual({ type: "value", value: "wahr" })
		// Without a locale, as before: ISO dates only.
		expect(parseCellInput("12/31/2024")).toEqual({ type: "value", value: "12/31/2024" })
	})

	it("show a cell's input as the locale types it", () => {
		expect(localizedInput("1.5", DE)).toBe("1,5")
		expect(localizedInput("=ROUND(A1,2)+0.5", DE)).toBe("=ROUND(A1;2)+0,5")
		expect(localizedInput("a.b", DE)).toBe("a.b")
		expect(localizedInput("2024-01-02", DE)).toBe("2024-01-02")
		expect(localizedInput("1.5", US)).toBe("1.5")
	})
})

describe("numbers typed with an Arabic decimal separator locale", () => {
	// Intl's ar-EG, ar-SA, fa-IR and ps-AF: U+066B as the decimal point, U+066C between groups.
	const AR: SheetLocale = { decimal: "٫", group: "٬", argument: ",", dateOrder: "dmy" }

	it("take '.' as the decimal point too, as LibreOffice 7.5 reads entries in ar-EG and fa-IR", () => {
		expect(parseCellInput("3.5", AR)).toEqual({ type: "value", value: 3.5 })
		expect(parseCellInput("3٫5", AR)).toEqual({ type: "value", value: 3.5 })
		expect(localizedNumber("1٬234.5", AR)?.value).toBe(1234.5)
		expect(localizedNumber("12.5%", AR)?.value).toBe(0.125)
	})

	it("read text in formulas with either decimal point", () => {
		expect(textNumber("3.5", AR)).toBe(3.5)
		expect(textNumber(".5", AR)).toBe(0.5)
		expect(textNumber("3٫5", AR)).toBe(3.5)
	})

	it("keep '.' as no decimal point where the locale's is a comma", () => {
		expect(localizedNumber("3.5", DE)).toBeNull()
	})
})
