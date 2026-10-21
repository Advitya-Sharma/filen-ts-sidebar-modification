import { cellKey, type CellStyleView, type CellView, type SheetView } from "../model/model"
import type { SheetLocale } from "../protocol"
import { wellFormedFormula } from "../refs/formulaRefs"
import { delocalizeFormula, localizedDate, localizedDateFormat, localizedNumber } from "./locale"

// What a typed entry becomes, the way spreadsheets read it: a leading "=" starts a formula, a leading "'"
// keeps the rest as text whatever it looks like, and otherwise numbers (with thousands separators, a
// percent sign, an exponent), TRUE/FALSE and ISO dates are recognised, with the format they imply for a
// cell that has none. Anything else is text, as typed.
export type ParsedInput =
	| { type: "empty" }
	| { type: "formula"; formula: string }
	| { type: "value"; value: string | number | boolean | Date; impliedFormat?: string }

const NUMBER = /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i
const GROUPED = /^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/
const PERCENT = /^[-+]?(\d+\.?\d*|\.\d+)%$/
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/

//
// With the app's locale, numbers read with its decimal point and grouping, dates in its order (ISO dates in any),
// and a formula with its separators, given back in the file's syntax (input/locale.ts).
export function parseCellInput(input: string, locale?: SheetLocale): ParsedInput {
	if (input === "") {
		return { type: "empty" }
	}

	if (input.startsWith("=") && input.length > 1) {
		return { type: "formula", formula: delocalizeFormula(input.slice(1), locale) }
	}

	if (input.startsWith("'")) {
		return { type: "value", value: input.slice(1) }
	}

	const trimmed = input.trim()

	if (locale !== undefined) return localizedValue(input, trimmed, locale)

	if (NUMBER.test(trimmed)) {
		return { type: "value", value: Number(trimmed) }
	}

	if (GROUPED.test(trimmed)) {
		return { type: "value", value: Number(trimmed.replaceAll(",", "")), impliedFormat: trimmed.includes(".") ? "#,##0.00" : "#,##0" }
	}

	if (PERCENT.test(trimmed)) {
		return { type: "value", value: Number(trimmed.slice(0, -1)) / 100, impliedFormat: trimmed.includes(".") ? "0.00%" : "0%" }
	}

	const upper = trimmed.toUpperCase()

	if (upper === "TRUE" || upper === "FALSE") {
		return { type: "value", value: upper === "TRUE" }
	}

	const date = ISO_DATE.exec(trimmed)

	if (date !== null) {
		const [, year = "", month = "", day = ""] = date
		const value = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)))

		if (value.getUTCMonth() === Number(month) - 1 && value.getUTCDate() === Number(day)) {
			return { type: "value", value, impliedFormat: "yyyy-mm-dd" }
		}
	}

	return { type: "value", value: input }
}

// A value typed in the app's locale (en-US's included: its dates are month first).
function localizedValue(input: string, trimmed: string, locale: SheetLocale): ParsedInput {
	const number = localizedNumber(trimmed, locale)

	if (number !== null) return { type: "value", ...number }

	const upper = trimmed.toUpperCase()

	if (upper === "TRUE" || upper === "FALSE") return { type: "value", value: upper === "TRUE" }

	const date = localizedDate(trimmed, locale)

	if (date !== null) {
		const iso = ISO_DATE.test(trimmed)

		return {
			type: "value",
			value: new Date(Date.UTC(date.year, date.month - 1, date.day)),
			impliedFormat: iso ? "yyyy-mm-dd" : localizedDateFormat(locale)
		}
	}

	return { type: "value", value: input }
}

// Whether a workbook cell takes what is typed as text, as the worker reads it: by the cell's own format
// when it has a cell, else by its row's format, else by its column's.
export function takesText(
	sheet: Pick<SheetView, "textFormat">,
	row: number,
	col: number,
	cell: CellView | undefined,
	styles: readonly CellStyleView[]
): boolean {
	if (cell !== undefined) {
		return cell.style !== undefined && styles[cell.style]?.numFmt === "@"
	}

	const format = sheet.textFormat

	return (
		format !== undefined &&
		!format.otherCells.includes(cellKey(row, col)) &&
		(format.rows.includes(row) || (!format.otherRows.includes(row) && format.cols.includes(col)))
	)
}

// Whether an entry would be stored as a formula that does not read as one (a parenthesis left open, ";"
// between arguments): refused rather than written to the file, where Excel would drop it. A Text ("@")
// cell takes it as text.
export function malformedFormula(input: string, textFormat = false, locale?: SheetLocale): boolean {
	if (textFormat) {
		return false
	}

	const parsed = parseCellInput(input, locale)

	return parsed.type === "formula" && !wellFormedFormula(parsed.formula)
}
