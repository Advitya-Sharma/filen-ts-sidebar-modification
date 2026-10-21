import type { SheetLocale } from "../protocol"

// Entries as typed in the user's locale, and the file's own syntax shown back in it. Files always store en-US syntax
// and English function names (ECMA-376 Part 1 §18.17: the formula grammar's "," between arguments, "." in numbers);
// the app injects the locale (the core reads no Intl). A locale with "." and "," and "," between arguments is en-US's
// and changes nothing.
//
// In a formula typed with ";" between arguments, an array constant separates its columns with "\" and its rows with
// ";" (as Excel does where "," is the decimal point); with "," between arguments, as en-US: "," and ";".

const QUOTE = 34
const APOSTROPHE = 39
const BRACKET_OPEN = 91
const BRACKET_CLOSE = 93

export const EN_US_LOCALE: SheetLocale = { decimal: ".", group: ",", argument: ",", dateOrder: "mdy" }

// Whether the locale types formulas and numbers as en-US does.
export function plainLocale(locale: SheetLocale | undefined): boolean {
	return locale === undefined || (locale.decimal === "." && locale.argument === ",")
}

// The end of a quoted run (string or sheet name) starting at `start`, its doubled quotes inside.
function quotedEnd(text: string, start: number, quote: number): number {
	let index = start + 1

	while (index < text.length) {
		if (text.charCodeAt(index) === quote) {
			if (text.charCodeAt(index + 1) === quote) {
				index += 2
				continue
			}

			return index + 1
		}

		index++
	}

	return text.length
}

function bracketEnd(text: string, start: number): number {
	let depth = 0

	for (let index = start; index < text.length; index++) {
		const code = text.charCodeAt(index)

		if (code === BRACKET_OPEN) depth++
		else if (code === BRACKET_CLOSE && --depth === 0) return index + 1
	}

	return text.length
}

function isDigit(code: number): boolean {
	return code >= 48 && code <= 57
}

function isWordChar(code: number): boolean {
	return (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || code === 95 || code === 36 || code > 127
}

// Rewrites a formula's separators outside strings, quoted sheet names and structured references: `argument` for the
// argument separator, `decimal` for a decimal point between digits (in a number, not a name such as LOG10.2 or a
// sheet's Q1.2024), and in array constants `column` and `row`.
function mapSeparators(
	formula: string,
	from: { argument: string; decimal: string; column: string; row: string },
	to: { argument: string; decimal: string; column: string; row: string }
): string {
	let out = ""
	let braces = 0
	let index = 0

	while (index < formula.length) {
		const code = formula.charCodeAt(index)

		if (code === QUOTE || code === APOSTROPHE) {
			const end = quotedEnd(formula, index, code)

			out += formula.slice(index, end)
			index = end

			continue
		}

		if (code === BRACKET_OPEN) {
			const end = bracketEnd(formula, index)

			out += formula.slice(index, end)
			index = end

			continue
		}

		// A run of word characters (a name, a cell, a function): copied whole, so a point in it is never a decimal.
		if (isWordChar(code)) {
			let end = index + 1

			while (
				end < formula.length &&
				(isWordChar(formula.charCodeAt(end)) || isDigit(formula.charCodeAt(end)) || formula[end] === ".")
			)
				end++

			out += formula.slice(index, end)
			index = end

			continue
		}

		const char = formula[index] ?? ""

		if (code === 123) braces++
		else if (code === 125) braces = Math.max(0, braces - 1)

		if (
			char === from.decimal &&
			isDigit(formula.charCodeAt(index + 1)) &&
			(index === 0 || !isWordChar(formula.charCodeAt(index - 1)))
		) {
			// A decimal point: a number's digits on its right, its own digits (or nothing) on its left.
			out += to.decimal
		} else if (braces > 0 && char === from.column) out += to.column
		else if (braces > 0 && char === from.row) out += to.row
		else if (braces === 0 && char === from.argument) out += to.argument
		else out += char

		index++
	}

	return out
}

function separators(locale: SheetLocale): { argument: string; decimal: string; column: string; row: string } {
	return locale.argument === ";"
		? { argument: ";", decimal: locale.decimal, column: "\\", row: ";" }
		: { argument: locale.argument, decimal: locale.decimal, column: ",", row: ";" }
}

const EN_US_SEPARATORS = { argument: ",", decimal: ".", column: ",", row: ";" }

// A formula as typed in the locale (without its "="), in the file's syntax.
export function delocalizeFormula(formula: string, locale: SheetLocale | undefined): string {
	if (locale === undefined || plainLocale(locale)) return formula

	return mapSeparators(formula, separators(locale), EN_US_SEPARATORS)
}

// A formula in the file's syntax (without its "="), as the locale types it.
export function localizeFormula(formula: string, locale: SheetLocale | undefined): string {
	if (locale === undefined || plainLocale(locale)) return formula

	return mapSeparators(formula, EN_US_SEPARATORS, separators(locale))
}

function escaped(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

interface NumberReader {
	readonly plain: RegExp
	readonly grouped: RegExp
	readonly percent: RegExp
	readonly decimal: string
	readonly group: string
}

const readers = new Map<string, NumberReader>()

function numberReader(locale: SheetLocale): NumberReader {
	const key = `${locale.decimal}\u0000${locale.group}`
	let reader = readers.get(key)

	if (reader === undefined) {
		const point = escaped(locale.decimal)
		// A space group separator (fr-FR's narrow no-break space too) also accepts a plain space as typed.
		const group = /\s/u.test(locale.group) ? "[\\s\\u00a0\\u202f]" : escaped(locale.group)

		reader = {
			plain: new RegExp(`^[-+]?(\\d+${point}?\\d*|${point}\\d+)(e[-+]?\\d+)?$`, "i"),
			grouped: new RegExp(`^[-+]?\\d{1,3}(${group}\\d{3})+(${point}\\d+)?$`),
			percent: new RegExp(`^[-+]?(\\d+${point}?\\d*|${point}\\d+)%$`),
			decimal: locale.decimal,
			group: locale.group
		}
		readers.set(key, reader)
	}

	return reader
}

// A number as typed in the locale: its value, and the format a grouped or percent entry implies. Null when it reads
// as no number.
export function localizedNumber(text: string, locale: SheetLocale): { value: number; impliedFormat?: string } | null {
	const reader = numberReader(locale)
	const canonical = (digits: string): number =>
		Number(
			digits
				.replace(/[\s\u00a0\u202f]/gu, "")
				.split(reader.group)
				.join("")
				.replace(reader.decimal, ".")
		)

	if (reader.plain.test(text)) return { value: Number(text.replace(reader.decimal, ".")) }

	if (reader.grouped.test(text)) {
		return { value: canonical(text), impliedFormat: text.includes(reader.decimal) ? "#,##0.00" : "#,##0" }
	}

	if (reader.percent.test(text)) {
		return {
			value: Number(text.slice(0, -1).replace(reader.decimal, ".")) / 100,
			impliedFormat: text.includes(reader.decimal) ? "0.00%" : "0%"
		}
	}

	return null
}

const DATE = /^(\d{1,4})([/.-])(\d{1,2})\2(\d{1,4})$/

// A date as typed in the locale's order (day, month and year separated by "/", "-" or "."; a two-digit year is
// 1930-2029, as Excel reads one: support.microsoft.com "How Excel works with two-digit year numbers"). ISO dates
// (yyyy-mm-dd) read in every locale. Null for anything else.
export function localizedDate(text: string, locale: SheetLocale): { year: number; month: number; day: number } | null {
	const match = DATE.exec(text)

	if (match === null) return null

	const [, first = "", separator = "", second = "", third = ""] = match
	const order = first.length === 4 ? "ymd" : locale.dateOrder

	// A "." separator with "." as the decimal point is a number's, not a date's.
	if (separator === "." && locale.decimal === ".") return null

	const parts = order === "ymd" ? [first, second, third] : order === "dmy" ? [third, second, first] : [third, first, second]
	const [yearText = "", monthText = "", dayText = ""] = parts

	if (order !== "ymd" && first.length > 2) return null
	if (yearText.length !== 2 && yearText.length !== 4) return null

	const short = Number(yearText)
	const year = yearText.length === 2 ? (short < 30 ? 2000 + short : 1900 + short) : short
	const month = Number(monthText)
	const day = Number(dayText)
	const date = new Date(Date.UTC(year, month - 1, day))

	if (year < 1900 || date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null

	return { year, month, day }
}

// The format a date typed in the locale implies: its order with four-digit years (Excel's short date for the
// locale, which the file stores as a custom code).
export function localizedDateFormat(locale: SheetLocale): string {
	switch (locale.dateOrder) {
		case "dmy":
			return "dd/mm/yyyy"
		case "ymd":
			return "yyyy-mm-dd"
		default:
			return "m/d/yyyy"
	}
}

// A cell's input (as the session shows it to edit, in the file's syntax) as the locale types it: a formula's
// separators, a plain number's decimal point. Text and anything else as it is.
export function localizedInput(input: string, locale: SheetLocale | undefined): string {
	if (locale === undefined || plainLocale(locale)) return input
	if (input.startsWith("=")) return `=${localizeFormula(input.slice(1), locale)}`
	if (/^-?\d+(\.\d+)?(e[-+]?\d+)?$/i.test(input)) return input.replace(".", locale.decimal)

	return input
}
