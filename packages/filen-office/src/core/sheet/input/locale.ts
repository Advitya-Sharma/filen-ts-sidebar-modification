import type { SheetLocale } from "../protocol"
import { TERRITORY_CURRENCIES } from "./currencies.generated"

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

interface TextNumberReader {
	readonly number: RegExp
	readonly grouped: RegExp
	readonly decimal: string
	readonly decimalCode: number
	readonly group: RegExp
	// Longest first, so "US$" is read before "$".
	readonly currencies: readonly string[]
	readonly currencyCodes: readonly number[]
}

const textReaders = new Map<string, TextNumberReader>()

const SPACE = 32
const PLUS = 43
const MINUS = 45
const PAREN = 40

function isSpace(code: number): boolean {
	return code === SPACE || code === 0xa0 || code === 0x202f || (code >= 9 && code <= 13)
}

function textNumberReader(locale: SheetLocale): TextNumberReader {
	const currencies = (locale.currencies ?? (locale.decimal === "." ? ["$"] : []))
		.filter(symbol => symbol !== "")
		.sort((a, z) => z.length - a.length)
	const key = `${locale.decimal}\u0000${locale.group}\u0000${currencies.join("\u0000")}`
	let reader = textReaders.get(key)

	if (reader === undefined) {
		const point = escaped(locale.decimal)
		const group = /\s/u.test(locale.group) ? "[\\s\\u00a0\\u202f]" : escaped(locale.group)

		reader = {
			number: new RegExp(`^(\\d+(${point}\\d*)?|${point}\\d+)(e[-+]?\\d+)?$`, "i"),
			grouped: new RegExp(`^\\d{1,3}(${group}\\d{3})+(${point}\\d*)?$`),
			decimal: locale.decimal,
			decimalCode: locale.decimal.charCodeAt(0),
			group: new RegExp(group, "g"),
			currencies,
			currencyCodes: currencies.map(symbol => symbol.charCodeAt(0))
		}
		textReaders.set(key, reader)
	}

	return reader
}

// The digits of a number as the locale writes them (no sign): plain, with an exponent, grouped, or a whole number
// and a fraction ("1 1/2"). Undefined for anything else.
function unsignedNumber(text: string, reader: TextNumberReader): number | undefined {
	if (reader.number.test(text)) return Number(reader.decimal === "." ? text : text.replace(reader.decimal, "."))
	if (reader.grouped.test(text)) return Number(text.replace(reader.group, "").replace(reader.decimal, "."))

	const space = text.indexOf(" ")
	const slash = text.indexOf("/")

	if (space > 0 && slash > space) {
		const whole = text.slice(0, space)
		const numerator = text.slice(space + 1, slash)
		const denominator = text.slice(slash + 1)

		if (DIGITS.test(whole) && DIGITS.test(numerator) && DIGITS.test(denominator) && Number(denominator) !== 0) {
			return Number(whole) + Number(numerator) / Number(denominator)
		}
	}

	return undefined
}

const DIGITS = /^\d+$/

// A number given as text where a formula takes a number ("1,000"*1, ABS("$4.00")), read as Excel reads text "that
// looks like an accepted number format" (support.microsoft.com "Calculation operators and precedence in Excel":
// ="1"+"2" is 3, =1+"$4.00" is 5) in the locale: its decimal point and grouping, a sign before or after (5- is -5),
// one of the locale's currency symbols before or after the number, with the sign on either side of it, a percent sign after,
// parentheses around a negative number, an exponent and a whole number with a fraction; surrounding spaces are
// ignored. Where the page is silent, as LibreOffice 7.5 reads text in en-US and de-DE. Dates and times are read
// apart (formula/textDates.ts). Undefined when the text reads as no number, most of it at its first character.
export function textNumber(text: string, locale: SheetLocale): number | undefined {
	return readTextNumber(text, textNumberReader(locale))
}

// textNumber for one locale, its reader resolved once (the engine reads every text operand through it).
export function textNumberParser(locale: SheetLocale): (text: string) => number | undefined {
	const reader = textNumberReader(locale)

	return text => readTextNumber(text, reader)
}

function readTextNumber(text: string, reader: TextNumberReader): number | undefined {
	let start = 0
	let end = text.length

	while (start < end && isSpace(text.charCodeAt(start))) start++
	while (end > start && isSpace(text.charCodeAt(end - 1))) end--

	if (start === end) return undefined

	const first = text.charCodeAt(start)
	let whole = true

	for (let index = start; whole && index < end; index++) whole = isDigit(text.charCodeAt(index))

	if (whole) {
		const value = Number(start === 0 && end === text.length ? text : text.slice(start, end))

		// Digits past what a number holds are no number: #VALUE!, not an overflow.
		return Number.isFinite(value) ? value : undefined
	}

	if (
		!isDigit(first) &&
		first !== PLUS &&
		first !== MINUS &&
		first !== PAREN &&
		first !== reader.decimalCode &&
		!reader.currencyCodes.includes(first)
	) {
		return undefined
	}

	let body = text.slice(start, end)
	let negative = false
	let signed = false

	// (1,000) and ($5): a negative number, with no sign or percent inside.
	const parenthesised = first === PAREN

	if (parenthesised) {
		if (body.charCodeAt(body.length - 1) !== 41) return undefined

		body = body.slice(1, -1).trim()
		negative = true
		signed = true
	}

	const sign = (): boolean => {
		const code = body.charCodeAt(0)

		if (code !== PLUS && code !== MINUS) return true
		if (signed) return false

		signed = true
		negative = code === MINUS
		body = body.slice(1)

		return true
	}

	if (!sign()) return undefined

	let currency = false
	const leading = reader.currencies.find(symbol => body.startsWith(symbol))

	if (leading !== undefined) {
		currency = true
		body = body.slice(leading.length).trimStart()

		if (!sign()) return undefined
	}

	if (!parenthesised && body.charCodeAt(body.length - 1) === MINUS) {
		if (signed) return undefined

		signed = true
		negative = true
		body = body.slice(0, -1).trimEnd()
	}

	let percent = false

	const trailing = currency ? undefined : reader.currencies.find(symbol => body.endsWith(symbol))

	if (trailing !== undefined) {
		currency = true
		body = body.slice(0, -trailing.length).trimEnd()
	} else if (!parenthesised && !currency && body.endsWith("%")) {
		percent = true
		body = body.slice(0, -1).trimEnd()
	}

	const value = unsignedNumber(body, reader)

	if (value === undefined) return undefined

	if (!Number.isFinite(value)) return undefined

	const scaled = percent ? value / 100 : value

	return negative ? -scaled : scaled
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

// A territory's current currency (ISO 3166-1 alpha-2 → ISO 4217, Unicode CLDR: currencies.generated.ts), for the
// app to name the locale's currency symbols. Undefined for a territory without one.
export function territoryCurrency(territory: string): string | undefined {
	if (territory.length !== 2) return undefined

	const key = territory.toUpperCase()
	let low = 0
	let high = TERRITORY_CURRENCIES.length / 5 - 1

	while (low <= high) {
		const middle = (low + high) >> 1
		const at = TERRITORY_CURRENCIES.slice(middle * 5, middle * 5 + 2)

		if (at === key) return TERRITORY_CURRENCIES.slice(middle * 5 + 2, middle * 5 + 5)
		if (at < key) low = middle + 1
		else high = middle - 1
	}

	return undefined
}
