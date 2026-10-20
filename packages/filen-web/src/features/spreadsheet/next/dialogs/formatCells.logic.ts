import { numFmt, numFmtCategory, type NumFmtCategory } from "@filen/office/core/sheet"

// The Number tab of Excel's Format Cells dialog as codes: each category's options (decimal places, the thousands
// separator, a currency symbol, how negatives show, a type from a list) and the code they make, and the options a
// code already in a cell reads as.

export interface FormatOptions {
	readonly category: NumFmtCategory
	readonly decimals: number
	readonly thousands: boolean
	// NEGATIVE_STYLES index.
	readonly negative: number
	readonly symbol: string
	// The chosen entry of the category's TYPES list.
	readonly type: number
	// The custom code as typed.
	readonly code: string
}

export const MAX_DECIMALS = 30

// Excel's currency symbols offered first; "" is none.
export const SYMBOLS: readonly string[] = ["$", "€", "£", "¥", ""]

// How negatives show (Excel's four): "-1,234.10", red "1,234.10", "(1,234.10)", red "(1,234.10)".
export const NEGATIVE_STYLES = 4

// The date, time and fraction types Excel lists (en-US); "*" marks those that follow the system's settings.
export const TYPES: Readonly<Partial<Record<NumFmtCategory, readonly string[]>>> = {
	date: [
		"m/d/yyyy",
		"[$-F800]dddd, mmmm dd, yyyy",
		"m/d;@",
		"m/d/yy;@",
		"mm/dd/yy;@",
		"d-mmm;@",
		"d-mmm-yy;@",
		"dd-mmm-yy;@",
		"mmm-yy;@",
		"mmmm-yy;@",
		"mmmm d, yyyy;@",
		"m/d/yy h:mm AM/PM;@",
		"m/d/yy h:mm;@",
		"yyyy-mm-dd;@"
	],
	time: ["[$-F400]h:mm:ss AM/PM", "h:mm;@", "h:mm AM/PM;@", "h:mm:ss;@", "h:mm:ss AM/PM;@", "mm:ss.0;@", "[h]:mm:ss;@", "m/d/yy h:mm;@"],
	fraction: ["# ?/?", "# ??/??", "# ???/???", "# ?/2", "# ?/4", "# ?/8", "# ??/16", "# ?/10", "# ??/100"]
}

function places(decimals: number): string {
	const count = Math.max(0, Math.min(MAX_DECIMALS, Math.round(decimals)))

	return count > 0 ? `.${"0".repeat(count)}` : ""
}

// The negative sections Excel writes for a positive section `base`.
function withNegative(base: string, negative: number): string {
	switch (negative) {
		case 1:
			return `${base};[Red]${base}`
		case 2:
			return `${base}_);(${base})`
		case 3:
			return `${base}_);[Red](${base})`
		default:
			return base
	}
}

function quoted(symbol: string): string {
	return symbol === "" ? "" : `"${symbol}"`
}

// The code the options make.
export function codeOf(options: FormatOptions): string {
	const decimals = places(options.decimals)

	switch (options.category) {
		case "general":
			return "General"
		case "number":
			return withNegative(`${options.thousands ? "#,##0" : "0"}${decimals}`, options.negative)
		case "currency":
			return withNegative(`${quoted(options.symbol)}#,##0${decimals}`, options.negative)
		case "accounting": {
			const symbol = quoted(options.symbol)
			const zero = decimals === "" ? "" : "?".repeat(decimals.length - 1)

			return `_(${symbol}* #,##0${decimals}_);_(${symbol}* \\(#,##0${decimals}\\);_(${symbol}* "-"${zero}_);_(@_)`
		}
		case "percentage":
			return `0${decimals}%`
		case "scientific":
			return `0${decimals}E+00`
		case "text":
			return "@"
		case "date":
		case "time":
		case "fraction":
			return TYPES[options.category]?.[options.type] ?? TYPES[options.category]?.[0] ?? "General"
		case "custom":
			return options.code
	}
}

// The options a cell's code reads as: its category, and the decimals, separator, symbol and negative style its
// first section shows.
export function optionsOf(code: string): FormatOptions {
	const category = numFmtCategory(code)
	const format = numFmt(code)
	const first = format.sections[0]
	const negative = format.sections[1]
	const symbol = SYMBOLS.find(entry => entry !== "" && code.includes(entry)) ?? (category === "accounting" ? "" : "$")
	const red = /\[red\]/iu.test(code)
	const parens = negative?.tokens.some(token => token.text === "(") === true
	const type = TYPES[category]?.indexOf(code) ?? -1

	return {
		// A date, time or fraction not in its list stays as the code it is.
		category: (category === "date" || category === "time" || category === "fraction") && type < 0 ? "custom" : category,
		decimals: Math.min(MAX_DECIMALS, first?.fractionDigits ?? 2),
		thousands: first?.grouping ?? false,
		negative: (red ? 1 : 0) + (parens ? 2 : 0),
		symbol,
		type: Math.max(0, type),
		code
	}
}

// What a category offers: decimal places, the thousands separator, a symbol, negative styles, a type list.
export function offers(category: NumFmtCategory): {
	decimals: boolean
	thousands: boolean
	symbol: boolean
	negative: boolean
	types: boolean
} {
	return {
		decimals:
			category === "number" ||
			category === "currency" ||
			category === "accounting" ||
			category === "percentage" ||
			category === "scientific",
		thousands: category === "number",
		symbol: category === "currency" || category === "accounting",
		negative: category === "number" || category === "currency",
		types: category === "date" || category === "time" || category === "fraction"
	}
}
