// Number formats for chart labels: the subset of SpreadsheetML format codes charts use (General, fixed
// decimals, thousands separators, percent, scientific, literal prefixes/suffixes, simple dates).

const MS_PER_DAY = 86400000
// Serial 0 of the 1900 date system is 1899-12-30 once Excel's phantom 1900-02-29 is accounted for.
const EPOCH_1900 = Date.UTC(1899, 11, 30)
const EPOCH_1904 = Date.UTC(1904, 0, 1)
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

function general(value: number): string {
	if (value === 0) return "0"

	const magnitude = Math.abs(value)

	if (magnitude >= 1e11 || magnitude < 1e-9) {
		return value
			.toExponential(5)
			.replace(/\.?0+e/, "E")
			.replace(/E([+-])(\d)$/, "E$10$2")
	}

	// Ten significant digits, trailing zeros dropped, as Excel's General shows them.
	const text = Number(value.toPrecision(10)).toString()

	return text.includes("e") ? value.toFixed(Math.max(0, 10 - Math.ceil(Math.log10(magnitude)))).replace(/\.?0+$/, "") : text
}

function groupThousands(integer: string): string {
	let out = ""

	for (let k = 0; k < integer.length; k++) {
		if (k > 0 && (integer.length - k) % 3 === 0) out += ","

		out += integer[k]
	}

	return out
}

// Literal text of a format section: quoted strings, escaped characters, and the characters Excel shows
// as themselves.
function literal(text: string): string {
	return text.replace(
		/"([^"]*)"|\\(.)|_.|\*.|\[[^\]]*\]/g,
		(_match, quoted: string | undefined, escaped: string | undefined) => quoted ?? escaped ?? ""
	)
}

function isDateCode(code: string): boolean {
	const bare = code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, "")

	return /[yd]/i.test(bare) || (/m/i.test(bare) && !/[0#?]/.test(bare))
}

function formatDate(value: number, code: string, date1904: boolean): string {
	const time = (date1904 ? EPOCH_1904 : EPOCH_1900) + Math.round(value * MS_PER_DAY)

	if (!Number.isFinite(time)) return general(value)

	const date = new Date(time)
	const year = date.getUTCFullYear()
	const month = date.getUTCMonth()
	const day = date.getUTCDate()
	const bare = code.replace(/\[[^\]]*\]/g, "")

	return bare.replace(
		/"([^"]*)"|\\(.)|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|h+|s+|AM\/PM/gi,
		(token: string, quoted: string | undefined, escaped: string | undefined) => {
			if (quoted !== undefined) return quoted
			if (escaped !== undefined) return escaped

			switch (token.toLowerCase()) {
				case "yyyy":
					return String(year)
				case "yy":
					return String(year % 100).padStart(2, "0")
				case "mmmmm":
					return (MONTHS[month] ?? "").slice(0, 1)
				case "mmmm":
					return MONTHS[month] ?? ""
				case "mmm":
					return (MONTHS[month] ?? "").slice(0, 3)
				case "mm":
					return String(month + 1).padStart(2, "0")
				case "m":
					return String(month + 1)
				case "dd":
					return String(day).padStart(2, "0")
				case "d":
					return String(day)
				default:
					// Times and weekday names are rare on chart axes and shown as nothing.
					return ""
			}
		}
	)
}

// Formats `value` with an Excel number format code. `date1904` selects the workbook's date system.
export function formatNumber(value: number, code: string, date1904: boolean = false): string {
	if (!Number.isFinite(value)) return ""

	const sections = code.split(";")
	const negative = value < 0 && sections.length > 1
	let section = (negative ? sections[1] : sections[0]) ?? "General"

	if (value === 0 && sections.length > 2) section = sections[2] ?? section

	// Locale/currency tags ([$€-407]) show their symbol; colours and conditions ([Red], [>100]) show nothing.
	const trimmed = section
		.replace(/\[\$([^\]-]*)(?:-[^\]]*)?\]/g, (_match, symbol: string) => `"${symbol.replace(/"/g, "")}"`)
		.replace(/\[[^\]]*\]/g, "")
		.trim()

	if (trimmed === "" || /^general$/i.test(trimmed) || trimmed === "@") return general(value)
	if (isDateCode(trimmed)) return formatDate(value, trimmed, date1904)

	const magnitude = Math.abs(value)
	const placeholders = /[0#?][0#?,]*(?:\.[0#?]*)?(?:[eE][+-][0#]+)?%?/.exec(
		trimmed.replace(/"[^"]*"|\\./g, match => " ".repeat(match.length))
	)

	if (placeholders === null) return literal(trimmed)

	const number = placeholders[0]
	const percent = trimmed.includes("%")
	const scaled = percent ? magnitude * 100 : magnitude
	const decimalsMatch = /\.([0#?]*)/.exec(number)
	const decimals = decimalsMatch === null ? 0 : (decimalsMatch[1] ?? "").length
	const exponent = /[eE][+-]/.test(number)
	let digits: string

	if (exponent) {
		digits = scaled.toExponential(decimals).toUpperCase()
		digits = digits.replace(/E([+-])(\d)$/, "E$10$2")
	} else {
		const fixed = scaled.toFixed(decimals)
		const [integer = "0", fraction] = fixed.split(".")
		const optional = decimalsMatch !== null && /#/.test(decimalsMatch[1] ?? "")
		let fractionText = fraction ?? ""

		if (optional) fractionText = fractionText.replace(/0+$/, "")

		const grouped = /,/.test(number.split(".")[0] ?? "") ? groupThousands(integer) : integer
		// An integer part without a "0" placeholder ("#", "??") shows no zero.
		const leading = !(number.split(".")[0] ?? "").includes("0") && integer === "0" ? "" : grouped

		digits = fractionText === "" ? leading : `${leading}.${fractionText}`
	}

	const start = placeholders.index
	const prefix = literal(trimmed.slice(0, start))
	const suffix = literal(trimmed.slice(start + number.length)).replace(/%/g, "")
	const sign = value < 0 && !negative ? "-" : ""

	return `${sign}${prefix}${digits}${percent ? "%" : ""}${suffix}`
}
