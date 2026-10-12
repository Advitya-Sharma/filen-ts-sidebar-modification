// The text a date content control shows for its date (ECMA-376 Part 1 §17.5.2.7 w:date): w:dateFormat is a
// Word date-time picture (the same syntax as a field's \@ switch, §17.16.4.1) read in w:lid's language and
// w:calendar's calendar.
//
// Picture items: d, dd (day), ddd, dddd (weekday name), M, MM (month), MMM, MMMM (month name), y, yy (two-digit
// year), yyy, yyyy (year), h, hh, H, HH, m, mm, s, ss (a date control's time is midnight), AM/PM, am/pm, A/P,
// a/p; 'quoted' text is literal, as is anything else.

// Word's w:calendar values as Intl calendars (Korea counts the Tangun era on the Gregorian calendar).
const CALENDARS: Readonly<Record<string, string>> = {
	gregorian: "gregory",
	gregorianUs: "gregory",
	gregorianMeFrench: "gregory",
	gregorianArabic: "gregory",
	gregorianXlitEnglish: "gregory",
	gregorianXlitFrench: "gregory",
	hijri: "islamic",
	hebrew: "hebrew",
	japan: "japanese",
	taiwan: "roc",
	thai: "buddhist",
	saka: "indian",
	korea: "gregory"
}

const TANGUN_OFFSET = 2333

interface Parts {
	readonly year: number
	readonly month: number
	readonly day: number
}

const formats = new Map<string, Intl.DateTimeFormat>()

function format(locale: string, options: Intl.DateTimeFormatOptions): Intl.DateTimeFormat {
	const key = `${locale}|${JSON.stringify(options)}`
	let found = formats.get(key)

	if (found === undefined) {
		found = new Intl.DateTimeFormat(locale, { timeZone: "UTC", ...options })
		formats.set(key, found)
	}

	return found
}

// A BCP 47 tag Intl accepts for w:lid (ST_Lang: a language tag or a hex LCID, which reads as English).
export function dateLocale(lid: string | null): string {
	if (lid === null || !/^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*$/.test(lid)) return "en-US"

	try {
		return Intl.DateTimeFormat.supportedLocalesOf([lid]).length > 0 ? lid : "en-US"
	} catch {
		return "en-US"
	}
}

function calendarOf(calendar: string | null): string {
	return calendar === null ? "gregory" : (CALENDARS[calendar] ?? "gregory")
}

function partsIn(date: Date, locale: string, calendar: string, korea: boolean): Parts {
	if (calendar === "gregory") {
		const year = date.getUTCFullYear()

		return { year: korea ? year + TANGUN_OFFSET : year, month: date.getUTCMonth() + 1, day: date.getUTCDate() }
	}

	const parts = format(`${locale}-u-ca-${calendar}`, { year: "numeric", month: "numeric", day: "numeric" }).formatToParts(date)
	const number = (type: string): number => {
		const value =
			parts.find(part => (part.type as string) === type || (type === "year" && (part.type as string) === "relatedYear"))?.value ?? ""
		const parsed = Number.parseInt(value.replace(/[^\d]/g, ""), 10)

		return Number.isFinite(parsed) ? parsed : 0
	}

	return { year: number("year"), month: number("month"), day: number("day") }
}

function name(date: Date, locale: string, calendar: string, field: "month" | "weekday", style: "long" | "short"): string {
	const options: Intl.DateTimeFormatOptions = field === "month" ? { month: style } : { weekday: style }
	const tag = calendar === "gregory" ? locale : `${locale}-u-ca-${calendar}`

	return format(tag, options).format(date)
}

function pad(value: number, width: number): string {
	return String(value).padStart(width, "0")
}

// Runs of one picture letter (case-sensitive for M and m; y, d, h, s either case).
function run(picture: string, at: number): number {
	const letter = picture[at]
	let end = at + 1

	while (end < picture.length && picture[end] === letter) end++

	return end - at
}

// `iso` (yyyy-mm-dd) as Word shows it with `picture` (Word's default short date of the language without one).
export function formatControlDate(iso: string, picture: string | null, lid: string | null, calendar: string | null = null): string {
	const date = new Date(`${iso}T00:00:00Z`)

	if (Number.isNaN(date.getTime())) return iso

	const locale = dateLocale(lid)
	const cal = calendarOf(calendar)

	if (picture === null || picture === "") return format(locale, { year: "numeric", month: "numeric", day: "numeric" }).format(date)

	const parts = partsIn(date, locale, cal, calendar === "korea")
	let out = ""

	for (let at = 0; at < picture.length;) {
		const char = picture[at] ?? ""

		if (char === "'") {
			const close = picture.indexOf("'", at + 1)

			out += close < 0 ? picture.slice(at + 1) : picture.slice(at + 1, close)
			at = close < 0 ? picture.length : close + 1

			continue
		}

		const rest = picture.slice(at)
		const ampm = /^(AM\/PM|am\/pm|A\/P|a\/p)/.exec(rest)

		if (ampm !== null) {
			// Midnight.
			const token = ampm[1] ?? ""

			out += token === "AM/PM" ? "AM" : token === "am/pm" ? "am" : token === "A/P" ? "A" : "a"
			at += token.length

			continue
		}

		const count = run(picture, at)

		switch (char) {
			case "d":
			case "D":
				out +=
					count === 1
						? String(parts.day)
						: count === 2
							? pad(parts.day, 2)
							: name(date, locale, cal, "weekday", count === 3 ? "short" : "long")
				break
			case "M":
				out +=
					count === 1
						? String(parts.month)
						: count === 2
							? pad(parts.month, 2)
							: name(date, locale, cal, "month", count === 3 ? "short" : "long")
				break
			case "y":
			case "Y":
				out += count <= 2 ? pad(parts.year % 100, 2) : String(parts.year)
				break
			case "h":
				out += "12"
				break
			case "H":
			case "m":
			case "s":
			case "S":
				out += count === 1 ? "0" : "00"
				break
			default:
				out += picture.slice(at, at + count)
		}

		at += count
	}

	return out
}

// The value a date bound to custom XML stores (w:storeMappedDataAs).
export function storedDate(iso: string, storeAs: "dateTime" | "date" | "text", shown: string): string {
	return storeAs === "date" ? iso : storeAs === "text" ? shown : `${iso}T00:00:00`
}
