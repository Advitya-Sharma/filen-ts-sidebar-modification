// The text a date content control shows for its date (ECMA-376 Part 1 §17.5.2.7 w:date): w:dateFormat is a
// Word date-time picture (the same syntax as a field's \@ switch, §17.16.4.1) read in w:lid's language and
// w:calendar's calendar.
//
// Picture items: d, dd (day), ddd, dddd (weekday name), M, MM (month), MMM, MMMM (month name; with a day in the
// picture the name as a date has it, the genitive of Slavic languages), y, yy (two-digit year), yyy, yyyy (year),
// g (the era's short name), gg (its first character), ggg (its name), e, ee (the year of the era), B, BB, BBBB (the
// Thai Buddhist year), h, hh, H, HH, m, mm, s, ss (a date control's time is midnight), AM/PM, am/pm, A/P, a/p;
// 'quoted' text is literal, as is anything else.

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
const BUDDHIST_OFFSET = 543

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

function tagOf(locale: string, calendar: string): string {
	return calendar === "gregory" ? locale : `${locale}-u-ca-${calendar}`
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

// A month name as a date shows it (`dated`: next to a day, the genitive form Slavic languages use), or a weekday's,
// of `date` (its UTC fields) in `calendar`. English names are Word's own (three letters short: "Sep", not "Sept").
function name(date: Date, locale: string, calendar: string, field: "month" | "weekday", style: "long" | "short", dated = false): string {
	if (calendar === "gregory" && /^en(?:-|$)/i.test(locale)) {
		const full = (field === "month" ? MONTHS[date.getUTCMonth()] : DAYS[date.getUTCDay()]) ?? ""

		return style === "long" ? full : full.slice(0, 3)
	}

	if (field === "month" && dated) {
		const part = format(tagOf(locale, calendar), { day: "numeric", month: style })
			.formatToParts(date)
			.find(entry => entry.type === "month")

		if (part !== undefined) return part.value
	}

	const options: Intl.DateTimeFormatOptions = field === "month" ? { month: style } : { weekday: style }

	return format(tagOf(locale, calendar), options).format(date)
}

// Whether a date picture shows the day of the month (d, dd): month names then take a date's form.
export function picturesDay(picture: string): boolean {
	return /(^|[^dD])[dD]{1,2}(?![dD])/.test(picture.replace(/'[^']*'?/g, ""))
}

// The name of month `month` (1-12) or of weekday `weekday` (0: Sunday) in language `lid` (a w:lid or w:lang value),
// as Word shows it in a date picture: shared by date content controls and date fields.
export function monthName(month: number, lid: string | null | undefined, style: "long" | "short", dated: boolean): string {
	return name(new Date(Date.UTC(2001, month - 1, 15)), dateLocale(lid ?? null), "gregory", "month", style, dated)
}

export function dayName(weekday: number, lid: string | null | undefined, style: "long" | "short"): string {
	// 2001-01-07 was a Sunday.
	return name(new Date(Date.UTC(2001, 0, 7 + weekday)), dateLocale(lid ?? null), "gregory", "weekday", style)
}

// The era of a date in a calendar (its name, short name) and the year within it.
function eraOf(date: Date, locale: string, calendar: string, style: "long" | "narrow"): { name: string; year: number } {
	const parts = format(tagOf(locale, calendar), { era: style, year: "numeric" }).formatToParts(date)
	const year = Number.parseInt((parts.find(part => part.type === "year")?.value ?? "").replace(/[^\d]/g, ""), 10)

	return { name: parts.find(part => part.type === "era")?.value ?? "", year: Number.isFinite(year) ? year : date.getUTCFullYear() }
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
	const dated = picturesDay(picture)
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

			// AM/PM either case shows upper case, as in a field (Microsoft's "Format field results"); A/P keeps its case.
			out += token === "A/P" ? "A" : token === "a/p" ? "a" : "AM"
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
							: name(date, locale, cal, "month", count === 3 ? "short" : "long", dated)
				break
			case "y":
			case "Y":
				out += count <= 2 ? pad(parts.year % 100, 2) : String(parts.year)
				break
			case "g": {
				const era = eraOf(date, locale, cal, count === 1 ? "narrow" : "long")

				out += count === 2 ? [...era.name].slice(0, 1).join("") : era.name
				break
			}
			case "e":
				out += count === 1 ? String(eraOf(date, locale, cal, "long").year) : pad(eraOf(date, locale, cal, "long").year % 100, 2)
				break
			case "B": {
				const year = date.getUTCFullYear() + BUDDHIST_OFFSET

				out += count <= 2 ? pad(year % 100, 2) : String(year)
				break
			}
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

const ESCAPE = /[.*+?^${}()|[\]\\]/g

function names(locale: string, field: "month" | "weekday"): string[] {
	const out: string[] = []

	for (let k = 0; k < (field === "month" ? 12 : 7); k++) {
		const date = new Date(Date.UTC(2026, field === "month" ? k : 0, field === "month" ? 1 : 4 + k))

		for (const style of ["long", "short"] as const)
			out.push(name(date, locale, "gregory", field, style, false), name(date, locale, "gregory", field, style, true))
	}

	return out
}

// The date (yyyy-mm-dd) text typed into a date control means, read with its picture in its language, or null:
// what Word cannot read as a date it keeps as text (ECMA-376 Part 1 §17.5.2.40: a binding stores the shown text
// when there is no date to translate). Gregorian pictures only; an ISO date is read too.
export function parseControlDate(text: string, picture: string | null, lid: string | null, calendar: string | null = null): string | null {
	const trimmed = text.trim()
	const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed)

	if (iso !== null) return validDate(Number(iso[1]), Number(iso[2]), Number(iso[3]))
	if (picture === null || picture === "" || calendarOf(calendar) !== "gregory" || calendar === "korea") return null

	const locale = dateLocale(lid)
	const months = names(locale, "month")
	const groups: ("day" | "month" | "monthName" | "year2" | "year")[] = []
	let pattern = ""

	for (let at = 0; at < picture.length;) {
		const char = picture[at] ?? ""

		if (char === "'") {
			const close = picture.indexOf("'", at + 1)

			pattern += (close < 0 ? picture.slice(at + 1) : picture.slice(at + 1, close)).replace(ESCAPE, "\\$&")
			at = close < 0 ? picture.length : close + 1

			continue
		}

		const count = run(picture, at)

		switch (char) {
			case "d":
			case "D":
				if (count <= 2) {
					pattern += "(\\d{1,2})"
					groups.push("day")
				} else
					pattern += `(?:${names(locale, "weekday")
						.map(entry => entry.replace(ESCAPE, "\\$&"))
						.join("|")})`
				break
			case "M":
				if (count <= 2) {
					pattern += "(\\d{1,2})"
					groups.push("month")
				} else {
					pattern += `(${[...months]
						.sort((a, b) => b.length - a.length)
						.map(entry => entry.replace(ESCAPE, "\\$&"))
						.join("|")})`
					groups.push("monthName")
				}
				break
			case "y":
			case "Y":
				pattern += count <= 2 ? "(\\d{2})" : "(\\d{4})"
				groups.push(count <= 2 ? "year2" : "year")
				break
			default:
				if (/\s/.test(char)) pattern += "\\s*"
				else if (/[A-Za-z]/.test(char)) return null
				else pattern += picture.slice(at, at + count).replace(ESCAPE, "\\$&")
		}

		at += count
	}

	const match = new RegExp(`^\\s*${pattern}\\s*$`, "iu").exec(trimmed)

	if (match === null) return null

	let day = Number.NaN
	let month = Number.NaN
	let year = Number.NaN

	groups.forEach((group, k) => {
		const value = match[k + 1] ?? ""

		if (group === "day") day = Number(value)
		else if (group === "month") month = Number(value)
		else if (group === "year") year = Number(value)
		else if (group === "year2") year = Number(value) + (Number(value) < 30 ? 2000 : 1900)
		else {
			const found = months.findIndex(entry => entry.toLocaleLowerCase(locale) === value.toLocaleLowerCase(locale))

			month = found < 0 ? Number.NaN : Math.floor(found / 4) + 1
		}
	})

	return validDate(year, month, day)
}

function validDate(year: number, month: number, day: number): string | null {
	if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day) || year < 1 || month < 1 || month > 12 || day < 1)
		return null

	const date = new Date(Date.UTC(year, month - 1, day))

	if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null

	return `${String(year).padStart(4, "0")}-${pad(month, 2)}-${pad(day, 2)}`
}
