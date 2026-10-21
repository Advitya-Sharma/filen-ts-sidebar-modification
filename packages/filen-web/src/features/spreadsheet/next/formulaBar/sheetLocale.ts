import type { SheetLocale } from "@filen/office/core/sheet"

// The locale entries are typed in, from the browser's formatting locale (the core reads no Intl locale; the app
// injects it): its decimal point and grouping, its date order, and ";" between arguments where "," is the decimal
// point, as Excel types formulas there. Files keep en-US syntax whatever it is.
export function sheetLocaleOf(language: string): SheetLocale {
	let decimal = "."
	let group = ","
	let dateOrder: SheetLocale["dateOrder"] = "mdy"

	try {
		for (const part of new Intl.NumberFormat(language).formatToParts(12345.6)) {
			if (part.type === "decimal") decimal = part.value
			else if (part.type === "group") group = part.value
		}

		const order = new Intl.DateTimeFormat(language, { year: "numeric", month: "2-digit", day: "2-digit" })
			.formatToParts(new Date(2000, 10, 22))
			.map(part => part.type)
			.filter(type => type === "year" || type === "month" || type === "day")
			.map(type => type.charAt(0))
			.join("")

		if (order === "dmy" || order === "ymd" || order === "mdy") dateOrder = order
	} catch {
		// An unknown language tag: en-US's.
	}

	return { decimal, group, argument: decimal === "," ? ";" : ",", dateOrder }
}

// The browser's formatting language.
export function browserLanguage(): string {
	return typeof navigator === "undefined" ? "en-US" : (navigator.languages[0] ?? navigator.language)
}
