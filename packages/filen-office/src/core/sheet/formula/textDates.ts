import type { ConfigParams } from "hyperformula"
import { EN_US_LOCALE, localizedDate } from "../input/locale"
import type { SheetLocale } from "../protocol"

// Text read as a date or time inside formulas (DATEVALUE("8/22/2011"), MONTH("5/1/2020"), "1/2/2020"+1), as a typed
// entry reads it (input/locale.ts): the locale's day, month and year order, ISO dates in every locale, and a time
// of hours and minutes (seconds optional, AM/PM optional) alone or after the date. Excel reads text dates by the
// system's date settings (support.microsoft.com "DATEVALUE function"); the engine's own default is day first.

type TextDateTime = ReturnType<ConfigParams["parseDateTime"]>

const TIME = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?(?: ?([ap])m?)?$/i

function timeOf(text: string): TextDateTime {
	const match = TIME.exec(text)

	if (match === null) return undefined

	const [, hourText = "", minuteText = "", secondText, half] = match
	const hours = Number(hourText)
	const minutes = Number(minuteText)
	const seconds = secondText === undefined ? 0 : Number(secondText)

	if (minutes > 59 || seconds >= 60) return undefined

	if (half === undefined) return { hours, minutes, seconds }
	if (hours > 12) return undefined

	return { hours: (hours % 12) + (half.toLowerCase() === "p" ? 12 : 0), minutes, seconds }
}

export function textDateParser(locale: SheetLocale = EN_US_LOCALE): ConfigParams["parseDateTime"] {
	return (text: string): TextDateTime => {
		const trimmed = text.trim().replace(/\s+/g, " ")
		const date = localizedDate(trimmed, locale)

		if (date !== null) return date

		const time = timeOf(trimmed)

		if (time !== undefined) return time

		const space = trimmed.indexOf(" ")

		if (space < 0) return undefined

		const day = localizedDate(trimmed.slice(0, space), locale)
		const clock = timeOf(trimmed.slice(space + 1))

		return day === null || clock === undefined ? undefined : { ...day, ...clock }
	}
}
