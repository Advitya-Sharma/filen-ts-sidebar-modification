// Cell values as ODF writes them (ODF 1.3 Part 3 §19.379-§19.389): dates as xsd:date or xsd:dateTime, times as
// xsd:duration, both turned into serial numbers counted from the table's null date (§19.639).

// Days from 1970-01-01 to a calendar date (proleptic Gregorian, any year).
function epochDays(year: number, month: number, day: number): number {
	const y = month <= 2 ? year - 1 : year
	const era = Math.floor(y / 400)
	const yoe = y - era * 400
	const doy = Math.floor((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1
	const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy

	return era * 146_097 + doe - 719_468
}

const DATE = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?(?:Z|[+-]\d{2}:\d{2})?$/

// Digits of `text` from `at` as a number, -1 when one is not a digit.
function digits(text: string, at: number, count: number): number {
	let value = 0

	for (let k = at; k < at + count; k++) {
		const code = text.charCodeAt(k) - 48

		if (code < 0 || code > 9) return -1

		value = value * 10 + code
	}

	return value
}

// The days since 1970-01-01 of a date or date-time value (a fraction for the time of day), NaN for anything else.
export function dateValue(text: string): number {
	// What LibreOffice writes, read without a match: YYYY-MM-DD, or with Thh:mm:ss.
	if (
		(text.length === 10 || (text.length === 19 && text[10] === "T" && text[13] === ":" && text[16] === ":")) &&
		text[4] === "-" &&
		text[7] === "-"
	) {
		const year = digits(text, 0, 4)
		const month = digits(text, 5, 2)
		const day = digits(text, 8, 2)
		const timed = text.length === 19
		const hours = timed ? digits(text, 11, 2) : 0
		const minutes = timed ? digits(text, 14, 2) : 0
		const seconds = timed ? digits(text, 17, 2) : 0

		if (year >= 0 && month >= 1 && month <= 12 && day >= 1 && day <= 31 && hours >= 0 && minutes >= 0 && seconds >= 0)
			return epochDays(year, month, day) + (hours * 3600 + minutes * 60 + seconds) / 86_400
	}

	const match = DATE.exec(text.trim())

	if (match === null) return Number.NaN

	const month = Number(match[2])
	const day = Number(match[3])

	if (month < 1 || month > 12 || day < 1 || day > 31) return Number.NaN

	const days = epochDays(Number(match[1]), month, day)
	const seconds = match[4] === undefined ? 0 : Number(match[4]) * 3600 + Number(match[5]) * 60 + Number(match[6])

	return days + seconds / 86_400
}

const DURATION =
	/^(-)?P(?:(\d+(?:\.\d+)?)Y)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/

// An xsd:duration as days (years and months as 365 and 30 days, which no time value uses), NaN for anything else.
export function durationValue(text: string): number {
	const match = DURATION.exec(text.trim())

	if (match === null || text.trim() === "P" || text.trim() === "-P") return Number.NaN

	const part = (index: number): number => (match[index] === undefined ? 0 : Number(match[index]))
	const days = part(2) * 365 + part(3) * 30 + part(4) + (part(5) * 3600 + part(6) * 60 + part(7)) / 86_400

	return match[1] === "-" ? -days : days
}

// Serial numbers count days from the null date: LibreOffice's default is 1899-12-30, where serial 61 and on equal
// Excel's 1900 date system; 1904-01-01 is Excel's 1904 system.
export const NULL_DATE_1899 = epochDays(1899, 12, 30)
export const NULL_DATE_1904 = epochDays(1904, 1, 1)

export function nullDate(text: string | undefined): number {
	if (text === undefined) return NULL_DATE_1899

	const value = dateValue(text)

	return Number.isNaN(value) ? NULL_DATE_1899 : Math.floor(value)
}
