// Date serials (ECMA-376 Part 1 §18.17.4): days since the epoch, the time of day as the fraction.
// - 1900 system: 1 is 1900-01-01, 0 is shown as 1900-01-00, and 60 is 1900-02-29, a day the Gregorian calendar
//   does not have (§18.17.4.1: kept for compatibility, so serials from 61 on are one day off a plain count).
// - 1904 system: 0 is 1904-01-01.

export interface DateParts {
	year: number
	// 1-12.
	month: number
	// 0 only for the 1900 system's serial 0.
	day: number
	// 0 Sunday to 6 Saturday.
	weekday: number
}

// The largest serial a format shows as a date: 9999-12-31 (§18.17.4.1, §18.17.4.2).
export const MAX_SERIAL_1900 = 2958465
export const MAX_SERIAL_1904 = MAX_SERIAL_1900 - 1462

// Days from 1970-01-01 to a proleptic Gregorian date and back (civil calendar arithmetic, integers only).
function civil(days: number, out: DateParts): void {
	const z = days + 719468
	const era = Math.floor(z / 146097)
	const doe = z - era * 146097
	const yoe = Math.floor((doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365)
	const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100))
	const mp = Math.floor((5 * doy + 2) / 153)
	const month = mp < 10 ? mp + 3 : mp - 9

	out.year = yoe + era * 400 + (month <= 2 ? 1 : 0)
	out.month = month
	out.day = doy - Math.floor((153 * mp + 2) / 5) + 1
}

// 1899-12-30 and 1904-01-01 as days from 1970-01-01.
const EPOCH_1900 = -25569
const EPOCH_1904 = -24107

// The date of a whole serial day (>= 0, at most the system's maximum).
export function serialDate(day: number, date1904: boolean, out: DateParts): DateParts {
	if (date1904) {
		civil(day + EPOCH_1904, out)
		out.weekday = (day + 5) % 7

		return out
	}

	if (day === 0) {
		out.year = 1900
		out.month = 1
		out.day = 0
	} else if (day === 60) {
		out.year = 1900
		out.month = 2
		out.day = 29
	} else civil(day + EPOCH_1900 + (day < 60 ? 1 : 0), out)

	out.weekday = (day + 6) % 7

	return out
}
