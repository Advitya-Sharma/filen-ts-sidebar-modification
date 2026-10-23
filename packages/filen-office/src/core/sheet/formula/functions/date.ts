import { FunctionArgumentType, type ImplementedFunctions } from "hyperformula"
import { MAX_SERIAL_1900, MAX_SERIAL_1904, serialDate, type DateParts } from "../../numfmt"
import { dateSerial, type DateBase } from "../../numfmt/date"
import { dateBaseOf, numError, SheetFunctionPlugin, type PluginCall, type PluginState } from "./plugin"

// The parts of a date serial as Excel reads them, and the differences of two dates (DATEDIF, DAYS360), through the
// same calendar the cells show: in the 1900 system 0 is 1900-01-00 and 60 is 1900-02-29 (ECMA-376 Part 1
// §18.17.4.1), where the engine read 0 as 1899-12-31.

const DATE_PART = { parameters: [{ argumentType: FunctionArgumentType.NUMBER, minValue: 0 }] }

const DATE = { argumentType: FunctionArgumentType.NUMBER, minValue: 0 }

// The days of a month in a date system (the 1900 system's February 1900 has 29).
function monthDays(year: number, month: number, base: DateBase): number {
	return month === 12 ? 31 : dateSerial(year, month + 1, 1, base) - dateSerial(year, month, 1, base)
}

// DATEDIF(start_date, end_date, unit) ("DATEDIF function"; ECMA-376 Part 1 §18.17.7.75): the dates truncated, the
// unit in either case (the engine took only upper case, so "m" was #NUM!), a start after the end #NUM!. "MD" counts
// from the end date's previous month when its day is earlier (Excel's documented limitation: it can be negative),
// "YD" from the start date's day in the latest year that is not after the end date (as LibreOffice 7.5 counts).
function dateDifference(start: DateParts, end: DateParts, from: number, to: number, unit: string, base: DateBase): number | undefined {
	const dayBefore = end.day < start.day ? 1 : 0

	switch (unit.toUpperCase()) {
		case "D":
			return to - from
		case "M":
			return (end.year - start.year) * 12 + end.month - start.month - dayBefore
		case "Y":
			return end.year - start.year - (end.month < start.month || (end.month === start.month && dayBefore === 1) ? 1 : 0)
		case "YM":
			return (end.month - start.month - dayBefore + 12) % 12
		case "MD": {
			if (dayBefore === 0) return end.day - start.day

			const january = end.month === 1

			return monthDays(january ? end.year - 1 : end.year, january ? 12 : end.month - 1, base) + end.day - start.day
		}
		case "YD": {
			const before = end.month < start.month || (end.month === start.month && dayBefore === 1)

			return to - dateSerial(before ? end.year - 1 : end.year, start.month, start.day, base)
		}
		default:
			return undefined
	}
}

export class DateFunctions extends SheetFunctionPlugin {
	static override implementedFunctions: ImplementedFunctions = {
		DAY: { method: "day", ...DATE_PART },
		MONTH: { method: "month", ...DATE_PART },
		YEAR: { method: "year", ...DATE_PART },
		DATEDIF: { method: "datedif", parameters: [DATE, DATE, { argumentType: FunctionArgumentType.STRING }] },
		DAYS360: { method: "days360", parameters: [DATE, DATE, { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false }] }
	}

	private readonly parts: DateParts = { year: 0, month: 0, day: 0, weekday: 0 }
	private readonly other: DateParts = { year: 0, month: 0, day: 0, weekday: 0 }

	// Reads two whole dates of the workbook's date system into parts and other: false past its last date.
	private load(from: number, to: number, base: DateBase): boolean {
		const max = base === true ? MAX_SERIAL_1904 : MAX_SERIAL_1900

		if (from > max || to > max) return false

		serialDate(from, base, this.parts)
		serialDate(to, base, this.other)

		return true
	}

	private part(ast: PluginCall, state: PluginState, name: string, pick: (parts: DateParts) => number): unknown {
		const date1904 = dateBaseOf(this.config)

		return this.runFunction(ast.args, state, this.metadata(name), (serial: number) => {
			const day = Math.floor(serial)

			if (day > (date1904 === true ? MAX_SERIAL_1904 : MAX_SERIAL_1900)) return numError()

			return pick(serialDate(day, date1904, this.parts))
		})
	}

	day(ast: PluginCall, state: PluginState): unknown {
		return this.part(ast, state, "DAY", parts => parts.day)
	}

	month(ast: PluginCall, state: PluginState): unknown {
		return this.part(ast, state, "MONTH", parts => parts.month)
	}

	year(ast: PluginCall, state: PluginState): unknown {
		return this.part(ast, state, "YEAR", parts => parts.year)
	}

	datedif(ast: PluginCall, state: PluginState): unknown {
		const base = dateBaseOf(this.config)

		return this.runFunction(ast.args, state, this.metadata("DATEDIF"), (start: number, end: number, unit: string) => {
			const from = Math.floor(start)
			const to = Math.floor(end)

			if (from > to || !this.load(from, to, base)) return numError()

			return dateDifference(this.parts, this.other, from, to, unit, base) ?? numError()
		})
	}

	// DAYS360(start_date, end_date, [method]) ("DAYS360 function"; ECMA-376 Part 1 §18.17.7.79). The U.S. method
	// moves a start date on its month's last day (February's included) to the 30th before it reads the end date,
	// whose 31st becomes the 30th only after a start on the 30th (the engine read the start's 31st and 30th before
	// moving February's last day, so DAYS360(2011-02-28, 2011-03-31) was 31 for the page's 30); an end on
	// February's last day after a start on it is the 30th too. The European method moves every 31st to the 30th.
	days360(ast: PluginCall, state: PluginState): unknown {
		const base = dateBaseOf(this.config)

		return this.runFunction(ast.args, state, this.metadata("DAYS360"), (start: number, end: number, european: boolean) => {
			if (!this.load(Math.floor(start), Math.floor(end), base)) return numError()

			const from = this.parts
			const to = this.other
			let fromDay = from.day
			let toDay = to.day

			if (european) {
				if (fromDay === 31) fromDay = 30
				if (toDay === 31) toDay = 30
			} else {
				const fromFebruaryEnd = from.month === 2 && fromDay === monthDays(from.year, 2, base)

				if (fromFebruaryEnd && to.month === 2 && toDay === monthDays(to.year, 2, base)) toDay = 30
				if (fromFebruaryEnd || fromDay === 31) fromDay = 30
				if (fromDay === 30 && toDay === 31) toDay = 30
			}

			return 360 * (to.year - from.year) + 30 * (to.month - from.month) + toDay - fromDay
		})
	}
}
