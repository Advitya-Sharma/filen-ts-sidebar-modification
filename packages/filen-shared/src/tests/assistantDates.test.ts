import { describe, expect, it } from "vitest"
import {
	dateFilterNote,
	dayRangeBounds,
	describeDayRange,
	inDayRange,
	localDayStartMs,
	localIsoDay,
	parseDayBound,
	periodRange,
	resolveDateFilter
} from "@filen/shared/assistant"

// Wednesday 2025-10-08, 09:30 in Berlin.
const calendar = { now: Date.UTC(2025, 9, 8, 7, 30), timeZone: "Europe/Berlin" }

describe("periodRange", () => {
	it.each([
		["today", "2025-10-08", "2025-10-08"],
		["yesterday", "2025-10-07", "2025-10-07"],
		["this_week", "2025-10-06", "2025-10-12"],
		["last_week", "2025-09-29", "2025-10-05"],
		["this_month", "2025-10-01", "2025-10-31"],
		["last_month", "2025-09-01", "2025-09-30"],
		["this_year", "2025-01-01", "2025-12-31"],
		["last_year", "2024-01-01", "2024-12-31"]
	] as const)("%s", (period, from, to) => {
		expect(periodRange(period, calendar)).toEqual({ from, to })
	})

	it("takes the day in the user's zone and steps back across a year", () => {
		// Still 2025-12-31 in New York.
		const nyc = { now: Date.UTC(2026, 0, 1, 3), timeZone: "America/New_York" }

		expect(periodRange("today", nyc)).toEqual({ from: "2025-12-31", to: "2025-12-31" })
		expect(periodRange("last_month", { now: Date.UTC(2026, 0, 15), timeZone: "UTC" })).toEqual({ from: "2025-12-01", to: "2025-12-31" })
		expect(periodRange("last_month", { now: Date.UTC(2024, 2, 10), timeZone: "UTC" })).toEqual({ from: "2024-02-01", to: "2024-02-29" })
	})

	it("starts weeks on the given weekday", () => {
		expect(periodRange("this_week", { ...calendar, weekStart: 7 })).toEqual({ from: "2025-10-05", to: "2025-10-11" })
		expect(periodRange("last_week", { ...calendar, weekStart: 6 })).toEqual({ from: "2025-09-27", to: "2025-10-03" })
	})
})

describe("parseDayBound", () => {
	it("reads days, months and years as their first or last day", () => {
		expect(parseDayBound("2025-10-08", "from")).toBe("2025-10-08")
		expect(parseDayBound(" 2025-2-3 ", "to")).toBe("2025-02-03")
		expect(parseDayBound("2024-02", "to")).toBe("2024-02-29")
		expect(parseDayBound("2024-02", "from")).toBe("2024-02-01")
		expect(parseDayBound("2024", "to")).toBe("2024-12-31")
		expect(parseDayBound("2025-10-08T12:00:00Z", "from")).toBe("2025-10-08")
	})

	it("rejects what is no real day", () => {
		for (const value of ["2025-02-30", "2025-13", "last week", "", "08.10.2025", "2025-10-08x"]) {
			expect(parseDayBound(value, "from"), value).toBeNull()
		}
	})
})

describe("resolveDateFilter", () => {
	it("is null without date arguments or for any, whatever days were filled in beside it", () => {
		expect(resolveDateFilter({}, calendar)).toEqual({ ok: true, range: null })
		expect(resolveDateFilter({ from: " ", to: "" }, calendar)).toEqual({ ok: true, range: null })
		expect(resolveDateFilter({ period: "any", from: "2025-10-01", to: "nonsense" }, calendar)).toEqual({ ok: true, range: null })
	})

	it("takes a named period over days filled in beside it", () => {
		expect(resolveDateFilter({ period: "last_week", from: "2025-10-05", to: "2025-10-30" }, calendar)).toEqual({
			ok: true,
			range: { from: "2025-09-29", to: "2025-10-05" }
		})
	})

	it("reads explicit days under custom or without a period, leaving open ends open", () => {
		expect(resolveDateFilter({ period: "custom", from: "2024", to: "2024-06" }, calendar)).toEqual({
			ok: true,
			range: { from: "2024-01-01", to: "2024-06-30" }
		})
		expect(resolveDateFilter({ from: "2025-09" }, calendar)).toEqual({ ok: true, range: { from: "2025-09-01", to: "9999-12-31" } })
	})

	it("refuses malformed, reversed or missing days", () => {
		expect(resolveDateFilter({ from: "yesterday" }, calendar)).toMatchObject({ ok: false })
		expect(resolveDateFilter({ period: "custom", to: "2025-02-30" }, calendar)).toMatchObject({ ok: false })
		expect(resolveDateFilter({ from: "2025-10-08", to: "2025-10-01" }, calendar)).toMatchObject({ ok: false })
		expect(resolveDateFilter({ period: "custom" }, calendar)).toMatchObject({ ok: false })
	})

	it("explains an empty filtered result", () => {
		expect(dateFilterNote("2025-10-07", 1)).toBe("Nothing is dated 2025-10-07; 1 item is outside those dates.")
		expect(dateFilterNote("2025-10-07", 3)).toBe("Nothing is dated 2025-10-07; 3 items are outside those dates.")
	})
})

describe("day bounds", () => {
	it("covers whole local days, across a DST change", () => {
		// Berlin leaves summer time on 2025-10-26, a 25-hour day.
		const bounds = dayRangeBounds({ from: "2025-10-26", to: "2025-10-26" }, "Europe/Berlin")

		expect(bounds).toEqual({ start: Date.UTC(2025, 9, 25, 22), end: Date.UTC(2025, 9, 26, 23) })
		expect(inDayRange(Date.UTC(2025, 9, 26, 22, 59), bounds)).toBe(true)
		expect(inDayRange(Date.UTC(2025, 9, 26, 23), bounds)).toBe(false)
		expect(localDayStartMs("2025-03-30", "Europe/Berlin")).toBe(Date.UTC(2025, 2, 29, 23))
	})

	it("leaves an open end unbounded", () => {
		const bounds = dayRangeBounds({ from: "2025-09-01", to: "9999-12-31" }, "UTC")

		expect(inDayRange(Date.UTC(2030, 0, 1), bounds)).toBe(true)
		expect(inDayRange(Date.UTC(2025, 7, 31), bounds)).toBe(false)
		expect(inDayRange(0, dayRangeBounds({ from: "0000-01-01", to: "2025-09-01" }, "UTC"))).toBe(true)
	})

	it("names a timestamp's local day", () => {
		expect(localIsoDay(Date.UTC(2025, 9, 7, 22, 30), "Europe/Berlin")).toBe("2025-10-08")
		expect(localIsoDay(Date.UTC(2025, 9, 7, 22, 30), "UTC")).toBe("2025-10-07")
		expect(localIsoDay(Number.NaN, "UTC")).toBe("")
	})

	it("describes a range without its open ends", () => {
		expect(describeDayRange({ from: "2025-10-01", to: "2025-10-31" })).toBe("2025-10-01 to 2025-10-31")
		expect(describeDayRange({ from: "2025-10-08", to: "2025-10-08" })).toBe("2025-10-08")
		expect(describeDayRange({ from: "2025-09-01", to: "9999-12-31" })).toBe("since 2025-09-01")
		expect(describeDayRange({ from: "0000-01-01", to: "2025-09-01" })).toBe("until 2025-09-01")
	})
})
