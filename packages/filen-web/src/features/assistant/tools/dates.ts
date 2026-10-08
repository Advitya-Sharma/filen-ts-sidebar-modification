import {
	dateFilterNote,
	dayRangeBounds,
	describeDayRange,
	inDayRange,
	localeWeekStart,
	resolveDateFilter,
	type AssistantCalendar,
	type AssistantPeriodFilter
} from "@filen/shared/assistant"
import { AssistantToolError } from "@/features/assistant/tools/format"
import { optionalString } from "@/features/assistant/tools/entities"

export type UserCalendar = AssistantCalendar & { timeZone: string; weekStart: number }

// The user's clock, zone and first weekday: what the prompt's date line counts days by. Read once per turn
// (session.ts) and handed to every tool call, so the line and the filters agree across midnight.
export function userCalendar(): UserCalendar {
	const { timeZone, locale } = Intl.DateTimeFormat().resolvedOptions()

	return { now: Date.now(), timeZone, weekStart: localeWeekStart(locale) }
}

// A call's date filter as code resolves it: the date it looks at (one of `fields`, the first by default) and
// the instants [start, end) its days cover, null without one.
export interface DateSelection {
	field: string
	bounds: { start: number; end: number } | null
	range: string | null
}

export function dateSelection(
	input: Record<string, unknown>,
	fields: readonly [string, ...string[]],
	calendar: UserCalendar
): DateSelection {
	const filter = resolveDateFilter(
		{
			// The loop validated it against the enum.
			period: optionalString(input, "period") as AssistantPeriodFilter | undefined,
			from: optionalString(input, "from"),
			to: optionalString(input, "to")
		},
		calendar
	)

	if (!filter.ok) {
		throw new AssistantToolError(filter.error, { invalid: true })
	}

	const date = optionalString(input, "date")

	return {
		field: date !== undefined && fields.includes(date) ? date : fields[0],
		bounds: filter.range === null ? null : dayRangeBounds(filter.range, calendar.timeZone),
		range: filter.range === null ? null : describeDayRange(filter.range)
	}
}

export function selects(selection: DateSelection, timestamp: number): boolean {
	return selection.bounds === null || inDayRange(timestamp, selection.bounds)
}

// Tells the model which days a filtered result covers, so it neither recounts them nor reads the list as
// everything.
export function datesField(selection: DateSelection): { dates?: string } {
	return selection.range === null ? {} : { dates: `${selection.field} ${selection.range}` }
}

// The note a date-filtered result carries when the filter alone emptied it.
export function emptyDatesNote(selection: DateSelection, kept: number, outside: number): { note?: string } {
	return selection.range !== null && kept === 0 && outside > 0 ? { note: dateFilterNote(selection.range, outside) } : {}
}
