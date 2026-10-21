import {
	localizedDate,
	localizedNumber,
	type SheetCommand,
	type SheetLocale,
	type SheetRange,
	type SeriesRequest
} from "@filen/office/core/sheet"

// The Series dialog's state (Home > Fill > Series in Excel) and the command it sends.

export interface SeriesForm {
	readonly orientation: "rows" | "cols"
	readonly type: SeriesRequest["type"]
	readonly dateUnit: SeriesRequest["dateUnit"]
	readonly step: string
	readonly stop: string
	readonly trend: boolean
}

// Excel's defaults: a series along the selection's longer side, linear, stepping 1, no stop value.
export function seriesDefaults(range: SheetRange): SeriesForm {
	const rows = range.endRow - range.startRow + 1
	const cols = range.endCol - range.startCol + 1

	return { orientation: cols > rows ? "rows" : "cols", type: "linear", dateUnit: "day", step: "1", stop: "", trend: false }
}

// A number as typed in the step field, in the session's locale as every entry is; null when it is not one.
export function seriesStep(text: string, locale: SheetLocale): number | null {
	const value = localizedNumber(text.trim(), locale)?.value

	return value !== undefined && Number.isFinite(value) ? value : null
}

// Whether the stop field holds what the session reads: a number or a date in the locale, an ISO date and time, or a
// time; empty is none.
export function seriesStopValid(text: string, locale: SheetLocale): boolean {
	const trimmed = text.trim()

	return (
		trimmed === "" ||
		seriesStep(trimmed, locale) !== null ||
		localizedDate(trimmed, locale) !== null ||
		/^(?:\d{4}-\d{2}-\d{2}[ T])?\d{1,2}:\d{2}(?::\d{2})?$/.test(trimmed)
	)
}

// The trend fits linear and growth series only, as Excel offers it.
export function seriesTrendOffered(form: SeriesForm): boolean {
	return form.type === "linear" || form.type === "growth"
}

export function seriesCommand(form: SeriesForm, sheet: number, range: SheetRange, locale: SheetLocale): SheetCommand | null {
	const step = seriesStep(form.step, locale)

	if (step === null || !seriesStopValid(form.stop, locale)) {
		return null
	}

	return {
		type: "fill.series",
		sheet,
		range,
		orientation: form.orientation,
		options: {
			type: form.type,
			dateUnit: form.dateUnit,
			step,
			stop: form.stop.trim() === "" ? null : form.stop.trim(),
			trend: form.trend && seriesTrendOffered(form)
		}
	}
}
