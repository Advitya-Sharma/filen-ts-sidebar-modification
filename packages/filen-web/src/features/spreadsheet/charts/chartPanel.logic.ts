import { rangeName, type SheetRange } from "@filen/office/core/sheet"
import {
	labelPositions,
	type ChartAxisInfo,
	type ChartEditType,
	type ChartGrouping,
	type ChartLabelPosition,
	type ChartLegendPosition,
	type ChartPlotInfo,
	type ChartType
} from "@filen/office/react/sheet/charts"
import type { PreviewKey } from "@/lib/i18n"

// What the chart editor panel offers and how it reads its fields: the chart types and layouts it makes, axis names,
// label positions every plot allows, and the number fields' text (empty is automatic).

export const CHART_TYPES: readonly { readonly value: ChartEditType; readonly key: PreviewKey }[] = [
	{ value: "column", key: "previewSheetChartTypeColumn" },
	{ value: "bar", key: "previewSheetChartTypeBar" },
	{ value: "line", key: "previewSheetChartTypeLine" },
	{ value: "area", key: "previewSheetChartTypeArea" },
	{ value: "pie", key: "previewSheetChartTypePie" },
	{ value: "doughnut", key: "previewSheetChartTypeDoughnut" },
	{ value: "scatter", key: "previewSheetChartTypeScatter" }
]

// The types a group of a combination chart can take (the session refuses the others there).
export const COMBO_TYPES: ReadonlySet<ChartEditType> = new Set(["column", "line", "area"])

export function typeKey(type: ChartType): PreviewKey {
	return (
		CHART_TYPES.find(entry => entry.value === type)?.key ??
		(type === "radar" ? "previewSheetChartTypeRadar" : "previewSheetChartTypeOther")
	)
}

export function editableType(type: ChartType): type is ChartEditType {
	return CHART_TYPES.some(entry => entry.value === type)
}

// The layouts a type has (none for pie, doughnut and scatter).
export function groupingOptions(type: ChartType): readonly { readonly value: ChartGrouping; readonly key: PreviewKey }[] {
	if (type === "column" || type === "bar")
		return [
			{ value: "clustered", key: "previewSheetChartGroupingClustered" },
			{ value: "stacked", key: "previewSheetChartGroupingStacked" },
			{ value: "percentStacked", key: "previewSheetChartGroupingPercent" }
		]

	if (type === "line" || type === "area")
		return [
			{ value: "standard", key: "previewSheetChartGroupingStandard" },
			{ value: "stacked", key: "previewSheetChartGroupingStacked" },
			{ value: "percentStacked", key: "previewSheetChartGroupingPercent" }
		]

	return []
}

export const LEGEND_OPTIONS: readonly { readonly value: ChartLegendPosition | "none"; readonly key: PreviewKey }[] = [
	{ value: "none", key: "previewSheetChartLegendNone" },
	{ value: "r", key: "previewSheetChartLegendRight" },
	{ value: "l", key: "previewSheetChartLegendLeft" },
	{ value: "t", key: "previewSheetChartLegendTop" },
	{ value: "b", key: "previewSheetChartLegendBottom" },
	{ value: "tr", key: "previewSheetChartLegendTopRight" }
]

const POSITION_KEYS: Readonly<Record<ChartLabelPosition, PreviewKey>> = {
	bestFit: "previewSheetChartPositionBestFit",
	ctr: "previewSheetChartPositionCenter",
	inEnd: "previewSheetChartPositionInsideEnd",
	inBase: "previewSheetChartPositionInsideBase",
	outEnd: "previewSheetChartPositionOutsideEnd",
	l: "previewSheetChartPositionLeft",
	r: "previewSheetChartPositionRight",
	t: "previewSheetChartPositionAbove",
	b: "previewSheetChartPositionBelow"
}

// Label positions every plot of the chart allows ("default": none written, each plot's own place).
export function labelPositionOptions(
	plots: readonly ChartPlotInfo[]
): readonly { readonly value: ChartLabelPosition | "default"; readonly key: PreviewKey }[] {
	let shared: readonly ChartLabelPosition[] | null = null

	for (const plot of plots) {
		const allowed = labelPositions(plot.type, plot.grouping)

		shared = shared === null ? allowed : shared.filter(position => allowed.includes(position))
	}

	return [
		{ value: "default", key: "previewSheetChartPositionDefault" },
		...(shared ?? []).map(value => ({ value, key: POSITION_KEYS[value] }))
	]
}

// An axis's name in the panel: horizontal or vertical by where it sits, "secondary" for the second of its direction.
export function axisKey(axis: ChartAxisInfo, axes: readonly ChartAxisInfo[]): PreviewKey {
	const horizontal = axis.position === "b" || axis.position === "t"
	const before = axes
		.slice(0, axes.indexOf(axis))
		.filter(other => (other.position === "b" || other.position === "t") === horizontal).length

	if (horizontal) return before === 0 ? "previewSheetChartAxisHorizontal" : "previewSheetChartAxisHorizontal2"

	return before === 0 ? "previewSheetChartAxisVertical" : "previewSheetChartAxisVertical2"
}

// A number field's text: "" is automatic (null); undefined when it is not a number.
export function numberField(text: string): number | null | undefined {
	const trimmed = text.trim()

	if (trimmed === "") return null

	const value = Number(trimmed)

	return Number.isFinite(value) ? value : undefined
}

export function numberText(value: number | null): string {
	return value === null ? "" : String(value)
}

// The selected range as a series' values ("B2:B9"; the chart qualifies it with its own sheet).
export function selectionReference(range: SheetRange | undefined): string | null {
	return range === undefined ? null : rangeName(range)
}

// A series' label in the panel: its name, else its number.
export function seriesLabel(name: string, index: number, numbered: (number: number) => string): string {
	return name.trim() === "" ? numbered(index + 1) : name
}
