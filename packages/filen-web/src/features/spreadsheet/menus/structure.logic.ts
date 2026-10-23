import {
	AxisFlag,
	SHEET_MAX_COLS,
	SHEET_MAX_ROWS,
	type SheetAxis,
	type SheetAxisLayout,
	type SheetLayout,
	type SheetPane,
	type SheetRange
} from "@filen/office/core/sheet"

export interface AxisSpan {
	readonly start: number
	readonly end: number
}

// The rows (columns) a selection spans, each range once, overlaps joined.
export function selectionSpans(ranges: readonly SheetRange[], axis: SheetAxis): AxisSpan[] {
	const spans = ranges
		.map(range => (axis === "rows" ? { start: range.startRow, end: range.endRow } : { start: range.startCol, end: range.endCol }))
		.sort((a, b) => a.start - b.start)
	const joined: AxisSpan[] = []

	for (const span of spans) {
		const last = joined.at(-1)

		if (last !== undefined && span.start <= last.end + 1)
			joined[joined.length - 1] = { start: last.start, end: Math.max(last.end, span.end) }
		else joined.push(span)
	}

	return joined
}

export function spanCount(spans: readonly AxisSpan[]): number {
	return spans.reduce((count, span) => count + span.end - span.start + 1, 0)
}

// Whether an axis hides any index in the spans (a row or column of no size reads as hidden too). An entry may stand
// for a run of indices (SheetAxisLayout.counts).
export function hidesWithin(layout: SheetAxisLayout, spans: readonly AxisSpan[]): boolean {
	for (let k = 0; k < layout.indices.length; k++) {
		if (((layout.flags[k] ?? 0) & AxisFlag.HIDDEN) === 0 && layout.sizes[k] !== 0) continue

		const first = layout.indices[k] ?? 0
		const last = first + (layout.counts?.[k] ?? 1) - 1

		if (spans.some(span => first <= span.end && last >= span.start)) return true
	}

	return false
}

// What "Unhide" shows again, as Excel does: the hidden rows (columns) inside the selection's spans, and those next to
// them (a selection of the rows either side of a hidden one), within the sheet. Null when there are none.
export function unhideSpans(spans: readonly AxisSpan[], layout: SheetLayout | null, axis: SheetAxis): AxisSpan[] | null {
	if (layout === null) return null

	const last = (axis === "rows" ? SHEET_MAX_ROWS : SHEET_MAX_COLS) - 1
	const widened = spans.map(span => ({ start: Math.max(0, span.start - 1), end: Math.min(last, span.end + 1) }))

	if (!hidesWithin(axis === "rows" ? layout.rows : layout.cols, widened)) return null

	return widened
}

// The panes the freeze items set: frozen rows and columns.
export type FreezeTarget = "panes" | "topRow" | "firstColumn" | "none"

export function freezePane(target: FreezeTarget, active: { readonly row: number; readonly col: number }): { rows: number; cols: number } {
	switch (target) {
		case "panes":
			return { rows: active.row, cols: active.col }
		case "topRow":
			return { rows: 1, cols: 0 }
		case "firstColumn":
			return { rows: 0, cols: 1 }
		case "none":
			return { rows: 0, cols: 0 }
	}
}

// Freezing at A1 freezes nothing: Excel then splits at the centre; here the item is off.
export function canFreezePanes(active: { readonly row: number; readonly col: number }): boolean {
	return active.row > 0 || active.col > 0
}

export function samePane(a: SheetPane | null, b: SheetPane | null): boolean {
	if (a === null || b === null) return a === b

	return (
		a.rows === b.rows &&
		a.cols === b.cols &&
		a.firstRow === b.firstRow &&
		a.firstCol === b.firstCol &&
		a.topRow === b.topRow &&
		a.leftCol === b.leftCol
	)
}
