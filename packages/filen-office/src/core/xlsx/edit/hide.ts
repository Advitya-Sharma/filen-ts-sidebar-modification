import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetAxis, type SheetId } from "../../sheet/protocol"
import { defineSheetFeature, type SheetChangeSink, type SheetHandled } from "../../sheet/registry"
import { defaultColWidth } from "../../sheet/sizes/sizes"
import type { SheetStore } from "../model/store"
import { ColFlag, RowFlag, type XlsxColRange } from "../model/types"
import type { XlsxDocument } from "./document"
import { NO_ROW, rowState, type RowState } from "./history"

// Rows and columns hidden and shown again (§18.3.1.73 row hidden, §18.3.1.13 col hidden). A row's own record
// carries the flag; a column's range is split where the flag changes, and a column with no range of its own gets
// one at the sheet's default width.

declare module "../../sheet/protocol" {
	interface SheetCommandMap {
		// Inclusive spans of rows (columns), hidden or shown.
		"axis.hide": {
			readonly sheet: SheetId
			readonly axis: SheetAxis
			readonly spans: readonly { readonly start: number; readonly end: number }[]
			readonly hidden: boolean
		}
	}
}

// Rows one command hides: each one becomes a row record of its own.
export const MAX_HIDDEN_ROWS = 65_536

// The columns `first`..`last` with HIDDEN set or cleared, ranges split where that changes and joined back with
// nothing: a gap gets a range of `width`.
export function withHidden(cols: readonly XlsxColRange[], first: number, last: number, hidden: boolean, width: number): XlsxColRange[] {
	const next: XlsxColRange[] = []
	const flag = (flags: number): number => (hidden ? flags | ColFlag.HIDDEN : flags & ~ColFlag.HIDDEN)
	let at = first

	for (const range of cols) {
		if (range.last < first || range.first > last) {
			if (range.first > last && at <= last) {
				if (hidden) next.push({ first: at, last, width, style: -1, flags: ColFlag.HIDDEN })

				at = last + 1
			}

			next.push(range)

			continue
		}

		if (range.first > at && hidden) next.push({ first: at, last: range.first - 1, width, style: -1, flags: ColFlag.HIDDEN })
		if (range.first < first) next.push({ ...range, last: first - 1 })

		// A column of no width reads as hidden too: shown, it takes the default width.
		const flat = !hidden && range.width === 0
		const inside = {
			...range,
			first: Math.max(range.first, first),
			last: Math.min(range.last, last),
			width: flat ? Number.NaN : range.width,
			flags: flat ? flag(range.flags) & ~ColFlag.CUSTOM_WIDTH : flag(range.flags)
		}
		const plain = Number.isNaN(inside.width) && inside.style < 0 && inside.flags === 0

		if (!plain) next.push(inside)
		if (range.last > last) next.push({ ...range, first: last + 1 })

		at = Math.min(range.last, last) + 1
	}

	if (at <= last && hidden) next.push({ first: at, last, width, style: -1, flags: ColFlag.HIDDEN })

	return next
}

function sameCols(a: readonly XlsxColRange[], b: readonly XlsxColRange[]): boolean {
	return (
		a.length === b.length &&
		a.every((range, k) => {
			const other = b[k] as XlsxColRange

			return (
				range.first === other.first &&
				range.last === other.last &&
				range.style === other.style &&
				range.flags === other.flags &&
				(Number.isNaN(range.width) ? Number.isNaN(other.width) : range.width === other.width)
			)
		})
	)
}

export async function hideAxis(
	document: XlsxDocument,
	sheet: SheetId,
	axis: SheetAxis,
	spans: readonly { readonly start: number; readonly end: number }[],
	hidden: boolean,
	changes: SheetChangeSink
): Promise<SheetHandled> {
	const limit = axis === "rows" ? SHEET_MAX_ROWS : SHEET_MAX_COLS

	if (
		spans.some(
			span =>
				!Number.isInteger(span.start) || !Number.isInteger(span.end) || span.start < 0 || span.end < span.start || span.end >= limit
		)
	)
		return { refused: { reason: "tooLarge" } }

	const protection = document.protects(sheet, axis === "rows" ? "formatRows" : "formatColumns")

	if (protection !== null) return { refused: protection }

	const model = await document.ready(sheet)

	if (axis === "cols") {
		const before = model.cols
		const format = model.format
		const width = defaultColWidth(format)
		let cols: readonly XlsxColRange[] = before

		for (const span of spans) cols = withHidden(cols, span.start, span.end, hidden, width)

		if (sameCols(cols, before)) return null

		const after = cols
		const set = (value: readonly XlsxColRange[], sink: SheetChangeSink): void => {
			model.cols = value
			document.workbook.changes.sheetSection(sheet, "cols")
			sink.layout(sheet)
		}

		set(after, changes)

		return { bytes: (before.length + after.length) * 64, undo: sink => set(before, sink), redo: sink => set(after, sink) }
	}

	const store = model.store as SheetStore
	const rows: number[] = []
	const befores: RowState[] = []
	const afters: RowState[] = []
	const visit = (row: number): void => {
		const state = rowState(store, row)
		// A row of no height reads as hidden too: shown, it takes the default height (as Excel unhides it).
		const flat = !hidden && state.height === 0
		const flags = hidden ? state.flags | RowFlag.HIDDEN : state.flags & ~(RowFlag.HIDDEN | (flat ? RowFlag.CUSTOM_HEIGHT : 0))

		if (flags === state.flags && !flat) return

		rows.push(row)
		befores.push(state)
		afters.push({ ...state, flags, height: flat ? Number.NaN : state.height })
	}

	if (hidden) {
		let count = 0

		for (const span of spans) count += span.end - span.start + 1

		if (count > MAX_HIDDEN_ROWS) return { refused: { reason: "tooLarge", detail: "rows" } }

		for (const span of spans) for (let row = span.start; row <= span.end; row++) visit(row)
	} else {
		// Only rows with a record can be hidden.
		for (const block of store.blocks) {
			for (let i = 0; i < block.rowCount; i++) {
				const row = block.firstRow + (block.rows[i] ?? 0)

				if (spans.some(span => row >= span.start && row <= span.end)) visit(row)
			}
		}
	}

	if (rows.length === 0) return null

	const set = (states: readonly RowState[], sink: SheetChangeSink): void => {
		store.setRows(
			rows.map((row, index) => {
				const state = states[index] ?? NO_ROW

				return { row, height: state.height, style: state.style, flags: state.flags }
			})
		)
		document.workbook.changes.sheetRows(sheet)
		sink.layout(sheet)
	}

	set(afters, changes)

	return { bytes: rows.length * 48, undo: sink => set(befores, sink), redo: sink => set(afters, sink) }
}

export const hideFeature = defineSheetFeature<XlsxDocument>({
	id: "hide",
	commands: {
		"axis.hide": {
			info: { write: true },
			run: (context, command) => hideAxis(context.doc, command.sheet, command.axis, command.spans, command.hidden, context.changes)
		}
	}
})
