import { SHEET_MAX_COLS, type SheetId } from "../../sheet/protocol"
import type { SheetChangeSink, SheetHandled } from "../../sheet/registry"
import { runSliced } from "../../slicing"
import { MAX_RESIZE_TARGETS, PX_PER_POINT, clampSize, pxToRowHeight, rowHeightToPx } from "../../sheet/sizes/sizes"
import type { SheetStore } from "../model/store"
import { RowFlag, type XlsxSheetModel } from "../model/types"
import type { XlsxDocument } from "./document"
import { NO_ROW, rowState, type RowState } from "./history"

// A row without a height of its own (no customHeight, ECMA-376 Part 1 §18.3.1.73) fits its content, as Excel grows
// and shrinks such a row when a font size or wrapped text changes; the file then keeps the fitted height as ht
// without customHeight, the value Excel writes for it (ISO/IEC 29500 asks only for the resulting ht). Only the
// rows an edit touched are fitted. The core measures no text: the grid measures wrapped text it shows and sends
// the rows' wrapped heights (axis.fit); a row with wrapped text the grid did not measure is never shrunk.

// A line of text in pixels at 100 %, and the row's padding beyond it: what the grid's autofit measures a line as.
const LINE_HEIGHT = 1.2
const ROW_PADDING = 2

// Rows fitted between the host's turns.
const ROWS_PER_STEP = 256

function linePx(points: number): number {
	return Math.round(points * PX_PER_POINT * LINE_HEIGHT) + ROW_PADDING
}

interface XfLine {
	readonly points: number
	readonly wraps: boolean
}

// Wrapped text the grid measured: each row's tallest wrapped cell in px at 100 % (NaN: none) over the columns it
// looked at.
export interface MeasuredWraps {
	readonly heights: ReadonlyMap<number, number>
	readonly firstCol: number
	readonly lastCol: number
}

// The rows' new properties where their fitted height differs from the one they have.
export function* fittedRows(
	document: XlsxDocument,
	model: XlsxSheetModel,
	rows: Iterable<number>,
	measured?: MeasuredWraps
): Generator<void, { readonly row: number; readonly state: RowState }[]> {
	const store = model.store as SheetStore
	const styles = document.workbook.styles
	const normalFont = styles.fonts[styles.cellStyleXfs[0]?.fontId ?? 0] ?? styles.fonts[0]
	const normal = normalFont?.size ?? 11
	const defaultPx = rowHeightToPx(model.format.defaultRowHeight)
	const lines = new Map<number, XfLine>()
	const lineOf = (xf: number): XfLine => {
		let line = lines.get(xf)

		if (line === undefined) {
			const record = styles.cellXfs[xf]

			line = {
				points: styles.fonts[record?.fontId ?? 0]?.size ?? normal,
				wraps: document.styles.alignment(xf)?.wrapText === true
			}
			lines.set(xf, line)
		}

		return line
	}
	const out: { row: number; state: RowState }[] = []
	let done = 0

	for (const row of rows) {
		const state = rowState(store, row)

		if ((state.flags & RowFlag.CUSTOM_HEIGHT) === 0) {
			const wrapped = measured?.heights.get(row)
			let points = normal
			// Wrapped text no measure covers.
			let unmeasured = false
			const take = (xf: number, col: number): void => {
				const line = lineOf(xf)

				points = Math.max(points, line.points)

				if (line.wraps && (wrapped === undefined || col < (measured?.firstCol ?? 0) || col > (measured?.lastCol ?? -1)))
					unmeasured = true
			}

			if ((state.flags & RowFlag.CUSTOM_FORMAT) !== 0 && state.style >= 0) take(state.style, -1)

			store.forEachCell({ startRow: row, startCol: 0, endRow: row, endCol: SHEET_MAX_COLS - 1 }, (_row, col, slot) => {
				take(store.style(slot), col)
			})

			// The default row holds the normal font; a taller font adds what its line adds.
			let px = defaultPx + Math.max(0, linePx(points) - linePx(normal))
			const current = Number.isNaN(state.height) ? defaultPx : rowHeightToPx(state.height)

			if (wrapped !== undefined && !Number.isNaN(wrapped)) px = Math.max(px, Math.round(wrapped))
			if (unmeasured) px = Math.max(px, current)

			px = clampSize("rows", px)

			const height = px === defaultPx ? Number.NaN : pxToRowHeight(px)

			if (px !== current) {
				out.push({ row, state: { height, style: state === NO_ROW ? -1 : state.style, flags: state === NO_ROW ? 0 : state.flags } })
			}
		}

		if (++done % ROWS_PER_STEP === 0) yield
	}

	return out
}

// The rows each fitted: the grid's measured wrapped heights (axis.fit), as one step.
export async function fitRows(
	document: XlsxDocument,
	sheet: SheetId,
	measured: MeasuredWraps,
	changes: SheetChangeSink
): Promise<SheetHandled> {
	if (measured.heights.size > MAX_RESIZE_TARGETS) return { refused: { reason: "tooLarge" } }

	const model = await document.ready(sheet)
	const store = model.store as SheetStore
	const fitted = runSliced(
		fittedRows(
			document,
			model,
			[...measured.heights.keys()].sort((a, b) => a - b),
			measured
		),
		undefined
	)
	const after = await fitted

	if (after.length === 0) return null

	const before = after.map(({ row }) => ({ row, state: rowState(store, row) }))
	const set = (states: readonly { readonly row: number; readonly state: RowState }[], sink: SheetChangeSink): void => {
		store.setRows(states.map(({ row, state }) => ({ row, height: state.height, style: state.style, flags: state.flags })))
		document.workbook.changes.sheetRows(sheet)
		sink.layout(sheet)
	}

	set(after, changes)

	return { bytes: after.length * 96, undo: sink => set(before, sink), redo: sink => set(after, sink) }
}
