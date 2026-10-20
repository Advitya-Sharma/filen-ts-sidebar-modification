import type { SheetAxis, SheetId, SheetQueryResult } from "../../sheet/protocol"
import type { SheetChangeSink, SheetHandled } from "../../sheet/registry"
import { MAX_RESIZE_TARGETS, clampSize, pxToColWidth, pxToRowHeight } from "../../sheet/sizes/sizes"
import type { SheetStore } from "../model/store"
import { ColFlag, RowFlag, type XlsxColRange } from "../model/types"
import type { XlsxDocument } from "./document"
import { NO_ROW, rowState, type RowState } from "./history"

// Column widths and row heights (axis.resize), kept in the file: a row's ht with customHeight, a column's <col>
// width with customWidth, split out of the range it shared. And what autofit measures (fitSample).

// Sets one column's width (NaN: back to the default) in a column list, splitting the range that held it. A column
// left with nothing of its own goes.
function withWidth(cols: readonly XlsxColRange[], col: number, width: number): XlsxColRange[] {
	const next: XlsxColRange[] = []
	let placed = false

	for (const range of cols) {
		if (col < range.first || col > range.last) {
			if (!placed && range.first > col) {
				if (!Number.isNaN(width)) next.push({ first: col, last: col, width, style: -1, flags: ColFlag.CUSTOM_WIDTH })

				placed = true
			}

			next.push(range)

			continue
		}

		placed = true

		if (range.first < col) next.push({ ...range, last: col - 1 })

		const flags = Number.isNaN(width) ? range.flags & ~ColFlag.CUSTOM_WIDTH : range.flags | ColFlag.CUSTOM_WIDTH
		const own = { ...range, first: col, last: col, width, flags }

		if (!Number.isNaN(width) || own.style >= 0 || (own.flags & ~ColFlag.BEST_FIT) !== 0) next.push(own)
		if (range.last > col) next.push({ ...range, first: col + 1 })
	}

	if (!placed && !Number.isNaN(width)) next.push({ first: col, last: col, width, style: -1, flags: ColFlag.CUSTOM_WIDTH })

	return next
}

function widthOf(cols: readonly XlsxColRange[], col: number): number {
	for (const range of cols)
		if (col >= range.first && col <= range.last) return (range.flags & ColFlag.CUSTOM_WIDTH) === 0 ? Number.NaN : range.width

	return Number.NaN
}

function sameHeight(a: number, b: number): boolean {
	return (Number.isNaN(a) && Number.isNaN(b)) || a === b
}

export async function resizeAxis(
	document: XlsxDocument,
	sheet: SheetId,
	axis: SheetAxis,
	indices: Uint32Array,
	sizes: Float32Array,
	changes: SheetChangeSink
): Promise<SheetHandled> {
	const model = await document.ready(sheet)

	if (indices.length !== sizes.length || indices.length > MAX_RESIZE_TARGETS) return { refused: { reason: "tooLarge" } }

	if (axis === "cols") {
		const before = model.cols
		let cols: readonly XlsxColRange[] = before

		for (let i = 0; i < indices.length; i++) {
			const col = indices[i] ?? 0
			const px = sizes[i] ?? Number.NaN
			const width = Number.isNaN(px) ? Number.NaN : pxToColWidth(clampSize("cols", px))

			if (!sameHeight(widthOf(cols, col), width)) cols = withWidth(cols, col, width)
		}

		if (cols === before) return null

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

	for (let i = 0; i < indices.length; i++) {
		const row = indices[i] ?? 0
		const px = sizes[i] ?? Number.NaN
		const state = rowState(store, row)
		const height = Number.isNaN(px) ? Number.NaN : pxToRowHeight(clampSize("rows", px))
		const flags = Number.isNaN(height) ? state.flags & ~RowFlag.CUSTOM_HEIGHT : state.flags | RowFlag.CUSTOM_HEIGHT

		if (sameHeight(state.height, height) && flags === state.flags) continue
		if (state === NO_ROW && Number.isNaN(height)) continue

		rows.push(row)
		befores.push(state)
		afters.push({ height, style: state.style, flags })
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

// The `limit` longest texts of a column (or row), with their looks and whether they wrap: what autofit measures.
export async function fitSample(
	document: XlsxDocument,
	sheet: SheetId,
	axis: SheetAxis,
	index: number,
	limit: number
): Promise<SheetQueryResult<"fitSample">> {
	const model = await document.ready(sheet)
	const store = model.store as SheetStore
	const kept: { text: string; style: number; wrap: boolean; index: number }[] = []
	const range =
		axis === "cols"
			? { startRow: 0, startCol: index, endRow: store.rowExtent, endCol: index }
			: { startRow: index, startCol: 0, endRow: index, endCol: store.colExtent }
	const cap = Math.max(0, Math.min(limit, 1024))

	store.forEachCell(range, (row, col) => {
		if (cap === 0) return

		const view = document.cellView(sheet, row, col)

		if (view.text === "") return
		if (kept.length === cap && (kept[cap - 1]?.text.length ?? 0) >= view.text.length) return

		const entry = {
			text: view.text,
			style: view.style,
			wrap: document.styles.renderStyle(view.xf).wrap === true,
			index: axis === "cols" ? row : col
		}
		let at = kept.length

		while (at > 0 && (kept[at - 1]?.text.length ?? 0) < view.text.length) at--

		kept.splice(at, 0, entry)

		if (kept.length > cap) kept.pop()
	})

	// Where each text is: its row in a column's sample, its column in a row's (a wrapped text fits its column).
	const sample = {
		texts: kept.map(entry => entry.text),
		styles: Uint32Array.from(kept, entry => entry.style),
		wraps: Uint8Array.from(kept, entry => (entry.wrap ? 1 : 0)),
		indices: Uint32Array.from(kept, entry => entry.index)
	}

	return sample
}
