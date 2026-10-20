import { SHEET_MAX_EDIT_CELLS, type SheetId, type SheetRange } from "../../sheet/protocol"
import { defineSheetFeature, type SheetChangeSink, type SheetHandled, type SheetReadContext, type SheetTarget } from "../../sheet/registry"
import type { SheetStore } from "../model/store"
import { CellKind, FormulaKind } from "../model/types"
import { recalculateLogged, splitsArray, unshare } from "./cells"
import type { XlsxDocument } from "./document"
import { StyleLimitError, XfPatcher, cellStep, reportLogged } from "./format"
import { CellLog, type CellState } from "./history"

// Merged cells (§18.3.1.55 mergeCell). As Excel merges them ("Merge and center", support.microsoft.com): the
// upper-left cell's value stays and the others' go (only once the user agreed: refused with mergeLosesValues
// first), every cell of the merge takes the upper-left one's format, and merges the range holds whole are taken in.
// "center" also centres it; "across" merges each row on its own. A merge that would cut through another one, a
// table or an array formula is refused.

declare module "../../sheet/protocol" {
	interface SheetCommandMap {
		"merge.cells": {
			readonly sheet: SheetId
			readonly ranges: readonly SheetRange[]
			readonly mode: "merge" | "center" | "across"
			// The user agreed that values outside the upper-left cells go.
			readonly discard?: boolean
		}
		"merge.unmerge": { readonly sheet: SheetId; readonly ranges: readonly SheetRange[] }
	}

	interface SheetRefusalReasons {
		// Cells other than the upper-left one hold values the merge would drop (the user is asked first).
		mergeLosesValues: true
	}
}

function intersects(a: SheetRange, b: SheetRange): boolean {
	return a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol
}

function contains(outer: SheetRange, inner: SheetRange): boolean {
	return (
		outer.startRow <= inner.startRow && outer.endRow >= inner.endRow && outer.startCol <= inner.startCol && outer.endCol >= inner.endCol
	)
}

function quads(merges: Uint32Array): SheetRange[] {
	const out: SheetRange[] = []

	for (let k = 0; k + 3 < merges.length; k += 4)
		out.push({ startRow: merges[k] ?? 0, startCol: merges[k + 1] ?? 0, endRow: merges[k + 2] ?? 0, endCol: merges[k + 3] ?? 0 })

	return out
}

function packed(ranges: readonly SheetRange[]): Uint32Array {
	const out = new Uint32Array(ranges.length * 4)

	ranges.forEach((range, k) => out.set([range.startRow, range.startCol, range.endRow, range.endCol], k * 4))

	return out
}

function area(range: SheetRange): number {
	return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)
}

// The merges a command makes: each range whole, or each of its rows ("across"); single cells are none.
function mergesOf(ranges: readonly SheetRange[], mode: "merge" | "center" | "across"): SheetRange[] {
	const out: SheetRange[] = []

	for (const range of ranges) {
		if (mode !== "across") {
			if (area(range) > 1) out.push(range)

			continue
		}

		if (range.endCol === range.startCol) continue

		for (let row = range.startRow; row <= range.endRow; row++) out.push({ ...range, startRow: row, endRow: row })
	}

	return out
}

// Why the merges cannot be made, or null.
async function mergeRefusal(
	document: XlsxDocument,
	sheet: SheetId,
	merges: readonly SheetRange[]
): Promise<{ reason: "merged" | "arrayFormula" | "protectedSheet" | "tooLarge"; detail?: string } | null> {
	const model = await document.ready(sheet)

	if (model.protection !== null) return { reason: "protectedSheet", detail: "merge" }

	let cells = 0

	for (const merge of merges) cells += area(merge)

	if (cells > SHEET_MAX_EDIT_CELLS) return { reason: "tooLarge" }

	for (let k = 0; k < merges.length; k++)
		for (let j = k + 1; j < merges.length; j++)
			if (intersects(merges[k] as SheetRange, merges[j] as SheetRange)) return { reason: "merged" }

	for (const existing of quads(model.merges))
		for (const merge of merges) if (intersects(existing, merge) && !contains(merge, existing)) return { reason: "merged" }

	for (const table of await document.sheetTables(sheet))
		for (const merge of merges) if (intersects(table.range, merge)) return { reason: "merged", detail: "table" }

	const store = model.store as SheetStore
	const rows: number[] = []
	const cols: number[] = []

	for (const merge of merges) {
		store.forEachCell(merge, (row, col, slot) => {
			const formula = store.formula(slot)

			if (formula >= 0) {
				rows.push(row)
				cols.push(col)

				const kind = store.formulas.kind(formula)

				if (kind === FormulaKind.ARRAY || kind === FormulaKind.DATA_TABLE) cells = -1
			}
		})
	}

	if (cells < 0 || splitsArray(model, rows, cols)) return { reason: "arrayFormula" }

	return null
}

function holds(state: CellState): boolean {
	return state.formula >= 0 || state.kind !== CellKind.EMPTY
}

export async function mergeCells(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	mode: "merge" | "center" | "across",
	discard: boolean,
	changes: SheetChangeSink
): Promise<SheetHandled> {
	const merges = mergesOf(ranges, mode)

	if (merges.length === 0) return null

	const refusal = await mergeRefusal(document, sheet, merges)

	if (refusal !== null) return { refused: refusal }

	const model = await document.ready(sheet)
	const store = model.store as SheetStore

	if (!discard) {
		for (const merge of merges) {
			let losing = false

			store.forEachCell(merge, (row, col) => {
				if ((row !== merge.startRow || col !== merge.startCol) && holds(document.state(sheet, row, col))) losing = true
			})

			if (losing) return { refused: { reason: "mergeLosesValues" } }
		}
	}

	// Built from the cells as they stand before the merge, so dropped values are what it recalculates.
	if (document.engine === null && document.allReady && document.hasFormulas) {
		await document.loadTables()
		document.buildEngine(changes)
	}

	const log = new CellLog()
	const centred = mode === "merge" ? null : new XfPatcher(document, { align: "center" })

	try {
		for (const merge of merges) {
			const anchor = document.state(sheet, merge.startRow, merge.startCol)
			const style = centred === null ? anchor.style : centred.xf(anchor.style)

			for (let row = merge.startRow; row <= merge.endRow; row++) {
				for (let col = merge.startCol; col <= merge.endCol; col++) {
					const before = document.state(sheet, row, col)
					const upperLeft = row === merge.startRow && col === merge.startCol
					let next: CellState

					if (upperLeft) next = { ...before, present: before.present || style !== 0, style }
					else
						next =
							style === 0 && !before.present
								? before
								: { present: style !== 0, kind: CellKind.EMPTY, value: 0, style, formula: -1, shown: undefined }

					if (
						next.present === before.present &&
						next.kind === before.kind &&
						next.style === before.style &&
						next.formula === before.formula
					)
						continue

					if (!upperLeft && before.formula >= 0 && store.formulas.kind(before.formula) === FormulaKind.SHARED_MASTER)
						unshare(document, sheet, row, col, log)

					log.touch(sheet, row, col, () => document.state(sheet, row, col), !upperLeft)
					document.write(sheet, row, col, next)
				}
			}
		}
	} catch (error) {
		if (error instanceof StyleLimitError) {
			log.restore(document.restoreCell, "before")

			return { refused: { reason: "styleLimit" } }
		}

		throw error
	}

	const before = model.merges
	const kept = quads(before).filter(existing => !merges.some(merge => contains(merge, existing)))
	const after = packed([...kept, ...merges])
	const setMerges = (value: Uint32Array, sink: SheetChangeSink): void => {
		model.merges = value
		document.workbook.changes.sheetSection(sheet, "mergeCells")
		sink.layout(sheet)
	}

	recalculateLogged(document, log, changes)
	log.seal((s, r, c) => document.state(s, r, c))
	reportLogged(log, changes, "content")
	setMerges(after, changes)

	const cells = cellStep(document, log, sink => reportLogged(log, sink, "content"), true)

	return {
		bytes: cells.bytes + (before.byteLength + after.byteLength),
		undo: async sink => {
			setMerges(before, sink)
			await cells.undo(sink)
		},
		redo: async sink => {
			await cells.redo(sink)
			setMerges(after, sink)
		}
	}
}

export async function unmergeCells(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	changes: SheetChangeSink
): Promise<SheetHandled> {
	const model = await document.ready(sheet)

	if (model.protection !== null) return { refused: { reason: "protectedSheet", detail: "merge" } }

	const before = model.merges
	const all = quads(before)
	const kept = all.filter(existing => !ranges.some(range => intersects(range, existing)))

	if (kept.length === all.length) return null

	const after = packed(kept)
	const set = (value: Uint32Array, sink: SheetChangeSink): void => {
		model.merges = value
		document.workbook.changes.sheetSection(sheet, "mergeCells")
		sink.layout(sheet)

		for (const range of all) sink.look(sheet, range)
	}

	set(after, changes)

	return { bytes: before.byteLength + after.byteLength, undo: sink => set(before, sink), redo: sink => set(after, sink) }
}

// Whether the selection has merges to undo (menus).
function unmergeCheck(context: SheetReadContext<XlsxDocument>, target: SheetTarget): { reason: "unsupported" } | null {
	// A sheet still being parsed answers from the preview of its first rows.
	const model = context.doc.model(target.sheet) ?? context.doc.previews.get(target.sheet)?.model ?? null

	if (model === null) return null

	return quads(model.merges).some(existing => target.ranges.some(range => intersects(range, existing))) ? null : { reason: "unsupported" }
}

export const mergeFeature = defineSheetFeature<XlsxDocument>({
	id: "merge",
	commands: {
		"merge.cells": {
			info: { write: true },
			run: (context, command) =>
				mergeCells(context.doc, command.sheet, command.ranges, command.mode, command.discard === true, context.changes)
		},
		"merge.unmerge": {
			info: { write: true },
			check: unmergeCheck,
			run: (context, command) => unmergeCells(context.doc, command.sheet, command.ranges, context.changes)
		}
	}
})
