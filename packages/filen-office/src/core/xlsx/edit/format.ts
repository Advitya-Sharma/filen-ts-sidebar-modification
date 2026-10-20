import { SHEET_MAX_EDIT_CELLS, type SheetColorInput, type SheetFormatPatch, type SheetId, type SheetRange } from "../../sheet/protocol"
import type { SheetStore } from "../model/store"
import { CellKind, XfApply, type XlsxAlignment, type XlsxColor, type XlsxFont, type XlsxXf } from "../model/types"
import { XlsxReadError } from "../read/index"
import type { XlsxDocument } from "./document"
import { CellLog } from "./history"
import type { SheetChangeSink, SheetEditStep, SheetHandled } from "../../sheet/registry"

// Cell formats an edit asks for, as appended records: the format a cell had, patched, becomes a new cell format
// (font, fill, alignment, number format and protection as asked, the rest as they were). Records are never
// changed in place, and each distinct format is patched once per edit.

// Excel's own limit on the cell formats one edit may append is the workbook's (styleLimit).
export class StyleLimitError extends Error {
	constructor() {
		super("The workbook has as many cell formats as Excel allows")
		this.name = "StyleLimitError"
	}
}

export function xlsxColor(color: SheetColorInput): XlsxColor {
	if (typeof color === "number") return { argb: (0xff000000 | (color & 0xffffff)) >>> 0 }

	return color.tint === undefined || color.tint === 0 ? { theme: color.theme } : { theme: color.theme, tint: color.tint }
}

function without<T extends object, K extends keyof T>(value: T, key: K): T {
	const copy = { ...value }

	delete copy[key]

	return copy
}

function patchedFont(font: XlsxFont, patch: SheetFormatPatch): XlsxFont {
	let next: XlsxFont = { ...font }

	// A cell's font is whole: off is absent (a written false would read as an override of nothing).
	if (patch.bold !== undefined) next = patch.bold ? { ...next, bold: true } : without(next, "bold")
	if (patch.italic !== undefined) next = patch.italic ? { ...next, italic: true } : without(next, "italic")
	if (patch.strike !== undefined) next = patch.strike ? { ...next, strike: true } : without(next, "strike")
	if (patch.underline !== undefined)
		next = patch.underline === null ? without(next, "underline") : { ...next, underline: patch.underline }
	if (patch.script !== undefined) next = patch.script === null ? without(next, "vertAlign") : { ...next, vertAlign: patch.script }
	if (patch.font !== undefined) next = patch.font === null ? without(next, "name") : { ...without(next, "scheme"), name: patch.font }
	if (patch.size !== undefined) next = patch.size === null ? without(next, "size") : { ...next, size: patch.size }
	if (patch.color !== undefined) next = patch.color === null ? without(next, "color") : { ...next, color: xlsxColor(patch.color) }

	return next
}

function patchedAlignment(alignment: XlsxAlignment | undefined, patch: SheetFormatPatch): XlsxAlignment | undefined {
	let next: XlsxAlignment = { ...alignment }

	if (patch.align !== undefined) next = patch.align === null ? without(next, "horizontal") : { ...next, horizontal: patch.align }
	if (patch.valign !== undefined) next = patch.valign === null ? without(next, "vertical") : { ...next, vertical: patch.valign }
	if (patch.wrap !== undefined) next = patch.wrap ? { ...next, wrapText: true } : without(next, "wrapText")
	if (patch.shrink !== undefined) next = patch.shrink ? { ...next, shrinkToFit: true } : without(next, "shrinkToFit")
	if (patch.indent !== undefined) next = patch.indent > 0 ? { ...next, indent: patch.indent } : without(next, "indent")
	if (patch.rotation !== undefined)
		next = patch.rotation !== 0 ? { ...next, textRotation: patch.rotation } : without(next, "textRotation")

	return Object.keys(next).length === 0 ? undefined : next
}

const FONT_KEYS: readonly (keyof SheetFormatPatch)[] = ["bold", "italic", "strike", "underline", "script", "font", "size", "color"]
const ALIGNMENT_KEYS: readonly (keyof SheetFormatPatch)[] = ["align", "valign", "wrap", "shrink", "indent", "rotation"]

function touches(patch: SheetFormatPatch, keys: readonly (keyof SheetFormatPatch)[]): boolean {
	return keys.some(key => patch[key] !== undefined)
}

// Patches cell formats of one workbook, each distinct one once. Throws StyleLimitError past Excel's limit.
export class XfPatcher {
	private readonly document: XlsxDocument
	private readonly patch: SheetFormatPatch
	private readonly done = new Map<number, number>()

	constructor(document: XlsxDocument, patch: SheetFormatPatch) {
		this.document = document
		this.patch = patch
	}

	xf(index: number): number {
		let next = this.done.get(index)

		if (next === undefined) {
			next = patchXf(this.document, index, this.patch)
			this.done.set(index, next)
		}

		return next
	}
}

function append(document: XlsxDocument, xf: XlsxXf): number {
	try {
		const index = document.workbook.styles.appendXf(xf)

		document.workbook.changes.stylesAppended()

		return index
	} catch (error) {
		if (error instanceof XlsxReadError && error.code === "styleLimit") throw new StyleLimitError()

		throw error
	}
}

export function patchXf(document: XlsxDocument, index: number, patch: SheetFormatPatch): number {
	const styles = document.workbook.styles
	const xf: XlsxXf = styles.cellXfs[index] ?? styles.cellXfs[0] ?? { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, flags: 0 }
	let next: XlsxXf = { ...xf }
	let flags = xf.flags

	if (touches(patch, FONT_KEYS)) {
		const font = styles.fonts[xf.fontId] ?? styles.fonts[0] ?? {}

		next = { ...next, fontId: styles.appendFont(patchedFont(font, patch)) }
		flags |= XfApply.FONT
	}

	if (patch.fill !== undefined) {
		next = {
			...next,
			fillId:
				patch.fill === null
					? 0
					: styles.appendFill({ kind: "pattern", pattern: "solid", fg: xlsxColor(patch.fill), bg: { indexed: 64 } })
		}
		flags |= XfApply.FILL
	}

	if (patch.numFmt !== undefined) {
		next = { ...next, numFmtId: patch.numFmt === "General" ? 0 : styles.numFmtId(patch.numFmt) }
		flags |= XfApply.NUMBER_FORMAT
	}

	if (touches(patch, ALIGNMENT_KEYS)) {
		const alignment = patchedAlignment(document.styles.alignment(index), patch)

		next = alignment === undefined ? without(next, "alignment") : { ...next, alignment }
		flags |= XfApply.ALIGNMENT
	}

	if (patch.locked !== undefined || patch.formulaHidden !== undefined) {
		const protection = { ...document.styles.protection(index) }

		if (patch.locked !== undefined) protection.locked = patch.locked
		if (patch.formulaHidden !== undefined) protection.hidden = patch.formulaHidden

		next = { ...next, protection }
		flags |= XfApply.PROTECTION
	}

	next = { ...next, flags }

	return append(document, next)
}

// The format with the quote prefix (§18.8.45: an entry typed with a leading apostrophe) set or cleared, and the
// number format `numFmt` when the format has none of its own (General): what typing a value implies.
export function typedXf(document: XlsxDocument, index: number, quoted: boolean, numFmt: string | undefined): number {
	const styles = document.workbook.styles
	const xf = styles.cellXfs[index]
	const general = (xf?.numFmtId ?? 0) === 0
	const prefixed = (xf?.flags ?? 0) & XfApply.QUOTE_PREFIX

	if ((prefixed !== 0) === quoted && (numFmt === undefined || !general)) return index

	const base: XlsxXf = xf ?? styles.cellXfs[0] ?? { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, flags: 0 }
	let flags = quoted ? base.flags | XfApply.QUOTE_PREFIX : base.flags & ~XfApply.QUOTE_PREFIX
	let numFmtId = base.numFmtId

	if (numFmt !== undefined && general) {
		numFmtId = styles.numFmtId(numFmt)
		flags |= XfApply.NUMBER_FORMAT
	}

	return append(document, { ...base, numFmtId, flags })
}

// Clips a range to what one format edit takes: the whole range when the limits allow, else (whole rows or
// columns selected) as far as the sheet is used. Null when even that is too large.
export function clippedRange(store: SheetStore, range: SheetRange, budget: number): SheetRange | null {
	const area = (r: SheetRange): number => (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1)

	if (area(range) <= budget) return range

	const clipped = {
		startRow: range.startRow,
		startCol: range.startCol,
		endRow: Math.min(range.endRow, Math.max(store.rowExtent - 1, range.startRow)),
		endCol: Math.min(range.endCol, Math.max(store.colExtent - 1, range.startCol))
	}

	return area(clipped) <= budget ? clipped : null
}

export function cellStep(document: XlsxDocument, log: CellLog, report: (changes: SheetChangeSink) => void, engine: boolean): SheetEditStep {
	const replay = (which: "before" | "after", changes: SheetChangeSink): void => {
		log.restore(document.restoreCell, which)

		if (engine) syncEngine(document, log)

		report(changes)
	}

	return {
		bytes: log.bytes,
		undo: changes => replay("before", changes),
		redo: changes => replay("after", changes)
	}
}

// After cells were restored: the engine gets the edited cells as they now stand; their dependents were restored
// from the log (the engine's own values of them already match what it calculated before).
export function syncEngine(document: XlsxDocument, log: CellLog): void {
	const engine = document.engine

	if (engine === null) return

	const bySheet = new Map<SheetId, { row: number; col: number }[]>()

	for (let i = 0; i < log.size; i++) {
		if (log.direct[i] !== true) continue

		const sheet = log.sheets[i] ?? 0
		let cells = bySheet.get(sheet)

		if (cells === undefined) {
			cells = []
			bySheet.set(sheet, cells)
		}

		cells.push({ row: log.rows[i] ?? 0, col: log.cols[i] ?? 0 })
	}

	for (const [sheet, cells] of bySheet) {
		const model = document.model(sheet)
		const index = document.engineIndex(sheet)

		if (model === null || index < 0) continue

		document.guarded(() =>
			engine.set(
				index,
				cells.map(cell => document.engineCell(model, cell.row, cell.col))
			)
		)
	}
}

// Every logged cell reported, as one range per sheet.
export function reportLogged(log: CellLog, changes: SheetChangeSink, kind: "content" | "look"): void {
	const bounds = new Map<SheetId, { startRow: number; startCol: number; endRow: number; endCol: number }>()

	for (let i = 0; i < log.size; i++) {
		const sheet = log.sheets[i] ?? 0
		const row = log.rows[i] ?? 0
		const col = log.cols[i] ?? 0
		const range = bounds.get(sheet)

		if (range === undefined) {
			bounds.set(sheet, { startRow: row, startCol: col, endRow: row, endCol: col })
		} else {
			range.startRow = Math.min(range.startRow, row)
			range.startCol = Math.min(range.startCol, col)
			range.endRow = Math.max(range.endRow, row)
			range.endCol = Math.max(range.endCol, col)
		}
	}

	for (const [sheet, range] of bounds) changes[kind](sheet, range)
}

// format.apply: each range's cells take their format patched; cells without one of their own start from their
// row's or column's.
export function applyFormat(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	patch: SheetFormatPatch,
	changes: SheetChangeSink
): SheetHandled {
	const model = document.model(sheet)

	if (model === null) return { refused: { reason: "tooLarge" } }

	const store = model.store as SheetStore
	const clipped: SheetRange[] = []
	let budget = SHEET_MAX_EDIT_CELLS

	for (const range of ranges) {
		const fitted = clippedRange(store, range, budget)

		if (fitted === null) return { refused: { reason: "tooLarge" } }

		clipped.push(fitted)
		budget -= (fitted.endRow - fitted.startRow + 1) * (fitted.endCol - fitted.startCol + 1)
	}

	const patcher = new XfPatcher(document, patch)
	const log = new CellLog()
	const targets: { row: number; col: number; style: number }[] = []

	try {
		for (const range of clipped) {
			for (let row = range.startRow; row <= range.endRow; row++) {
				for (let col = range.startCol; col <= range.endCol; col++) {
					const slot = store.locate(row, col)
					const xf = slot < 0 ? document.startingStyle(model, row, col) : store.style(slot)
					const next = patcher.xf(xf)

					if (slot >= 0 && next === xf) continue

					targets.push({ row, col, style: next })
				}
			}
		}
	} catch (error) {
		if (error instanceof StyleLimitError) return { refused: { reason: "styleLimit" } }

		throw error
	}

	if (targets.length === 0) return null

	for (const { row, col, style } of targets) {
		const state = document.state(sheet, row, col)

		log.touch(sheet, row, col, () => state, true)
		document.write(sheet, row, col, state.present ? { ...state, style } : { ...state, present: true, kind: CellKind.EMPTY, style })
	}

	log.seal((s, r, c) => document.state(s, r, c))

	const report = (sink: SheetChangeSink): void => {
		for (const range of clipped) sink.look(sheet, range)
	}

	report(changes)

	return cellStep(document, log, report, false)
}
