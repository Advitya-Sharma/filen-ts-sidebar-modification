import { numFmtProblem, unwritableText } from "../../sheet/numfmt/edit"
import { SHEET_MAX_EDIT_CELLS, type SheetColorInput, type SheetFormatPatch, type SheetId, type SheetRange } from "../../sheet/protocol"
import { BLOCK_ROWS, type CellWrite, type SheetStore } from "../model/store"
import { CellKind, XfApply, type XlsxAlignment, type XlsxColor, type XlsxFont, type XlsxXf } from "../model/types"
import { XlsxReadError } from "../read/index"
import type { XlsxDocument } from "./document"
import { runSliced, sliceClock, type SliceClock, type Slicing } from "../../slicing"
import { AxisChange, partialAxis, restoreAxisChange, wholeAxis, wholeAxisTargets, type PositionedPatch } from "./formatAxes"
import { CellLog } from "./history"
import { fittedRows } from "./formatHeights"
import type { SheetChangeSink, SheetEditStep, SheetHandled } from "../../sheet/registry"

// Cell formats an edit asks for, as appended records: the format a cell had, patched, becomes a new cell format
// (font, fill, alignment, number format and protection as asked, the rest as they were). Records are never
// changed in place, and each distinct format is patched once per edit.

// Excel's own limit on the cell formats one edit may append is the workbook's (styleLimit).
const DEFAULT_XF: XlsxXf = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, flags: 0 }

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

const PARTS = [
	[XfApply.NUMBER_FORMAT, "numFmtId"],
	[XfApply.FONT, "fontId"],
	[XfApply.FILL, "fillId"],
	[XfApply.BORDER, "borderId"]
] as const

// An apply bit (§18.8.45) on a part equal to the cell style's tells nothing (applied or not, it is the same part):
// it goes, so a change back gives the format the cell started from rather than a look-alike.
function withoutIdleApply(document: XlsxDocument, xf: XlsxXf): XlsxXf {
	const parent = xf.xfId >= 0 ? document.workbook.styles.cellStyleXfs[xf.xfId] : undefined

	if (parent === undefined) return xf

	let flags = xf.flags

	for (const [bit, key] of PARTS) if ((flags & bit) !== 0 && xf[key] === parent[key]) flags &= ~bit

	return flags === xf.flags ? xf : { ...xf, flags }
}

// A cell format's index: an equal one the workbook holds, else one appended. `copy`: a record of its own even where
// the file holds an equal one. Throws StyleLimitError past the workbook's limit.
export function appendXf(document: XlsxDocument, xf: XlsxXf, copy = false): number {
	try {
		const index = document.workbook.styles.appendXf(copy ? xf : withoutIdleApply(document, xf), copy)

		document.workbook.changes.stylesAppended()

		return index
	} catch (error) {
		if (error instanceof XlsxReadError && error.code === "styleLimit") throw new StyleLimitError()

		throw error
	}
}

export function patchXf(document: XlsxDocument, index: number, patch: SheetFormatPatch): number {
	const styles = document.workbook.styles
	const xf: XlsxXf = styles.cellXfs[index] ?? styles.cellXfs[0] ?? DEFAULT_XF
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

	return appendXf(document, next)
}

// The format with the quote prefix (§18.8.45: an entry typed with a leading apostrophe) set or cleared, and the
// number format `numFmt` when the format has none of its own (General): what typing a value implies.
export function typedXf(document: XlsxDocument, index: number, quoted: boolean, numFmt: string | undefined): number {
	const styles = document.workbook.styles
	const xf = styles.cellXfs[index]
	const general = (xf?.numFmtId ?? 0) === 0
	const prefixed = (xf?.flags ?? 0) & XfApply.QUOTE_PREFIX

	if ((prefixed !== 0) === quoted && (numFmt === undefined || !general)) return index

	const base: XlsxXf = xf ?? styles.cellXfs[0] ?? DEFAULT_XF
	let flags = quoted ? base.flags | XfApply.QUOTE_PREFIX : base.flags & ~XfApply.QUOTE_PREFIX
	let numFmtId = base.numFmtId

	if (numFmt !== undefined && general) {
		numFmtId = styles.numFmtId(numFmt)
		flags |= XfApply.NUMBER_FORMAT
	}

	return appendXf(document, { ...base, numFmtId, flags })
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

// Cells the formula engine is given at a time in a sliced edit, its evaluation held between parts.
export const ENGINE_SLICE = 8192

// With `slicing` (the in-thread host), undo and redo write the cells back in slices.
export function cellStep(
	document: XlsxDocument,
	log: CellLog,
	report: (changes: SheetChangeSink) => void,
	engine: boolean,
	slicing?: Slicing
): SheetEditStep {
	const replay = async (which: "before" | "after", changes: SheetChangeSink): Promise<void> => {
		const clock = slicing === undefined ? undefined : sliceClock(slicing)

		await document.restoreLogSliced(log, which, clock)

		if (engine) await syncEngine(document, log, clock)

		report(changes)
	}

	return {
		bytes: log.bytes,
		undo: changes => replay("before", changes),
		redo: changes => replay("after", changes)
	}
}

// After cells were restored: the engine gets the edited cells as they now stand; their dependents were restored
// from the log (the engine's own values of them already match what it calculated before). With a clock, in parts of
// ENGINE_SLICE cells, evaluation held between them, yielding once the slice is up.
export async function syncEngine(document: XlsxDocument, log: CellLog, clock?: SliceClock): Promise<void> {
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

		if (clock === undefined) {
			document.guarded(() =>
				engine.set(
					index,
					cells.map(cell => document.engineCell(model, cell.row, cell.col))
				)
			)

			continue
		}

		// In parts, evaluation held until all are in: what reads them is recalculated once.
		try {
			engine.holdEvaluation()

			for (let start = 0; start < cells.length; start += ENGINE_SLICE) {
				engine.set(
					index,
					cells.slice(start, start + ENGINE_SLICE).map(cell => document.engineCell(model, cell.row, cell.col))
				)

				if (clock.due) await clock.pause()
			}

			await engine.releaseEvaluation(false, clock)
		} catch {
			document.dropEngine()

			return
		}
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

function areaOf(range: SheetRange): number {
	return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)
}

// Puts every logged cell of a format edit back as `which` side has it, through one batch (see planSteps): such an
// edit changed styles only, so the cells' contents, formulas and shown results are the same on both sides.
function restoreStyles(store: SheetStore, log: CellLog, which: "before" | "after"): void {
	const states = which === "before" ? log.before : log.after
	const cells: CellWrite[] = []

	for (let i = 0; i < log.size; i++) {
		const state = states[i]

		if (state === undefined) continue

		cells.push(
			state.present
				? {
						row: log.rows[i] ?? 0,
						col: log.cols[i] ?? 0,
						kind: state.kind,
						value: state.value,
						style: state.style,
						formula: state.formula,
						origin: state.origin ?? -1
					}
				: { row: log.rows[i] ?? 0, col: log.cols[i] ?? 0, kind: CellKind.EMPTY, value: 0, style: 0, formula: -1, origin: -1 }
		)
	}

	store.setCells(cells)
}

// Where a refused edit's partial changes are reported while they are taken back: nowhere (nothing was reported).
const UNSEEN: SheetChangeSink = {
	content: () => undefined,
	look: () => undefined,
	moved: () => undefined,
	layout: () => undefined,
	sheets: () => undefined,
	select: () => undefined
}

// One range of a format edit: its cells take `patch` (by where they are). `existing`: a neighbour whose facing
// border goes, never formatted through its axes; an empty cell gets a cell only where the format its row or column
// gives it changes.
export interface FormatPlanEntry {
	readonly range: SheetRange
	readonly patch: PositionedPatch
	readonly existing?: boolean
}

// Runs a format edit range by range, each seeing what the ones before it did: whole columns and rows through their
// axes (formatAxes.ts), other ranges cell by cell. Refused, with nothing changed, past the edit's cell budget or the
// workbook's style limit.
// `fitRows`: the edit may change font sizes, so the rows it touched fit their fonts again (formatHeights.ts).
export function formatPlan(
	document: XlsxDocument,
	sheet: SheetId,
	entries: readonly FormatPlanEntry[],
	changes: SheetChangeSink,
	slicing?: Slicing,
	fitRows = false
): Promise<SheetHandled> {
	return runSliced(planSteps(document, sheet, entries, changes, fitRows), slicing === undefined ? undefined : sliceClock(slicing))
}

// Cells written between the host's turns are counted in these.
const CELLS_PER_STEP = 64
// Cells written to the store at once: about a block's worth, each block rebuilt about once.
const CELLS_PER_BATCH = 1024

function* planSteps(
	document: XlsxDocument,
	sheet: SheetId,
	entries: readonly FormatPlanEntry[],
	changes: SheetChangeSink,
	fitRows: boolean
): Generator<void, SheetHandled> {
	const model = document.model(sheet)

	if (model === null) return { refused: { reason: "tooLarge" } }

	const store = model.store as SheetStore
	const log = new CellLog()
	const axes = new AxisChange(sheet)
	const looked: SheetRange[] = []
	let budget = SHEET_MAX_EDIT_CELLS
	let written = 0
	// A format changes cells' styles only (and adds or drops cells without content): written to the store in
	// batches, so a block is rebuilt once per batch rather than once per cell it gains.
	const pending: CellWrite[] = []
	const flush = (): void => {
		if (pending.length === 0) return

		store.setCells(pending)
		pending.length = 0
		document.workbook.changes.sheetRows(sheet)
	}
	const write = (row: number, col: number, style: number): boolean => {
		const state = document.state(sheet, row, col)

		log.touch(sheet, row, col, () => state, true)
		pending.push(
			state.present
				? { row, col, kind: state.kind, value: state.value, style, formula: state.formula, origin: -1 }
				: { row, col, kind: CellKind.EMPTY, value: 0, style, formula: -1, origin: -1 }
		)

		if (pending.length >= CELLS_PER_BATCH) {
			flush()

			return true
		}

		return ++written % CELLS_PER_STEP === 0
	}
	const undoAll = (): void => {
		pending.length = 0
		restoreStyles(store, log, "before")
		document.workbook.changes.sheetRows(sheet)
		restoreAxisChange(document, axes, "before", UNSEEN)
	}
	const refuse = (reason: "tooLarge" | "styleLimit"): SheetHandled => {
		undoAll()

		return { refused: { reason } }
	}

	try {
		for (const entry of entries) {
			let whole = entry.existing === true ? null : wholeAxis(entry.range)

			// A range reaching the grid's edge from inside it is formatted cell by cell while that fits.
			if (whole !== null && partialAxis(entry.range, whole) && areaOf(entry.range) <= budget) whole = null

			if (whole !== null) {
				const taken = yield* wholeAxisTargets(model, entry.range, whole, entry.patch, budget)

				if (taken === null) return refuse("tooLarge")

				for (const target of taken.cells) if (write(target.row, target.col, target.style)) yield

				if (taken.cols !== null) {
					axes.setCols(model, taken.cols)
					document.workbook.changes.sheetSection(sheet, "cols")
				}

				// A block of rows at a time: each is rebuilt once.
				for (let at = 0; at < taken.rows.length; at += BLOCK_ROWS) {
					axes.setRows(store, taken.rows.slice(at, at + BLOCK_ROWS))
					yield
				}

				if (taken.rows.length > 0) document.workbook.changes.sheetRows(sheet)

				budget -= taken.cells.length
				flush()
				looked.push(entry.range)
				continue
			}

			// A range is formatted whole or not at all. An `existing` one adds a cell only where the format its row or
			// column gives it changes.
			const range = entry.range

			if (entry.existing !== true) {
				if (areaOf(range) > budget) return refuse("tooLarge")

				budget -= areaOf(range)
			}

			for (let row = range.startRow; row <= range.endRow; row++) {
				for (let col = range.startCol; col <= range.endCol; col++) {
					const slot = store.locate(row, col)
					const xf = slot < 0 ? document.startingStyle(model, row, col) : store.style(slot)
					const next = entry.patch(xf, row, col, true)

					if (next === xf) continue
					if (slot < 0 && entry.existing === true && --budget < 0) return refuse("tooLarge")

					if (write(row, col, next)) yield
				}
			}

			flush()
			looked.push(range)
		}
	} catch (error) {
		if (error instanceof StyleLimitError) return refuse("styleLimit")

		throw error
	}

	if (log.size === 0 && !axes.changed) return null

	// A cell of the default format without content is no cell, so it would show its row's or column's format: where
	// that is not the default, it is kept as a copy of the default format (an equal record of its own).
	let copy = -1

	try {
		for (let i = 0; i < log.size; i++) {
			const row = log.rows[i] ?? 0
			const col = log.cols[i] ?? 0

			if (store.locate(row, col) < 0 && log.direct[i] === true && document.startingStyle(model, row, col) !== 0) {
				if (copy < 0) copy = appendXf(document, { ...(document.workbook.styles.cellXfs[0] ?? DEFAULT_XF) }, true)

				pending.push({ row, col, kind: CellKind.EMPTY, value: 0, style: copy, formula: -1, origin: -1 })

				if (pending.length >= CELLS_PER_BATCH) flush()
			}

			if ((i + 1) % CELLS_PER_STEP === 0) yield
		}
	} catch (error) {
		if (error instanceof StyleLimitError) return refuse("styleLimit")

		throw error
	}

	flush()

	if (fitRows) {
		const rows = new Set<number>(log.rows)

		for (const row of axes.rows.keys()) rows.add(row)

		const fitted = yield* fittedRows(
			document,
			model,
			[...rows].sort((a, b) => a - b)
		)

		for (let at = 0; at < fitted.length; at += BLOCK_ROWS) {
			axes.setRows(store, fitted.slice(at, at + BLOCK_ROWS))
			yield
		}

		if (fitted.length > 0) document.workbook.changes.sheetRows(sheet)
	}

	// What each touched cell now holds (CellLog.seal), in steps.
	log.after.length = 0

	for (let i = 0; i < log.size; i++) {
		log.after.push(document.state(log.sheets[i] ?? sheet, log.rows[i] ?? 0, log.cols[i] ?? 0))

		if ((i + 1) % CELLS_PER_STEP === 0) yield
	}

	const report = (sink: SheetChangeSink): void => {
		for (const range of looked) sink.look(sheet, range)
	}
	const replay = (which: "before" | "after", sink: SheetChangeSink): void => {
		restoreStyles(store, log, which)
		document.workbook.changes.sheetRows(sheet)
		restoreAxisChange(document, axes, which, sink)
		report(sink)
	}

	if (axes.changed) changes.layout(sheet)

	report(changes)

	return {
		bytes: log.bytes + axes.bytes,
		undo: sink => replay("before", sink),
		redo: sink => replay("after", sink)
	}
}

// What a patch asks that the styles part cannot hold: a number format code the Format Cells dialog refuses too
// (numFmtProblem), or a font name with a character XML cannot carry.
function unwritable(patch: SheetFormatPatch): string | null {
	if (patch.numFmt !== undefined && numFmtProblem(patch.numFmt) !== null) return "numFmt"
	if (patch.font !== undefined && patch.font !== null && (patch.font.trim() === "" || unwritableText(patch.font))) return "font"

	return null
}

// format.apply: each range's cells take their format patched (whole columns and rows their axes' too); cells
// without one of their own start from their row's or column's.
export function applyFormat(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	patch: SheetFormatPatch,
	changes: SheetChangeSink,
	slicing?: Slicing
): Promise<SheetHandled> {
	const detail = unwritable(patch)

	if (detail !== null) return Promise.resolve({ refused: { reason: "unsupported", detail } })

	const patcher = new XfPatcher(document, patch)
	const patchCell: PositionedPatch = xf => patcher.xf(xf)

	return formatPlan(
		document,
		sheet,
		ranges.map(range => ({ range, patch: patchCell })),
		changes,
		slicing,
		patch.size !== undefined
	)
}

// format.clear: every cell back to the default format (cellXfs 0), whole columns' and rows' axes too.
export function clearFormats(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	changes: SheetChangeSink,
	slicing?: Slicing
): Promise<SheetHandled> {
	const cleared: PositionedPatch = () => 0

	return formatPlan(
		document,
		sheet,
		ranges.map(range => ({ range, patch: cleared })),
		changes,
		slicing,
		true
	)
}
