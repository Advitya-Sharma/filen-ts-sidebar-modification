import {
	SHEET_MAX_CELL_TEXT,
	SHEET_MAX_COLS,
	SHEET_MAX_EDIT_CELLS,
	SHEET_MAX_FORMULA_TEXT,
	SHEET_MAX_ROWS,
	type SheetId,
	type SheetLocale,
	type SheetRange
} from "../../sheet/protocol"
import type { SheetChangeSink, SheetEditStep, SheetHandled } from "../../sheet/registry"
import { parseCellInput } from "../../sheet/input/cellInput"
import { storedFormula, wellFormedFormula } from "../../sheet/refs/formulaRefs"
import type { SheetStore } from "../model/store"
import { CellKind, FormulaKind, type XlsxSheetModel } from "../model/types"
import { liveRelationships } from "../refs/state"
import { dateSerial } from "../../sheet/numfmt/date"
import { encodeXstring, type XlsxDocument } from "./document"
import type { RecalculatedCell } from "../../sheet/formula/index"
import { ENGINE_SLICE, StyleLimitError, cellStep, clearFormats, reportLogged, typedXf } from "./format"
import { sliceClock, type SliceClock, type Slicing } from "../../slicing"
import { ABSENT, CellLog, type CellState } from "./history"

// Cell entries (cells.set) and clearing (cells.clear): what is typed, read as Excel reads it in that cell, written
// into the store with its formula and its implied format, then recalculated. Undo writes the cells back as they were.

// What a typed entry is, once read.
type Entry =
	| { readonly type: "empty" }
	| { readonly type: "formula"; readonly text: string }
	| { readonly type: "text"; readonly text: string }
	| { readonly type: "number"; readonly value: number; readonly format?: string }
	| { readonly type: "boolean"; readonly value: boolean }
	| { readonly type: "error"; readonly text: string }

// The error values a cell holds as constants (ECMA-376 Part 1, the formula grammar's error-constant): typed or
// pasted, each is that error, as in Excel, not text.
const ERRORS: ReadonlyMap<string, string> = new Map(
	["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"].map(error => [error, error])
)

const DATE_TIME = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$/
const TIME = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/

function validDay(year: number, month: number, day: number): boolean {
	if (month < 1 || month > 12 || day < 1) return false

	const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0
	const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1] ?? 0

	return day <= days
}

// A date with a time ("2024-03-05 14:30"), or a time alone ("14:30", on the epoch day), as editing a date-time or
// time cell starts from. Null for anything else.
function dateTimeEntry(input: string, date1904: boolean): Entry | null {
	const trimmed = input.trim()
	const dateTime = DATE_TIME.exec(trimmed)
	const time = dateTime === null ? TIME.exec(trimmed) : null
	const [hours, minutes, seconds] = (dateTime === null ? time?.slice(1, 4) : dateTime.slice(4, 7)) ?? []

	if (hours === undefined || minutes === undefined || Number(hours) > 23 || Number(minutes) > 59 || Number(seconds ?? 0) > 59) return null

	const ofDay = (Number(hours) * 3600 + Number(minutes) * 60 + Number(seconds ?? 0)) / 86_400
	const format = seconds === undefined ? "h:mm" : "h:mm:ss"

	if (dateTime === null) return { type: "number", value: ofDay, format }

	const [, year = "", month = "", day = ""] = dateTime

	if (!validDay(Number(year), Number(month), Number(day))) return null

	return {
		type: "number",
		value: dateSerial(Number(year), Number(month), Number(day), date1904) + ofDay,
		format: `yyyy-mm-dd ${format.replace("h", "hh")}`
	}
}

// What is typed into a cell, read as Excel reads it there: into a Text ("@") cell, as text, a formula too. With the
// session's locale, as typed in it (cellInput.ts); without one, in the file's own (en-US) syntax.
export function readEntry(input: string, textFormat: boolean, date1904: boolean, locale?: SheetLocale): Entry {
	if (textFormat && input !== "" && !input.startsWith("'")) return { type: "text", text: input }

	const dateTime = dateTimeEntry(input, date1904)

	if (dateTime !== null) return dateTime

	const parsed = parseCellInput(input, locale)

	switch (parsed.type) {
		case "empty":
			return parsed
		case "formula":
			return { type: "formula", text: parsed.formula }
	}

	const value = parsed.value

	if (typeof value === "string") {
		const error = ERRORS.get(value.toUpperCase())

		return error === undefined || input.startsWith("'") ? { type: "text", text: value } : { type: "error", text: error }
	}
	if (typeof value === "boolean") return { type: "boolean", value }
	if (typeof value === "number")
		return parsed.impliedFormat === undefined ? { type: "number", value } : { type: "number", value, format: parsed.impliedFormat }

	return {
		type: "number",
		value: dateSerial(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate(), date1904),
		...(parsed.impliedFormat !== undefined && { format: parsed.impliedFormat })
	}
}

// What of an entry no cell can hold: a number past a double (1e400 reads as Infinity), text or a formula longer
// than Excel's limits. Null when it fits.
function entryTooLarge(entry: Entry): "number" | "text" | "formula" | null {
	switch (entry.type) {
		case "number":
			return Number.isFinite(entry.value) ? null : "number"
		case "text":
			return entry.text.length > SHEET_MAX_CELL_TEXT ? "text" : null
		case "formula":
			return entry.text.length > SHEET_MAX_FORMULA_TEXT ? "formula" : null
		default:
			return null
	}
}

// The array formulas and data tables of a sheet: their ranges, whose cells change only together.
function arrayRanges(model: XlsxSheetModel): SheetRange[] {
	const formulas = model.store.formulas
	const ranges: SheetRange[] = []

	for (let id = 0; id < formulas.length; id++) {
		const kind = formulas.kind(id)

		if (kind !== FormulaKind.ARRAY && kind !== FormulaKind.DATA_TABLE) continue

		const ref = formulas.formula(id).ref

		// Still the formula of its range's first cell (an edit may have replaced the whole array).
		if (ref !== null && (model.store as SheetStore).formula(model.store.locate(ref.startRow, ref.startCol)) === id) ranges.push(ref)
	}

	return ranges
}

// Whether the cells take part of an array formula's range without the rest of it.
export function splitsArray(model: XlsxSheetModel, rows: ArrayLike<number>, cols: ArrayLike<number>): boolean {
	const ranges = arrayRanges(model)

	if (ranges.length === 0) return false

	for (const range of ranges) {
		let inside = 0

		for (let i = 0; i < rows.length; i++) {
			const row = rows[i] ?? 0
			const col = cols[i] ?? 0

			if (row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol) inside++
		}

		if (inside > 0 && inside < (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)) return true
	}

	return false
}

// The header rows of the sheet's tables: a header cell names its table column (§18.5.1.2), so editing its text
// would rename the column, which the table part keeps on its own.
export async function tableHeaders(document: XlsxDocument, model: XlsxSheetModel): Promise<SheetRange[]> {
	const headers: SheetRange[] = []

	for (const table of await document.sheetTables(model.id))
		if (table.headerRows > 0) headers.push({ ...table.range, endRow: table.range.startRow + table.headerRows - 1 })

	return headers
}

// A shared formula's children (§18.3.1.40) are written as the master's text moved to them: before its master is
// overwritten or moved, each becomes a formula of its own with that text (logged, so undo gives the group back).
export function unshare(document: XlsxDocument, sheet: SheetId, masterRow: number, masterCol: number, log: CellLog): void {
	const model = document.model(sheet)

	if (model === null) return

	const store = model.store as SheetStore
	const formulas = store.formulas
	const master = store.formula(store.locate(masterRow, masterCol))

	if (master < 0 || formulas.kind(master) !== FormulaKind.SHARED_MASTER) return

	const shared = formulas.sharedIndex(master)
	const ref = formulas.formula(master).ref ?? { startRow: masterRow, startCol: masterCol, endRow: masterRow, endCol: masterCol }
	const cells: { row: number; col: number; text: string }[] = []

	store.forEachCell(ref, (row, col, slot) => {
		const id = store.formula(slot)

		if (id >= 0 && (id === master || (formulas.kind(id) === FormulaKind.SHARED_CHILD && formulas.sharedIndex(id) === shared)))
			cells.push({ row, col, text: document.formulaText(model, id, row, col) })
	})

	for (const cell of cells) {
		const state = document.state(sheet, cell.row, cell.col)
		const formula = formulas.add({ kind: FormulaKind.NORMAL, text: cell.text, sharedIndex: -1, ref: null, flags: 0, cellMetadata: -1 })

		log.touch(sheet, cell.row, cell.col, () => state, false)
		document.write(sheet, cell.row, cell.col, { ...state, formula })
	}
}

export function sharedString(document: XlsxDocument, text: string): number {
	const strings = document.workbook.sharedStrings
	const encoded = encodeXstring(text)
	const found = strings.find(encoded)

	if (found >= 0) return found

	const index = strings.append(encoded)

	document.workbook.changes.sharedStringsAppended()

	return index
}

// The cell an entry becomes.
function entryState(
	document: XlsxDocument,
	model: XlsxSheetModel,
	row: number,
	col: number,
	input: string,
	entry: Entry,
	existing: CellState
): CellState {
	if (entry.type === "empty") {
		return existing.present && existing.style !== 0
			? { present: true, kind: CellKind.EMPTY, value: 0, style: existing.style, formula: -1, shown: undefined }
			: ABSENT
	}

	const start = existing.present ? existing.style : document.startingStyle(model, row, col)
	const quoted = input.startsWith("'") && entry.type !== "formula"
	const style = typedXf(document, start, quoted, entry.type === "number" ? entry.format : undefined)

	switch (entry.type) {
		case "formula": {
			const formula = model.store.formulas.add({
				kind: FormulaKind.NORMAL,
				text: encodeXstring(storedFormula(entry.text)),
				sharedIndex: -1,
				ref: null,
				flags: 0,
				cellMetadata: -1
			})

			return { present: true, kind: CellKind.EMPTY, value: 0, style, formula, shown: undefined }
		}
		case "text":
			return {
				present: true,
				kind: CellKind.SHARED_STRING,
				value: sharedString(document, entry.text),
				style,
				formula: -1,
				shown: undefined
			}
		case "number":
			return { present: true, kind: CellKind.NUMBER, value: entry.value, style, formula: -1, shown: undefined }
		case "boolean":
			return { present: true, kind: CellKind.BOOLEAN, value: entry.value ? 1 : 0, style, formula: -1, shown: undefined }
		case "error":
			return {
				present: true,
				kind: CellKind.ERROR,
				value: model.store.strings.intern(entry.text),
				style,
				formula: -1,
				shown: undefined
			}
	}
}

export interface CellInputs {
	readonly rows: ArrayLike<number>
	readonly cols: ArrayLike<number>
	readonly inputs: readonly string[]
	// A format each cell takes as it is (a paste of this workbook's own cells), -1 for the one typing implies.
	readonly styles?: ArrayLike<number>
	// The source cell each carries (XlsxSheetStore.carry, a paste of this workbook's own cells), -1 for none.
	readonly origins?: ArrayLike<number>
	// 1 where the input is in the file's own (en-US) syntax rather than typed in the session's locale: a value or
	// formula a paste carries exactly (this workbook's own copy, clipboard HTML's typed values).
	readonly canonical?: ArrayLike<number>
}

// Writes typed entries (cells.set, a paste). `prepare` runs once the edit is known to go ahead, before the cells
// are written (a cut paste clears its source there), logging into the same step.
export async function setCells(
	document: XlsxDocument,
	sheet: SheetId,
	cells: CellInputs,
	changes: SheetChangeSink,
	prepare?: (log: CellLog) => void,
	slicing?: Slicing
): Promise<SheetHandled> {
	const model = await document.ready(sheet)
	const count = cells.inputs.length
	const clock = slicing === undefined ? undefined : sliceClock(slicing)

	if (count > SHEET_MAX_EDIT_CELLS || cells.rows.length !== count || cells.cols.length !== count)
		return { refused: { reason: "tooLarge" } }

	for (let i = 0; i < count; i++) {
		if ((cells.rows[i] ?? 0) >= SHEET_MAX_ROWS || (cells.cols[i] ?? 0) >= SHEET_MAX_COLS) return { refused: { reason: "tooLarge" } }
		if (document.locked(sheet, cells.rows[i] ?? 0, cells.cols[i] ?? 0)) return { refused: { reason: "protectedSheet" } }
		if ((i & SLICE_MASK) === SLICE_MASK) await due(clock)
	}

	const store = model.store as SheetStore
	const entries: Entry[] = []

	for (let i = 0; i < count; i++) {
		const row = cells.rows[i] ?? 0
		const col = cells.cols[i] ?? 0
		const slot = store.locate(row, col)
		const xf = slot < 0 ? document.startingStyle(model, row, col) : store.style(slot)
		const entry = readEntry(
			cells.inputs[i] ?? "",
			document.numFmt(xf).isText,
			document.date1904,
			cells.canonical?.[i] === 1 ? undefined : document.settings.locale
		)

		if (entry.type === "formula" && !wellFormedFormula(entry.text)) return { refused: { reason: "formula" } }

		const tooLarge = entryTooLarge(entry)

		if (tooLarge !== null) return { refused: { reason: "tooLarge", detail: tooLarge } }

		entries.push(entry)

		if ((i & SLICE_MASK) === SLICE_MASK) await due(clock)
	}

	if (splitsArray(model, cells.rows, cells.cols)) return { refused: { reason: "arrayFormula" } }

	if (
		liveRelationships(document.workbook, model.part, model.relationships.all).some(relationship => relationship.type.endsWith("/table"))
	) {
		const headers = await tableHeaders(document, model)

		for (let i = 0; i < count; i++) {
			const row = cells.rows[i] ?? 0
			const col = cells.cols[i] ?? 0

			if (
				headers.some(range => row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol) &&
				document.cellView(sheet, row, col).text !== (cells.inputs[i] ?? "")
			)
				return { refused: { reason: "tableHeader" } }
		}
	}

	// Built from the cells as they stand before this edit, so the edit is what it recalculates.
	if (document.engine === null && document.allReady && (document.hasFormulas || entries.some(entry => entry.type === "formula"))) {
		await document.loadTables()
		document.buildEngine(changes)
	}

	const log = new CellLog()
	const states: CellState[] = []

	try {
		for (let i = 0; i < count; i++) {
			const row = cells.rows[i] ?? 0
			const col = cells.cols[i] ?? 0

			const state = entryState(document, model, row, col, cells.inputs[i] ?? "", entries[i] as Entry, document.state(sheet, row, col))
			const style = cells.styles?.[i] ?? -1
			const origin = cells.origins?.[i] ?? -1
			const styled = style < 0 ? state : { ...state, present: true, style }

			states.push(origin < 0 ? styled : { ...styled, origin })

			if ((i & SLICE_MASK) === SLICE_MASK) await due(clock)
		}
	} catch (error) {
		if (error instanceof StyleLimitError) return { refused: { reason: "styleLimit" } }

		throw error
	}

	prepare?.(log)

	return writeLogged(document, sheet, cells.rows, cells.cols, states, log, changes, slicing)
}

// Cells between two looks at the slice clock in a sliced edit (a power of two, less one).
const SLICE_MASK = 255

function due(clock: SliceClock | undefined): Promise<void> | undefined {
	return clock?.due === true ? clock.pause() : undefined
}

// Writes an edit's cells (shared formulas split first, each cell logged), recalculates what reads them and reports
// them: the edit's step. With `slicing` (the in-thread host) every stage yields once the slice is up; the session
// answers nothing else until the edit is done, so nothing reads the cells half written.
export async function writeLogged(
	document: XlsxDocument,
	sheet: SheetId,
	rows: ArrayLike<number>,
	cols: ArrayLike<number>,
	states: readonly CellState[],
	log: CellLog,
	changes: SheetChangeSink,
	slicing: Slicing | undefined
): Promise<SheetEditStep> {
	const clock = slicing === undefined ? undefined : sliceClock(slicing)
	const store = (document.model(sheet)?.store ?? null) as SheetStore | null
	const writes: { row: number; col: number; state: CellState }[] = []

	for (let i = 0; i < states.length; i++) {
		const row = rows[i] ?? 0
		const col = cols[i] ?? 0
		const before = document.state(sheet, row, col)

		if (store !== null && before.formula >= 0 && store.formulas.kind(before.formula) === FormulaKind.SHARED_MASTER)
			unshare(document, sheet, row, col, log)

		log.touch(sheet, row, col, () => document.state(sheet, row, col), true)
		writes.push({ row, col, state: states[i] as CellState })

		if ((i & SLICE_MASK) === SLICE_MASK) await due(clock)
	}

	await document.writeCellsSliced(sheet, writes, false, clock)
	await recalculateSliced(document, log, changes, clock)

	const state = (s: SheetId, r: number, c: number): CellState => document.state(s, r, c)

	for (let from = 0; ; from += WRITE_PART) {
		log.seal(state, from, Math.min(log.size, from + WRITE_PART))

		if (from + WRITE_PART >= log.size) break

		await due(clock)
	}

	reportLogged(log, changes, "content")

	return cellStep(document, log, sink => reportLogged(log, sink, "content"), true, slicing)
}

// Log entries sealed between two looks at the slice clock.
const WRITE_PART = 4096

// The engine gets the cells the edit wrote, sheet by sheet, and its results go into the log.
export function recalculateLogged(document: XlsxDocument, log: CellLog, changes: SheetChangeSink): void {
	if (document.engine === null) return

	for (const [sheet, list] of directCells(log)) document.recalculate(sheet, list, log, changes)
}

// recalculateLogged with the engine given the cells ENGINE_SLICE at a time, its evaluation held until all are in (what
// reads them is recalculated once), yielding between parts once the clock's slice is up.
async function recalculateSliced(
	document: XlsxDocument,
	log: CellLog,
	changes: SheetChangeSink,
	clock: SliceClock | undefined
): Promise<void> {
	const engine = document.engine

	if (clock === undefined || engine === null) {
		recalculateLogged(document, log, changes)

		return
	}

	const bySheet = directCells(log)
	let results: RecalculatedCell[]

	try {
		engine.holdEvaluation()

		for (const [sheet, list] of bySheet) {
			const model = document.model(sheet)
			const index = document.engineIndex(sheet)

			if (model === null || index < 0) continue

			for (let start = 0; start < list.length; start += ENGINE_SLICE) {
				engine.set(
					index,
					list.slice(start, start + ENGINE_SLICE).map(cell => document.engineCell(model, cell.row, cell.col))
				)
				await due(clock)
			}
		}

		results = await engine.releaseEvaluation(true, clock)
	} catch {
		// Built again from the cells when next needed, as a failing recalculate leaves it.
		document.dropEngine()

		return
	}

	let first = true

	for (const [sheet, list] of bySheet) {
		document.applyRecalculated(sheet, list, first ? results : [], log, changes)
		first = false
	}
}

// The cells an edit wrote itself, by sheet.
function directCells(log: CellLog): Map<SheetId, { row: number; col: number }[]> {
	const bySheet = new Map<SheetId, { row: number; col: number }[]>()

	for (let i = 0; i < log.size; i++) {
		if (log.direct[i] !== true) continue

		const sheet = log.sheets[i] ?? 0
		let list = bySheet.get(sheet)

		if (list === undefined) {
			list = []
			bySheet.set(sheet, list)
		}

		list.push({ row: log.rows[i] ?? 0, col: log.cols[i] ?? 0 })
	}

	return bySheet
}

// cells.clear: contents (values and formulas; formats stay), formats (back to the default; values stay) or all.
export async function clearCells(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	what: "contents" | "formats" | "all",
	changes: SheetChangeSink,
	slicing?: Slicing
): Promise<SheetHandled> {
	const model = await document.ready(sheet)
	const store = model.store as SheetStore
	const clock = slicing === undefined ? undefined : sliceClock(slicing)
	const targets: { row: number; col: number }[] = []

	for (const range of ranges) {
		let tooMany = false

		store.forEachCell(range, (row, col) => {
			if (targets.length >= SHEET_MAX_EDIT_CELLS) tooMany = true
			else targets.push({ row, col })
		})

		if (tooMany) return { refused: { reason: "tooLarge" } }
	}

	for (const { row, col } of targets) if (document.locked(sheet, row, col)) return { refused: { reason: "protectedSheet" } }

	// Formats clear as format.clear does (whole columns' and rows' own formats too): one implementation.
	if (what === "formats") return clearFormats(document, sheet, ranges, changes)

	const rows = Uint32Array.from(targets, target => target.row)
	const cols = Uint32Array.from(targets, target => target.col)

	if (splitsArray(model, rows, cols)) return { refused: { reason: "arrayFormula" } }

	const log = new CellLog()
	const changedRows: number[] = []
	const changedCols: number[] = []
	const cleared: CellState[] = []

	for (let i = 0; i < targets.length; i++) {
		const { row, col } = targets[i] as { row: number; col: number }
		const before = document.state(sheet, row, col)
		let next: CellState

		if (what === "all") next = ABSENT
		else
			next =
				before.style === 0
					? ABSENT
					: { present: true, kind: CellKind.EMPTY, value: 0, style: before.style, formula: -1, shown: undefined }

		if (
			next.present === before.present &&
			next.kind === before.kind &&
			next.value === before.value &&
			next.style === before.style &&
			next.formula === before.formula
		)
			continue

		changedRows.push(row)
		changedCols.push(col)
		cleared.push(next)

		if ((i & SLICE_MASK) === SLICE_MASK) await due(clock)
	}

	const contents =
		cleared.length === 0 ? null : await writeLogged(document, sheet, changedRows, changedCols, cleared, log, changes, slicing)

	if (what === "contents") return contents

	// Clear All: the contents, then the formats as format.clear clears them; refused formats take the contents back.
	const formats = await clearFormats(document, sheet, ranges, changes)

	if (formats !== null && "refused" in formats) {
		await contents?.undo(changes)

		return formats
	}

	if (formats === null || contents === null) return formats ?? contents

	const done = contents

	return {
		bytes: done.bytes + formats.bytes,
		undo: async sink => {
			await formats.undo(sink)
			await done.undo(sink)
		},
		redo: async sink => {
			await done.redo(sink)
			await formats.redo(sink)
		}
	}
}
