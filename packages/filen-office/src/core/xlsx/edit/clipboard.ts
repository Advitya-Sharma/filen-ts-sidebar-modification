import {
	SHEET_MAX_COPY_CELLS,
	SHEET_MAX_EDIT_CELLS,
	type SheetId,
	type SheetQueryResult,
	type SheetRange,
	type SheetRefusal
} from "../../sheet/protocol"
import type { SheetChangeSink, SheetHandled } from "../../sheet/registry"
import { pastedCells, type GridClip } from "../../sheet/clipboard/tsv"
import { cellKey } from "../../sheet/model/model"
import { rewriteFormula, sameSheet } from "../../sheet/refs/formulaRefs"
import { sliceClock, type Slicing } from "../../slicing"
import type { SheetStore } from "../model/store"
import { FormulaKind, type XlsxDefinedName, type XlsxSheetModel, type XlsxSourceCell } from "../model/types"
import { planMove, type MovePlan } from "../refs/move"
import { moveReference, type CellMove } from "../refs/moveRefs"
import { nameWithFormula } from "../refs/plan"
import { parseRef } from "../refs/sqref"
import type { XlsxDocument } from "./document"
import { setCells, splitsArray, unshare, type CellInputs } from "./cells"
import { ABSENT, type CellLog } from "./history"

// The TSV clipboard: a copy is the cells' text as other spreadsheets paste it, with the session's own record of
// what the cells held (formulas, values a format dresses up, formats); a paste of that same text back writes what
// they held, a cut moving its cells.

export interface XlsxClip {
	readonly id: number
	readonly sheet: SheetId
	readonly grid: GridClip
	// Each copied cell's format, by offset in the range (cellKey).
	readonly styles: ReadonlyMap<number, number>
	// The source cells of copied cells holding markup the model does not (a picture in the cell, rich runs), by
	// offset: a paste of this clip carries it (a cut moves it, a copy shares it), a paste of other text does not.
	readonly sources: ReadonlyMap<number, XlsxSourceCell>
}

function field(text: string): string {
	return /[\t\n\r"]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text
}

function lineEnds(text: string): string {
	return text.replace(/\r\n?/g, "\n")
}

export async function copyCells(
	document: XlsxDocument,
	sheet: SheetId,
	range: SheetRange,
	cut: boolean
): Promise<SheetQueryResult<"copy">> {
	const rows = range.endRow - range.startRow + 1
	const cols = range.endCol - range.startCol + 1

	if (rows <= 0 || cols <= 0 || rows * cols > SHEET_MAX_COPY_CELLS) return { refused: "tooLarge" }

	const model = await document.ready(sheet)
	const store = model.store as SheetStore
	const lines: string[][] = Array.from({ length: rows }, () => new Array<string>(cols).fill(""))
	const inputs = new Map<number, string>()
	const styles = new Map<number, number>()
	const sources = new Map<number, XlsxSourceCell>()

	store.forEachCell(range, (row, col, slot) => {
		const view = document.cellView(sheet, row, col)
		const offset = cellKey(row - range.startRow, col - range.startCol)
		const line = lines[row - range.startRow]

		if (line !== undefined) line[col - range.startCol] = field(view.text)
		if (view.input !== null) inputs.set(offset, view.input)
		if (view.xf !== 0) styles.set(offset, view.xf)

		if (store.carriesMarkup(slot)) {
			const source = store.sourceCell(slot, model.part)

			if (source !== null) sources.set(offset, source)
		}
	})

	const text = lines.map(line => line.join("\t")).join("\r\n")
	const entry = document.workbook.sheet(sheet)
	const id = document.nextClip++

	document.clip = {
		id,
		sheet,
		grid: { tsv: text, sheet: entry?.name ?? "", range, cut, inputs },
		styles,
		sources
	}

	return { clip: id, text }
}

// --- A cut's move ---

// A cut moves its cells: every formula and name of the workbook moves its references as moveReference says (they
// follow the cut cells; one wholly to cells the paste writes over becomes #REF!), and so do the formulas of sections
// and parts, with the formats, validations, hyperlinks and merges of the cut cells (refs/move.ts). Its source is
// cleared as cells.clear clears it, and checked the same way first.

interface Repoint {
	readonly sheet: SheetId
	readonly row: number
	readonly col: number
	readonly text: string
	// A shared formula's master whose text changes: its group is unshared first.
	readonly master: boolean
}

interface CutPlan {
	// The cut's cells as they stand.
	readonly cells: readonly { readonly row: number; readonly col: number }[]
	readonly repoints: readonly Repoint[]
	readonly names: { readonly before: readonly XlsxDefinedName[]; readonly after: readonly XlsxDefinedName[] } | null
	readonly parts: MovePlan
}

// The workbook's names as before or after the cut. The engine cannot take a name's new formula: it is dropped, and
// built again from the cells and names as they then are when an edit next needs it.
function setNames(document: XlsxDocument, names: readonly XlsxDefinedName[]): void {
	const workbook = document.workbook

	workbook.definedNames.splice(0, workbook.definedNames.length, ...names)
	workbook.changes.workbookSection("definedNames")
	document.dropEngine()
}

function isLetter(code: number): boolean {
	return (code >= 65 && code <= 90) || (code >= 97 && code <= 122)
}

function isDigit(code: number): boolean {
	return code >= 48 && code <= 57
}

function inside(range: SheetRange, row: number, col: number): boolean {
	return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}

// Whether formula text may refer to cells of `range`: a run of letters then digits that reads as a cell inside it,
// or a whole row or column range. Every reference is such a run, so a formula this says no to refers to none.
function mayReferInto(text: string, range: SheetRange): boolean {
	let i = 0

	while (i < text.length) {
		let j = i

		if (text.charCodeAt(j) === 36) j++

		const lettersStart = j
		let col = 0

		while (j < text.length && isLetter(text.charCodeAt(j))) {
			if (j - lettersStart < 3) col = col * 26 + ((text.charCodeAt(j) | 0x20) - 96)

			j++
		}

		const letters = j - lettersStart

		if (letters === 0) {
			// 1:3 (whole rows) or A:C (whole columns, after its letters) may meet the range.
			if (text.charCodeAt(j) === 58 || (isDigit(text.charCodeAt(j)) && /^\$?\d+:\$?\d/.test(text.slice(j)))) return true

			i = Math.max(j, i + 1)

			continue
		}

		if (text.charCodeAt(j) === 58 && letters <= 3) return true
		if (text.charCodeAt(j) === 36) j++

		const digitsStart = j
		let row = 0

		while (j < text.length && isDigit(text.charCodeAt(j))) {
			if (j - digitsStart < 8) row = row * 10 + text.charCodeAt(j) - 48

			j++
		}

		if (letters <= 3 && j > digitsStart && j - digitsStart <= 7 && inside(range, row - 1, col - 1)) return true

		i = j
	}

	return false
}

// Whether text may name `sheet` as a reference's sheet.
function mayName(text: string, sheet: string): boolean {
	if (!text.includes("!")) return false

	const upper = text.toUpperCase()
	const name = sheet.toUpperCase()

	return upper.includes(name) || (name.includes("'") && upper.includes(name.replaceAll("'", "''")))
}

// A formula on `own` (null for a workbook-wide name) after the move.
function moved(text: string, own: string | null, move: CellMove): string {
	return rewriteFormula(text, ref => moveReference(ref, own, move))
}

// Whether a formula on `own` may refer to the moved or the written-over cells.
function mayMeet(text: string, own: string | null, move: CellMove): boolean {
	return (
		(own !== null && sameSheet(own, move.sheet) && mayReferInto(text, move.range)) ||
		(own !== null && sameSheet(own, move.to) && mayReferInto(text, move.destination)) ||
		mayName(text, move.sheet) ||
		mayName(text, move.to)
	)
}

// The cut's checks (its cells are cleared as cells.clear would clear them) and what it moves, before anything
// changes.
async function planCut(
	document: XlsxDocument,
	cut: XlsxClip,
	sheet: SheetId,
	destination: SheetRange,
	slicing: Slicing | undefined
): Promise<CutPlan | { refused: SheetRefusal }> {
	const source = cut.grid.range
	const sourceModel = await document.ready(cut.sheet)
	const cells: { row: number; col: number }[] = []

	;(sourceModel.store as SheetStore).forEachCell(source, (row, col) => cells.push({ row, col }))

	for (const { row, col } of cells) if (document.locked(cut.sheet, row, col)) return { refused: { reason: "protectedSheet" } }

	const rows = Uint32Array.from(cells, cell => cell.row)
	const cols = Uint32Array.from(cells, cell => cell.col)

	if (splitsArray(sourceModel, rows, cols)) return { refused: { reason: "arrayFormula" } }

	await document.readyAllSheets(slicing)

	if (!document.allReady) return { refused: { reason: "structureLocked", detail: "unreadSheet" } }

	const workbook = document.workbook
	const move: CellMove = {
		sheet: workbook.sheet(cut.sheet)?.name ?? "",
		range: source,
		to: workbook.sheet(sheet)?.name ?? "",
		destination
	}

	// Names move like formulas (their attributes kept).
	const before = workbook.definedNames.slice()
	const after: XlsxDefinedName[] = []

	for (const defined of before) {
		const own = defined.sheet === null ? null : (workbook.sheet(defined.sheet)?.name ?? null)
		const formula = mayMeet(defined.formula, own, move) ? moved(defined.formula, own, move) : defined.formula
		const next = formula === defined.formula ? defined : nameWithFormula(defined, formula)

		if (next === null) return { refused: { reason: "structureLocked", detail: "definedNames" } }

		after.push(next)
	}

	const names = after.some((defined, i) => defined !== before[i]) ? { before, after } : null

	const repoints: Repoint[] = []
	const clock = slicing === undefined ? undefined : sliceClock(slicing)

	for (const entry of workbook.sheets) {
		const model = entry.model as XlsxSheetModel | null

		if (model === null) continue

		const store = model.store as SheetStore
		const formulas = store.formulas
		const own = entry.id === cut.sheet

		for (const block of store.blocks) {
			if (block.formulas === null) continue

			// A block at a time: calls wait in the session's order, so the store does not change meanwhile.
			if (clock?.due === true) await clock.pause()

			for (let i = 0; i < block.rowCount; i++) {
				const row = block.firstRow + (block.rows[i] ?? 0)

				for (let k = block.rowStarts[i] ?? 0; k < (block.rowStarts[i + 1] ?? 0); k++) {
					const id = block.formulas[k] ?? -1

					if (id < 0) continue

					const col = block.cols[k] ?? 0

					// Cleared by the cut, or written over by the paste.
					if (own && inside(source, row, col)) continue
					if (entry.id === sheet && inside(destination, row, col)) continue

					const kind = formulas.kind(id)

					if (kind === FormulaKind.DATA_TABLE) {
						const formula = formulas.formula(id)

						for (const input of [formula.r1, formula.r2]) {
							const cell = input === undefined ? null : parseRef(input)

							if (
								cell !== null &&
								((own && inside(source, cell.startRow, cell.startCol)) ||
									(entry.id === sheet && inside(destination, cell.startRow, cell.startCol)))
							)
								return { refused: { reason: "structureLocked", detail: "dataTable" } }
						}

						continue
					}

					const text = document.formulaText(model, id, row, col)

					if (!mayMeet(text, entry.name, move)) continue

					const next = moved(text, entry.name, move)

					if (next === text) continue

					// As many cells as one edit writes.
					if (repoints.length === SHEET_MAX_EDIT_CELLS) return { refused: { reason: "tooLarge" } }

					repoints.push({ sheet: entry.id, row, col, text: next, master: kind === FormulaKind.SHARED_MASTER })
				}
			}
		}
	}

	const parts = await planMove(workbook, move, slicing)

	return parts.ok ? { cells, repoints, names, parts: parts.plan } : { refused: parts.refusal }
}

// Clears the cut's cells (a shared formula's master unshared first, so its children keep their formulas) and
// moves the references into them, logged into the paste's step; sections and parts with them.
function applyCut(document: XlsxDocument, cut: XlsxClip, plan: CutPlan, log: CellLog, changes: SheetChangeSink): void {
	const source = document.model(cut.sheet)

	if (source === null) return

	for (const { row, col } of plan.cells) {
		const formula = document.state(cut.sheet, row, col).formula

		if (formula >= 0 && source.store.formulas.kind(formula) === FormulaKind.SHARED_MASTER) unshare(document, cut.sheet, row, col, log)
	}

	for (const repoint of plan.repoints) if (repoint.master) unshare(document, repoint.sheet, repoint.row, repoint.col, log)

	for (const { row, col } of plan.cells) {
		log.touch(cut.sheet, row, col, () => document.state(cut.sheet, row, col), true)
		document.write(cut.sheet, row, col, ABSENT)
	}

	for (const repoint of plan.repoints) {
		const model = document.model(repoint.sheet)
		const state = document.state(repoint.sheet, repoint.row, repoint.col)

		if (model === null || state.formula < 0) continue

		const formulas = model.store.formulas
		const before = formulas.formula(state.formula)
		// An array keeps its range; anything else becomes a formula of its own.
		const formula = formulas.add(
			before.kind === FormulaKind.ARRAY
				? { ...before, text: repoint.text }
				: { kind: FormulaKind.NORMAL, text: repoint.text, sharedIndex: -1, ref: null, flags: 0, cellMetadata: -1 }
		)

		log.touch(repoint.sheet, repoint.row, repoint.col, () => state, true)
		document.write(repoint.sheet, repoint.row, repoint.col, { ...state, formula })
	}

	plan.parts.apply()

	if (plan.names !== null) setNames(document, plan.names.after)

	changes.content(cut.sheet, cut.grid.range)
}

// clipboard.paste: the session's own copy when `clip` names it and the text is still its text, else the text read
// as tab-separated values.
export async function pasteCells(
	document: XlsxDocument,
	sheet: SheetId,
	at: { readonly row: number; readonly col: number },
	text: string,
	clipId: number | undefined,
	changes: SheetChangeSink,
	slicing?: Slicing
): Promise<SheetHandled> {
	const held = document.clip
	const own = held !== null && clipId === held.id && lineEnds(text) === lineEnds(held.grid.tsv) ? held : null
	const target = document.workbook.sheet(sheet)

	if (target === undefined) return { refused: { reason: "tooLarge" } }

	const block = pastedCells(text, own?.grid ?? null, at, target.name, SHEET_MAX_EDIT_CELLS)

	if (block === null || block.cells.length === 0) return block === null ? { refused: { reason: "tooLarge" } } : null

	const range = { startRow: at.row, startCol: at.col, endRow: at.row + block.rows - 1, endCol: at.col + block.cols - 1 }
	const styles =
		own === null ? undefined : Int32Array.from(block.cells, cell => own.styles.get(cellKey(cell.row - at.row, cell.col - at.col)) ?? 0)
	const cut = own?.grid.cut === true ? own : null
	const plan = cut === null ? null : await planCut(document, cut, sheet, range, slicing)

	if (plan !== null && "refused" in plan) return plan

	// Cells of this clip that hold markup the model does not carry their source cells (the writer keeps that markup).
	const store = (await document.ready(sheet)).store as SheetStore
	const origins =
		own === null || own.sources.size === 0
			? undefined
			: Float64Array.from(block.cells, cell => {
					const source = own.sources.get(cellKey(cell.row - at.row, cell.col - at.col))

					return source === undefined ? -1 : store.carry(source)
				})
	const inputs: CellInputs = {
		rows: Uint32Array.from(block.cells, cell => cell.row),
		cols: Uint32Array.from(block.cells, cell => cell.col),
		inputs: block.cells.map(cell => cell.input),
		...(styles !== undefined && { styles }),
		...(origins !== undefined && { origins })
	}

	const handled = await setCells(
		document,
		sheet,
		inputs,
		changes,
		cut === null || plan === null ? undefined : log => applyCut(document, cut, plan, log, changes)
	)

	if (handled === null || "refused" in handled) return handled
	if (cut !== null && block.clip !== null) document.clip = { ...cut, sheet, grid: block.clip }

	changes.select(sheet, { activeRow: at.row, activeCol: at.col, ranges: [range] })

	if (plan === null) return handled

	const { parts, names } = plan

	return {
		bytes: handled.bytes,
		undo: async sink => {
			await handled.undo(sink)
			parts.revert()

			if (names !== null) setNames(document, names.before)
		},
		redo: async sink => {
			parts.apply()

			if (names !== null) setNames(document, names.after)

			await handled.redo(sink)
		}
	}
}
