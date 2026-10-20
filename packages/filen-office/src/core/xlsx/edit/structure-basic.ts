import { partKey } from "../../opc/partNames"
import { Relationships } from "../../opc/relationships"
import { sliceClock, type SliceClock, type Slicing } from "../../slicing"
import { decodeUtf8 } from "../../xml/index"
import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetAxis, type SheetId, type SheetRefusal } from "../../sheet/protocol"
import type { SheetChangeSink, SheetEditStep, SheetHandled } from "../../sheet/registry"
import type { AxisEdit } from "../../sheet/model/edits"
import { renameSheetInFormula, rewriteFormula, sameSheet, shiftFormula, shiftSpan } from "../../sheet/refs/formulaRefs"
import { FormulaTable } from "../model/formulas"
import { SheetStore } from "../model/store"
import { SheetStrings } from "../model/strings"
import { CellKind, FormulaKind, type XlsxColRange, type XlsxDefinedName, type XlsxSheetEntry, type XlsxSheetModel } from "../model/types"
import { XlsxSheet } from "../read/index"
import { RefRefused } from "../refs/catalogue"
import { FormulaShift } from "../refs/formulas"
import type { RefJob } from "../refs/markup"
import { rewriteConnections, rewriteControlProperties, rewriteVml } from "../refs/parts"
import { nameWithFormula } from "../refs/plan"
import { readTable, rewriteTable } from "../refs/tables"
import { relationshipType } from "../write/index"
import { utf8Part } from "../write/markup"
import type { XlsxDocument } from "./document"
import { syncEngine, reportLogged } from "./format"
import { CellLog, StoreSnapshot } from "./history"

// Basic structure edits: rows and columns inserted or deleted where every construct referring to cells is one
// this module moves (cells, formulas of every sheet, merges, column formats), a typed refusal naming the
// construct otherwise; sheets added and renamed.

// Worksheet children that hold no reference to the sheet's cells (or one the writer keeps right by itself:
// dimension). Any other child refers to cells this module does not move.
const MOVABLE_SECTIONS: ReadonlySet<string> = new Set([
	"sheetPr",
	"dimension",
	"sheetViews",
	"sheetFormatPr",
	"cols",
	"sheetData",
	"sheetCalcPr",
	"sheetProtection",
	"mergeCells",
	"phoneticPr",
	"printOptions",
	"pageMargins",
	"pageSetup",
	"headerFooter",
	"picture",
	"legacyDrawingHF",
	"drawingHF",
	"customProperties"
])

// Children of other sheets that may name this sheet in a formula or reference.
const NAMING_SECTIONS: ReadonlySet<string> = new Set([
	"conditionalFormatting",
	"dataValidations",
	"hyperlinks",
	"dataConsolidate",
	"oleObjects",
	"controls",
	"webPublishItems",
	"extLst",
	"mc:AlternateContent"
])

// Workbook children that may name a sheet.
const NAMING_WORKBOOK_SECTIONS: ReadonlySet<string> = new Set(["webPublishObjects", "extLst", "mc:AlternateContent"])

// Parts (by content type, the subtype after the last "." or "/") whose formulas may name a sheet, which this
// module does not rewrite: tables (calculated and totals formulas), legacy drawings and form control properties
// (linked cells, input ranges), connections (parameters bound to cells). The catalogue's rewrite of each tells
// whether the edit would change it.
const NAMING_PARTS: ReadonlySet<string> = new Set(["table+xml", "vmldrawing", "controlproperties+xml", "connections+xml"])

// Parts that refer to cells by a sheet's position or id: rows and columns cannot move under them.
const POSITION_PARTS: ReadonlySet<string> = new Set([
	"volatiledependencies+xml",
	"revisionheaders+xml",
	"revisionlog+xml",
	"pivotcachedefinition+xml"
])

// Relationships of the edited sheet to parts that refer to its cells, held outside its markup.
const CELL_RELATIONSHIPS: ReadonlySet<string> = new Set([
	"comments",
	"threadedComment",
	"pivotTable",
	"tableSingleCells",
	"namedSheetView",
	"table",
	"ctrlProp",
	"slicer",
	"timeline"
])

const SHEET_NAME = /^[^\\/?*[\]:]{1,31}$/

function locked(detail: string): { refused: SheetRefusal } {
	return { refused: { reason: "structureLocked", detail } }
}

// Whether a formula refers to `target` from `own` (its own sheet's unqualified references included).
function refersTo(text: string, own: string, target: string): boolean {
	const ownSheet = sameSheet(own, target)

	if (!ownSheet && !text.includes("!")) return false

	let found = false

	rewriteFormula(text, ref => {
		if (
			!ref.external &&
			ref.area !== null &&
			(ref.sheet === null ? ownSheet : sameSheet(ref.sheet, target) || (ref.lastSheet !== null && sameSheet(ref.lastSheet, target)))
		)
			found = true

		return undefined
	})

	return found
}

// Whether a shared formula's references into `target`, as its children repeat them over the group's range
// (`group`, master first), reach the moved rows or columns.
function reaches(
	text: string,
	own: string,
	target: string,
	group: { startRow: number; startCol: number; endRow: number; endCol: number },
	edit: AxisEdit
): boolean {
	const ownSheet = sameSheet(own, target)
	const rows = edit.axis === "rows"
	const span = rows ? group.endRow - group.startRow : group.endCol - group.startCol
	let found = false

	if (!ownSheet && !text.includes("!")) return false

	rewriteFormula(text, ref => {
		const area = ref.area

		if (found || ref.external || area === null) return undefined
		if (!(ref.sheet === null ? ownSheet : sameSheet(ref.sheet, target) || (ref.lastSheet !== null && sameSheet(ref.lastSheet, target))))
			return undefined
		if ((rows && area.kind === "cols") || (!rows && area.kind === "rows")) return undefined

		const end = rows ? Math.max(area.startRow, area.endRow) : Math.max(area.startCol, area.endCol)
		const relative = rows ? !(area.startRowAbs && area.endRowAbs) : !(area.startColAbs && area.endColAbs)

		if (end + (relative ? span : 0) >= edit.at) found = true

		return undefined
	})

	return found
}

// Whether a section or workbook child names the edited sheet (in a formula or a sheet attribute).
function namesSheet(bytes: Uint8Array | null, formulas: FormulaShift): boolean {
	return bytes !== null && bytes.length > 0 && formulas.namesTarget(decodeUtf8(bytes))
}

function sectionNames(model: XlsxSheetModel, formulas: FormulaShift): boolean {
	return model.sections.some(section => NAMING_SECTIONS.has(section.name) && namesSheet(section.bytes, formulas))
}

function subtype(contentType: string): string {
	const type = contentType.toLowerCase().split(";")[0]?.trim() ?? ""

	return type.slice(Math.max(type.lastIndexOf("."), type.lastIndexOf("/")) + 1)
}

// Why the workbook's parts or workbook children keep the edit `op` (rows or columns of a sheet moved, or the sheet
// renamed) from going ahead, or null: one whose formulas it would change, which this module does not rewrite; one
// that refers to cells by a sheet's position (rows or columns moved); one another edit rewrote.
async function partsLock(document: XlsxDocument, formulas: FormulaShift, slicing: Slicing | undefined): Promise<SheetRefusal | null> {
	const workbook = document.workbook
	const axis = formulas.op.kind === "axis"

	for (const section of workbook.workbookSections)
		if (NAMING_WORKBOOK_SECTIONS.has(section.name) && namesSheet(section.bytes, formulas))
			return { reason: "structureLocked", detail: section.name }

	const changed = new Map<string, boolean>()
	const owners = new Map<string, { entry: XlsxSheetEntry; relationship: string }>()

	for (const [part, write] of workbook.changes.parts) changed.set(partKey(part), write !== null)

	for (const entry of workbook.sheets)
		for (const relationship of entry.model?.relationships.all ?? [])
			if (!relationship.external && relationship.partName !== undefined)
				owners.set(partKey(relationship.partName), { entry, relationship: relationship.id })

	try {
		for (const part of workbook.parts.parts()) {
			const type = subtype(part.contentType)

			if (axis && POSITION_PARTS.has(type)) return { reason: "structureLocked", detail: type }
			if (!NAMING_PARTS.has(type)) continue

			const key = partKey(part.name)
			const write = changed.get(key)

			if (write === false) continue
			if (write === true) return { reason: "structureLocked", detail: `${type}: edited` }

			const owner = owners.get(key) ?? null
			const job: RefJob = { own: owner?.entry.name ?? null, shift: null, formulas, touch: () => undefined }
			const bytes = utf8Part(await workbook.parts.read(part.name, slicing))
			let after = bytes

			if (type === "table+xml") {
				if (owner !== null)
					after = rewriteTable(readTable(part.name, part.contentType, owner.entry.id, owner.relationship, bytes, []), null, job)
			} else if (type === "vmldrawing") {
				after = rewriteVml(bytes, job, part.name)
			} else if (type === "controlproperties+xml") {
				after = rewriteControlProperties(bytes, job, part.name)
			} else {
				after = rewriteConnections(bytes, job, part.name)
			}

			if (after !== bytes) return { reason: "structureLocked", detail: type }
		}
	} catch (error) {
		if (error instanceof RefRefused) return error.refusal

		throw error
	}

	return null
}

// Why rows or columns of `target` cannot move, or null.
async function structureLock(
	document: XlsxDocument,
	target: XlsxSheetEntry,
	edit: AxisEdit,
	slicing: Slicing | undefined
): Promise<SheetRefusal | null> {
	const model = target.model

	if (model === null) return { reason: "structureLocked", detail: "unreadSheet" }

	for (const section of model.sections)
		if (!MOVABLE_SECTIONS.has(section.name)) return { reason: "structureLocked", detail: section.name }

	for (const relationship of model.relationships.all) {
		const type = relationship.type.slice(relationship.type.lastIndexOf("/") + 1)

		if (!relationship.external && CELL_RELATIONSHIPS.has(type)) return { reason: "structureLocked", detail: type }
	}

	const workbook = document.workbook
	const moved = new FormulaShift({ kind: "axis", target: target.name, shift: edit })

	if (workbook.definedNames.length > 0) return { reason: "structureLocked", detail: "definedNames" }
	if (workbook.workbookSections.some(section => section.name === "pivotCaches")) return { reason: "pivot", detail: "pivotCaches" }

	for (const entry of workbook.sheets) {
		if (entry.kind !== "worksheet") return { reason: "structureLocked", detail: entry.kind }
		if (entry.model === null) return { reason: "structureLocked", detail: "unreadSheet" }
		if (entry.model.sections.some(section => section.name === "drawing")) return { reason: "structureLocked", detail: "drawing" }
		if (entry !== target && sectionNames(entry.model, moved)) return { reason: "structureLocked", detail: "reference" }

		const formulas = entry.model.store.formulas

		for (let id = 0; id < formulas.length; id++) {
			const kind = formulas.kind(id)

			if (
				(kind === FormulaKind.ARRAY || kind === FormulaKind.DATA_TABLE) &&
				(entry === target || refersTo(formulas.text(id), entry.name, target.name))
			)
				return { reason: "arrayFormula" }
		}
	}

	return partsLock(document, moved, slicing)
}

// Formula text a quick scan can move: cell references and ranges among operators, numbers and function names; no
// sheet names, strings, structured references, errors, arrays, whole rows or columns. Anything else (or anything
// next to a reference that could make it a name) goes through the formula parser (shiftFormula), whose text this
// matches on what it takes.

function isLetter(code: number): boolean {
	return (code >= 65 && code <= 90) || (code >= 97 && code <= 122)
}

function isDigit(code: number): boolean {
	return code >= 48 && code <= 57
}

// What may continue a name or a number, as the parser reads a word: letters, digits, _ . $ ? \ and every code
// unit from 128 on.
function isNameChar(code: number): boolean {
	return isLetter(code) || isDigit(code) || code === 95 || code === 46 || code === 36 || code === 63 || code === 92 || code >= 128
}

// Characters the scan leaves to the parser: ' " ! [ ] # { } @ \ and a colon outside a cell range.
function isParserOnly(code: number): boolean {
	return (
		code === 39 ||
		code === 34 ||
		code === 33 ||
		code === 91 ||
		code === 93 ||
		code === 35 ||
		code === 123 ||
		code === 125 ||
		code === 64 ||
		code === 92 ||
		code === 58
	)
}

// One cell reference at `at`: [$]letters[$]digits, or -1 when there is none.
interface QuickCell {
	end: number
	colAbs: boolean
	col: number
	rowAbs: boolean
	row: number
}

const scratch: QuickCell = { end: -1, colAbs: false, col: 0, rowAbs: false, row: 0 }
const second: QuickCell = { end: -1, colAbs: false, col: 0, rowAbs: false, row: 0 }

function quickCell(text: string, at: number, out: QuickCell): boolean {
	let i = at

	out.colAbs = text.charCodeAt(i) === 36

	if (out.colAbs) i++

	let col = 0
	const lettersStart = i

	while (i < text.length && isLetter(text.charCodeAt(i)) && i - lettersStart < 4) {
		col = col * 26 + ((text.charCodeAt(i) | 0x20) - 96)
		i++
	}

	if (i === lettersStart || i - lettersStart > 3) return false

	out.rowAbs = text.charCodeAt(i) === 36

	if (out.rowAbs) i++

	let row = 0
	const digitsStart = i

	while (i < text.length && isDigit(text.charCodeAt(i)) && i - digitsStart < 8) {
		row = row * 10 + text.charCodeAt(i) - 48
		i++
	}

	if (i === digitsStart || i - digitsStart > 7) return false

	out.col = col - 1
	out.row = row - 1
	out.end = i

	return true
}

function colLetters(col: number): string {
	let name = ""

	for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name

	return name
}

function cellText(colAbs: boolean, col: number, rowAbs: boolean, row: number): string {
	return `${colAbs ? "$" : ""}${colLetters(col)}${rowAbs ? "$" : ""}${String(row + 1)}`
}

// A formula on the target sheet after rows or columns of it moved, or null when the quick scan cannot tell.
export function quickShift(text: string, edit: AxisEdit): string | null {
	const rows = edit.axis === "rows"
	const limit = rows ? SHEET_MAX_ROWS : SHEET_MAX_COLS
	let out = ""
	let copied = 0
	let i = 0

	while (i < text.length) {
		const code = text.charCodeAt(i)

		if (isParserOnly(code)) return null

		if (!isLetter(code) && code !== 36) {
			// A number or a name not starting with a letter: skipped whole (1E5 is not E5).
			if (isNameChar(code)) {
				while (i < text.length && isNameChar(text.charCodeAt(i))) i++
			} else {
				i++
			}

			continue
		}

		if (!quickCell(text, i, scratch)) {
			// A word (SUM, TRUE): skipped whole.
			while (i < text.length && isNameChar(text.charCodeAt(i))) i++

			continue
		}

		let end = scratch.end
		const range = text.charCodeAt(end) === 58

		if (range && !quickCell(text, end + 1, second)) return null
		if (range) end = second.end

		const next = end < text.length ? text.charCodeAt(end) : -1

		// Followed by more of a name, or a call (LOG10( ): the parser tells.
		if (next !== -1 && (isNameChar(next) || next === 40)) return null

		const startCol = scratch.col
		const startRow = scratch.row
		const endCol = range ? second.col : startCol
		const endRow = range ? second.row : startRow

		if (
			Math.max(startCol, endCol) >= SHEET_MAX_COLS ||
			Math.max(startRow, endRow) >= SHEET_MAX_ROWS ||
			Math.min(startRow, endRow, startCol, endCol) < 0 ||
			(rows ? startRow > endRow : startCol > endCol)
		)
			return null

		const span = shiftSpan(rows ? startRow : startCol, rows ? endRow : endCol, edit, limit, range)

		if (span === null || span[0] !== (rows ? startRow : startCol) || span[1] !== (rows ? endRow : endCol)) {
			let moved = "#REF!"

			if (span !== null) {
				const first = rows
					? cellText(scratch.colAbs, startCol, scratch.rowAbs, span[0])
					: cellText(scratch.colAbs, span[0], scratch.rowAbs, startRow)

				moved = range
					? `${first}:${rows ? cellText(second.colAbs, endCol, second.rowAbs, span[1]) : cellText(second.colAbs, span[1], second.rowAbs, endRow)}`
					: first
			}

			out += text.slice(copied, i) + moved
			copied = end
		}

		i = end
	}

	return copied === 0 ? text : out + text.slice(copied)
}

interface FormulaMove {
	readonly sheet: SheetId
	readonly ids: number[]
	readonly rows: number[]
	readonly cols: number[]
	readonly before: string[]
	readonly after: string[]
}

// Each formula cell of every sheet whose text the edit changes takes the moved text, at its position before the
// move: an ordinary formula in place (its record is its cell's alone; `moves` keeps both texts for undo); a shared
// formula group that refers to the target or moves is split into formulas of their own first (logged), as its
// children could no longer be derived from their master.
async function moveFormulas(
	document: XlsxDocument,
	target: XlsxSheetEntry,
	edit: AxisEdit,
	log: CellLog,
	clock: SliceClock | undefined
): Promise<FormulaMove[]> {
	const moved: FormulaMove[] = []

	for (const entry of document.workbook.sheets) {
		const model = entry.model

		if (model === null) continue

		const store = model.store as SheetStore
		const formulas = store.formulas as FormulaTable
		const groups = new Map<number, boolean>()
		const own = entry === target
		const splits: { row: number; col: number; text: string }[] = []
		const move: FormulaMove = { sheet: entry.id, ids: [], rows: [], cols: [], before: [], after: [] }
		const seen = new Uint8Array(formulas.length)

		for (const block of store.blocks) {
			if (block.formulas === null) continue

			// A block at a time: the in-thread host's slices end between blocks (the store does not change meanwhile:
			// calls wait in the session's order).
			if (clock?.due === true) await clock.pause()

			for (let i = 0; i < block.rowCount; i++) {
				const row = block.firstRow + (block.rows[i] ?? 0)

				for (let k = block.rowStarts[i] ?? 0; k < (block.rowStarts[i + 1] ?? 0); k++) {
					const id = block.formulas[k] ?? -1

					if (id < 0) continue

					const col = block.cols[k] ?? 0
					const kind = formulas.kind(id)
					const text = formulas.text(id)

					if (kind === FormulaKind.SHARED_MASTER || kind === FormulaKind.SHARED_CHILD) {
						const shared = formulas.sharedIndex(id)
						let split = groups.get(shared)

						if (split === undefined) {
							const master = formulas.master(shared)
							const ref = master < 0 ? null : formulas.ref(master)
							const reach = ref === null ? -1 : edit.axis === "rows" ? ref.endRow : ref.endCol

							split =
								master < 0 ||
								ref === null ||
								(own && reach >= edit.at) ||
								reaches(formulas.text(master), entry.name, target.name, ref, edit)
							groups.set(shared, split)
						}

						if (split)
							splits.push({
								row,
								col,
								text: shiftFormula(document.formulaText(model, id, row, col), entry.name, target.name, edit).formula
							})

						continue
					}

					if (!own && !text.includes("!")) continue

					const shifted = (own ? quickShift(text, edit) : null) ?? shiftFormula(text, entry.name, target.name, edit).formula

					if (shifted === text) continue

					// A record two cells share (never written so) is split rather than changed for both.
					if (kind !== FormulaKind.NORMAL || seen[id] === 1) {
						splits.push({ row, col, text: shifted })

						continue
					}

					seen[id] = 1
					move.ids.push(id)
					move.rows.push(row)
					move.cols.push(col)
					move.before.push(text)
					move.after.push(shifted)
				}
			}
		}

		for (const { row, col, text } of splits) {
			const state = document.state(entry.id, row, col)
			const formula = formulas.add({ kind: FormulaKind.NORMAL, text, sharedIndex: -1, ref: null, flags: 0, cellMetadata: -1 })

			log.touch(entry.id, row, col, () => state, true)
			document.write(entry.id, row, col, { ...state, formula })
		}

		if (move.ids.length > 0) {
			moved.push(move)
			await setFormulaTexts(document, move, "after", clock)
		}
	}

	return moved
}

// Ordinary formulas' texts as before or after a move, at their positions before it; their rows are rewritten.
async function setFormulaTexts(
	document: XlsxDocument,
	move: FormulaMove,
	which: "before" | "after",
	clock: SliceClock | undefined
): Promise<void> {
	const model = document.model(move.sheet)

	if (model === null) return

	const store = model.store as SheetStore
	const formulas = store.formulas as FormulaTable
	const texts = which === "after" ? move.after : move.before

	for (let i = 0; i < move.ids.length; i++) {
		if ((i & 1023) === 0 && clock?.due === true) await clock.pause()

		const id = move.ids[i] ?? 0
		const row = move.rows[i] ?? 0
		const col = move.cols[i] ?? 0
		const slot = store.locate(row, col)

		formulas.set(id, { ...formulas.formula(id), text: texts[i] ?? "" })

		// The same cell written over itself: its row is rewritten at save.
		if (slot >= 0) store.setCell(row, col, store.kind(slot), store.value(slot), store.style(slot), id)
	}

	document.workbook.changes.sheetRows(move.sheet)
	document.workbook.changes.formulasChanged()
}

function movedMerges(merges: Uint32Array, edit: AxisEdit): Uint32Array {
	const next: number[] = []
	const rows = edit.axis === "rows"

	for (let i = 0; i + 3 < merges.length; i += 4) {
		const quad = [merges[i] ?? 0, merges[i + 1] ?? 0, merges[i + 2] ?? 0, merges[i + 3] ?? 0]
		const span = shiftSpan(rows ? quad[0]! : quad[1]!, rows ? quad[2]! : quad[3]!, edit, rows ? SHEET_MAX_ROWS : SHEET_MAX_COLS, true)

		if (span === null) continue

		if (rows) {
			quad[0] = span[0]
			quad[2] = span[1]
		} else {
			quad[1] = span[0]
			quad[3] = span[1]
		}

		// A merge left one cell wide is no merge.
		if (quad[0] === quad[2] && quad[1] === quad[3]) continue

		next.push(...quad)
	}

	return Uint32Array.from(next)
}

function movedCols(cols: readonly XlsxColRange[], edit: AxisEdit): readonly XlsxColRange[] {
	if (edit.axis === "rows") return cols

	const next: XlsxColRange[] = []

	for (const range of cols) {
		const span = shiftSpan(range.first, range.last, edit, SHEET_MAX_COLS, true)

		if (span !== null) next.push({ ...range, first: span[0], last: span[1] })
	}

	return next
}

function engineSet(document: XlsxDocument, sheet: SheetId, cells: readonly { row: number; col: number }[]): void {
	const engine = document.engine
	const model = document.model(sheet)
	const index = document.engineIndex(sheet)

	if (engine === null || model === null || index < 0 || cells.length === 0) return

	document.guarded(() =>
		engine.set(
			index,
			cells.map(cell => document.engineCell(model, cell.row, cell.col))
		)
	)
}

// axis.insert and axis.delete.
export async function moveAxis(
	document: XlsxDocument,
	sheet: SheetId,
	axis: SheetAxis,
	type: "insert" | "delete",
	at: number,
	count: number,
	changes: SheetChangeSink,
	slicing?: Slicing
): Promise<SheetHandled> {
	const target = document.entry(sheet)
	const limit = axis === "rows" ? SHEET_MAX_ROWS : SHEET_MAX_COLS

	if (target === undefined || target.kind !== "worksheet") return locked("sheet")
	if (!Number.isInteger(at) || !Number.isInteger(count) || at < 0 || count <= 0 || at + count > limit)
		return { refused: { reason: "tooLarge" } }

	const action = type === "insert" ? (axis === "rows" ? "insertRows" : "insertColumns") : axis === "rows" ? "deleteRows" : "deleteColumns"
	const protection = document.protects(sheet, action)

	if (protection !== null) return { refused: protection }

	await document.readyAllSheets()

	if (!document.allReady) return locked("unreadSheet")

	const lock = await structureLock(document, target, { type, axis, at, count }, slicing)

	if (lock !== null) return { refused: lock }

	const model = target.model as XlsxSheetModel
	const store = model.store as SheetStore

	// Excel will not push cells off the sheet (nor would an undo bring them back).
	if (type === "insert" && (axis === "rows" ? store.rowExtent : store.colExtent) + count > limit)
		return { refused: { reason: "tooLarge", detail: "cells" } }

	const edit: AxisEdit = { type, axis, at, count }
	const index = document.engineIndex(sheet)
	const engine = document.engine

	const engineRefusal = engine === null || index < 0 ? null : engine.moveRefusal(index, edit)

	if (engineRefusal === "spill") return { refused: { reason: "arrayFormula", detail: "spill" } }
	if (engineRefusal === "tooLarge") return { refused: { reason: "tooLarge", detail: "engine" } }

	// The cells a deletion takes: given back to the engine on undo, and their shared strings' count.
	const range =
		axis === "rows"
			? { startRow: at, startCol: 0, endRow: at + count - 1, endCol: SHEET_MAX_COLS - 1 }
			: { startRow: 0, startCol: at, endRow: SHEET_MAX_ROWS - 1, endCol: at + count - 1 }
	const pre = new CellLog()
	const post = new CellLog()
	const merges = { before: model.merges, after: movedMerges(model.merges, edit) }
	const cols = { before: model.cols, after: movedCols(model.cols, edit) }
	const removed: { row: number; col: number }[] = []
	let strings = 0
	let snapshot: StoreSnapshot | null = null

	const moves = await moveFormulas(document, target, edit, pre, slicing === undefined ? undefined : sliceClock(slicing))
	pre.seal((s, r, c) => document.state(s, r, c))

	if (type === "delete") {
		store.forEachCell(range, (row, col, slot) => {
			removed.push({ row, col })

			if (store.kind(slot) === CellKind.SHARED_STRING) strings++
		})
	}

	const changed = document.workbook.changes
	const mark = (): void => {
		changed.sheetRows(sheet)

		if (merges.after.length !== merges.before.length || merges.after.some((value, i) => value !== merges.before[i]))
			changed.sheetSection(sheet, "mergeCells")

		if (cols.after !== cols.before) changed.sheetSection(sheet, "cols")
		if (pre.size > 0 || moves.length > 0) changed.formulasChanged()
	}
	const forward = (): void => {
		if (type === "delete") snapshot = new StoreSnapshot(store, axis === "cols")

		if (axis === "rows") {
			if (type === "insert") store.insertRows(at, count)
			else store.deleteRows(at, count)
		} else if (type === "insert") {
			store.insertCols(at, count)
		} else {
			store.deleteCols(at, count)
		}

		if (strings > 0) {
			document.workbook.sharedStrings.adjustReferences(-strings)
			changed.sharedStringsAppended()
		}

		model.merges = merges.after
		model.cols = cols.after
		document.shiftShown(sheet, edit)
		mark()
	}
	const engineForward = () =>
		engine === null || index < 0
			? []
			: document.guarded(() => (type === "insert" ? engine.insert(index, axis, at, count) : engine.remove(index, axis, at, count)))
	const report = (sink: SheetChangeSink): void => {
		// Formula texts that moved in place: on the target, those before the moved rows or columns (the rest are in
		// the moved region).
		for (const move of moves) {
			let startRow = Infinity
			let startCol = Infinity
			let endRow = -1
			let endCol = -1

			for (let i = 0; i < move.rows.length; i++) {
				const row = move.rows[i] ?? 0
				const col = move.cols[i] ?? 0

				if (move.sheet === sheet && (axis === "rows" ? row : col) >= at) continue

				startRow = Math.min(startRow, row)
				endRow = Math.max(endRow, row)
				startCol = Math.min(startCol, col)
				endCol = Math.max(endCol, col)
			}

			if (endRow >= 0) sink.content(move.sheet, { startRow, startCol, endRow, endCol })
		}

		reportLogged(pre, sink, "content")
		reportLogged(post, sink, "content")
		sink.moved(sheet, axis, at, type === "insert" ? count : -count)
		sink.layout(sheet)
	}

	const clockOf = (): SliceClock | undefined => (slicing === undefined ? undefined : sliceClock(slicing))
	const turn = async (clock: SliceClock | undefined): Promise<void> => {
		if (clock?.due === true) await clock.pause()
	}
	const clock = clockOf()

	await turn(clock)
	forward()
	await turn(clock)
	document.apply(engineForward(), post, changes)
	post.seal((s, r, c) => document.state(s, r, c))
	report(changes)

	const step: SheetEditStep = {
		bytes:
			moves.reduce(
				(bytes, move) =>
					bytes +
					move.ids.length * 40 +
					move.before.reduce((sum, text, i) => sum + (text.length + (move.after[i]?.length ?? 0)) * 2, 0),
				0
			) +
			pre.bytes +
			post.bytes +
			removed.length * 16 +
			((snapshot as StoreSnapshot | null)?.bytes ?? 0) +
			merges.after.byteLength,
		undo: async sink => {
			const clock = clockOf()

			post.restore(document.restoreCell, "before")

			if (type === "insert") {
				if (axis === "rows") store.deleteRows(at, count)
				else store.deleteCols(at, count)
			} else {
				;(snapshot as StoreSnapshot | null)?.restore(store)

				if (strings > 0) document.workbook.sharedStrings.adjustReferences(strings)
			}

			model.merges = merges.before
			model.cols = cols.before
			document.shiftShown(sheet, { type: type === "insert" ? "delete" : "insert", axis, at, count })
			mark()
			pre.restore(document.restoreCell, "before")

			for (const move of moves) await setFormulaTexts(document, move, "before", clock)

			await turn(clock)

			const live = document.engine

			if (live !== null && index >= 0) {
				document.guarded(() => (type === "insert" ? live.remove(index, axis, at, count) : live.insert(index, axis, at, count)))

				if (type === "delete") engineSet(document, sheet, removed)

				syncEngine(document, pre)

				// A deletion's #REF! does not come back with the rows: the engine gets the formulas again.
				if (type === "delete")
					for (const move of moves)
						engineSet(
							document,
							move.sheet,
							move.rows.map((row, i) => ({ row, col: move.cols[i] ?? 0 }))
						)
			}

			report(sink)
		},
		redo: async sink => {
			const clock = clockOf()

			pre.restore(document.restoreCell, "after")

			for (const move of moves) await setFormulaTexts(document, move, "after", clock)

			await turn(clock)
			forward()
			await turn(clock)
			engineForward()
			post.restore(document.restoreCell, "after")
			report(sink)
		}
	}

	return step
}

// --- Sheets ---

// Excel's rules for a sheet name: 1-31 characters, none of \ / ? * [ ] :, no apostrophe first or last, not
// "History" (reserved), unique whatever the case.
export function sheetNameRefusal(document: XlsxDocument, name: string, except: SheetId | null): SheetRefusal | null {
	if (!SHEET_NAME.test(name) || name.startsWith("'") || name.endsWith("'") || name.toUpperCase() === "HISTORY")
		return { reason: "sheetName", detail: "invalid" }

	if (document.workbook.sheets.some(entry => entry.id !== except && sameSheet(entry.name, name)))
		return { reason: "sheetName", detail: "taken" }

	return null
}

function newSheetPart(document: XlsxDocument): string {
	const taken = new Set(document.workbook.sheets.map(entry => entry.part.toLowerCase()))

	for (let n = 1; ; n++) {
		const name = `xl/worksheets/sheet${String(n)}.xml`

		if (!taken.has(name) && !document.workbook.opc.has(name)) return name
	}
}

function relativeTarget(from: string, part: string): string {
	const directory = from.slice(0, from.lastIndexOf("/") + 1)

	return part.startsWith(directory) ? part.slice(directory.length) : `/${part}`
}

// sheet.add: an empty worksheet at `index` (after the last by default), named as asked or "SheetN".
export async function addSheet(
	document: XlsxDocument,
	requested: string | undefined,
	index: number | undefined,
	changes: SheetChangeSink
): Promise<SheetHandled> {
	const workbook = document.workbook
	const sheets = workbook.sheets
	let name = requested

	if (name === undefined) {
		const taken = new Set(sheets.map(entry => entry.name.toUpperCase()))
		let n = sheets.length + 1

		while (taken.has(`SHEET${String(n)}`)) n++

		name = `Sheet${String(n)}`
	}

	const refusal = sheetNameRefusal(document, name, null)

	if (refusal !== null) return { refused: refusal }

	const at = index === undefined ? sheets.length : Math.max(0, Math.min(sheets.length, Math.floor(index)))

	// Local names keep their sheet by position (localSheetId): a sheet put before others would move them.
	if (at < sheets.length && workbook.definedNames.some(defined => defined.sheet !== null)) return locked("definedNames")
	if (sheets.length >= workbook.limits.maxSheets) return { refused: { reason: "tooLarge" } }

	await workbook.parts.relationships(workbook.workbookPart)

	const part = newSheetPart(document)
	const relationship = workbook.changes.addRelationship(workbook.workbookPart, {
		type: relationshipType(workbook.strict, "worksheet"),
		target: relativeTarget(workbook.workbookPart, part)
	})
	const id = workbook.allocateSheetId()
	const model = new XlsxSheet({
		id,
		part,
		store: new SheetStore([], new SheetStrings(), new FormulaTable()),
		sections: [],
		relationships: new Relationships([]),
		rootStart: 0,
		rootOpenEnd: 0,
		rootEnd: 0,
		prefix: "",
		utf16: false
	})
	const entry: XlsxSheetEntry = {
		id,
		name,
		state: "visible",
		kind: "worksheet",
		sheetId: sheets.reduce((highest, sheet) => Math.max(highest, sheet.sheetId), 0) + 1,
		relationshipId: relationship.id,
		part,
		model
	}
	const activeBefore = workbook.activeTab
	const activeAfter = at <= activeBefore && at < sheets.length ? activeBefore + 1 : activeBefore
	const added = (sink: SheetChangeSink): void => {
		sheets.splice(at, 0, entry)
		workbook.activeTab = activeAfter
		workbook.changes.workbookSection("sheets")

		if (activeAfter !== activeBefore) workbook.changes.workbookSection("bookViews")

		const engine = document.engine

		if (engine !== null) {
			engine.addSheet(entry.name)
			document.engineSheets.push(id)
		}

		sink.sheets()
	}

	added(changes)

	return {
		bytes: 512,
		undo: sink => {
			const position = sheets.indexOf(entry)

			if (position >= 0) sheets.splice(position, 1)

			workbook.activeTab = activeBefore
			workbook.changes.removeRelationship(workbook.workbookPart, relationship.id)

			const engine = document.engine

			if (engine !== null && document.engineSheets.at(-1) === id) {
				document.guarded(() => engine.removeLastSheet())
				document.engineSheets.pop()
			}

			sink.sheets()
		},
		redo: sink => {
			workbook.changes.addRelationship(workbook.workbookPart, { type: relationship.type, target: relationship.target })
			added(sink)
		}
	}
}

// sheet.rename: the name, and every formula and defined name that names the sheet.
export async function renameSheet(document: XlsxDocument, sheet: SheetId, name: string, changes: SheetChangeSink): Promise<SheetHandled> {
	const target = document.entry(sheet)

	if (target === undefined) return { refused: { reason: "sheetName", detail: "missing" } }
	if (target.name === name) return null

	const refusal = sheetNameRefusal(document, name, sheet)

	if (refusal !== null) return { refused: refusal }

	await document.readyAllSheets()

	const workbook = document.workbook
	const from = target.name

	const naming = new FormulaShift({ kind: "renameSheet", target: from, name })

	if (!document.allReady) return locked("unreadSheet")
	if (workbook.workbookSections.some(section => section.name === "pivotCaches"))
		return { refused: { reason: "pivot", detail: "pivotCaches" } }

	for (const entry of workbook.sheets) {
		if (entry.kind !== "worksheet") return locked(entry.kind)

		const model = entry.model as XlsxSheetModel

		if (model.sections.some(section => section.name === "drawing" || section.name === "tableParts")) return locked("drawing")
		if (sectionNames(model, naming)) return locked("reference")
	}

	const lock = await partsLock(document, naming, undefined)

	if (lock !== null) return { refused: lock }

	// Formula records renamed in place (a shared master's children follow it); each one's cell marked.
	const renamed: { sheet: SheetId; id: number; row: number; col: number; before: string; after: string }[] = []

	for (const entry of workbook.sheets) {
		const model = entry.model as XlsxSheetModel
		const store = model.store as SheetStore
		const formulas = store.formulas

		for (const block of store.blocks) {
			if (block.formulas === null) continue

			for (let i = 0; i < block.rowCount; i++) {
				for (let k = block.rowStarts[i] ?? 0; k < (block.rowStarts[i + 1] ?? 0); k++) {
					const id = block.formulas[k] ?? -1

					if (id < 0 || formulas.kind(id) === FormulaKind.SHARED_CHILD) continue

					const text = formulas.text(id)
					const next = renameSheetInFormula(text, from, name)

					if (next !== text)
						renamed.push({
							sheet: entry.id,
							id,
							row: block.firstRow + (block.rows[i] ?? 0),
							col: block.cols[k] ?? 0,
							before: text,
							after: next
						})
				}
			}
		}
	}

	const names = { before: workbook.definedNames.slice(), after: [] as XlsxDefinedName[] }

	for (const defined of names.before) {
		const formula = renameSheetInFormula(defined.formula, from, name)
		const renamedName = formula === defined.formula ? defined : nameWithFormula(defined, formula)

		if (renamedName === null) return locked("definedNames")

		names.after.push(renamedName)
	}

	const namesChanged = names.after.some((defined, i) => defined !== names.before[i])
	const apply = (which: "before" | "after", sink: SheetChangeSink): void => {
		for (const change of renamed) {
			const model = document.model(change.sheet)

			if (model === null) continue

			const formulas = model.store.formulas

			formulas.set(change.id, { ...formulas.formula(change.id), text: which === "after" ? change.after : change.before })
			document.write(change.sheet, change.row, change.col, document.state(change.sheet, change.row, change.col))
			sink.content(change.sheet, { startRow: change.row, startCol: change.col, endRow: change.row, endCol: change.col })
		}

		if (namesChanged) {
			workbook.definedNames.splice(0, workbook.definedNames.length, ...(which === "after" ? names.after : names.before))
			workbook.changes.workbookSection("definedNames")
		}

		target.name = which === "after" ? name : from
		workbook.changes.workbookSection("sheets")

		const engine = document.engine
		const index = document.engineIndex(sheet)

		if (engine !== null && index >= 0) {
			try {
				engine.renameSheet(index, target.name)
			} catch {
				document.dropEngine()
			}
		}

		sink.sheets()
	}

	apply("after", changes)

	return { bytes: 256 + renamed.length * 96, undo: sink => apply("before", sink), redo: sink => apply("after", sink) }
}
