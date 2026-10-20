import { recalculateLogged, setCells, splitsArray, tableHeaders, unshare } from "../../xlsx/edit/cells"
import type { XlsxDocument } from "../../xlsx/edit/document"
import { cellStep, reportLogged } from "../../xlsx/edit/format"
import { ABSENT, CellLog, type CellState } from "../../xlsx/edit/history"
import type { SheetStore } from "../../xlsx/model/store"
import { CellKind, FormulaKind, type XlsxSheetModel } from "../../xlsx/model/types"
import { defineSheetFeature, type SheetChangeSink, type SheetHandled } from "../registry"
import { formulaTranslator } from "../refs/formulaRefs"
import type { SheetId, SheetRange } from "../protocol"
import { entryCells, fillPlan, type FillDirection, type FillPlan } from "./fill"

// Fill over an xlsx sheet: each target takes its source cell as it is (its value of whatever type, its format), a
// formula moved by the target's offset, so a copy is exact where a typed entry would be read again. Refused where a
// paste would be: locked cells, part of an array formula, a table header's text; and where merged cells are in the
// way, or the source holds an array formula or a dynamic array (whose range a copy cannot carry).

function intersects(a: SheetRange, startRow: number, startCol: number, endRow: number, endCol: number): boolean {
	return a.startRow <= endRow && startRow <= a.endRow && a.startCol <= endCol && startCol <= a.endCol
}

// Whether a source holds a formula a copy cannot carry: an array formula's, a data table's or a dynamic array's.
function arraySource(model: XlsxSheetModel, formula: number): boolean {
	if (formula < 0) return false

	const kind = model.store.formulas.kind(formula)

	return kind === FormulaKind.ARRAY || kind === FormulaKind.DATA_TABLE || model.store.formulas.formula(formula).cellMetadata >= 0
}

// The state a target takes from its source.
function copied(document: XlsxDocument, model: XlsxSheetModel, plan: FillPlan, i: number, sheet: SheetId): CellState {
	const row = plan.rows[i] ?? 0
	const col = plan.cols[i] ?? 0
	const sourceRow = plan.sourceRows[i] ?? 0
	const sourceCol = plan.sourceCols[i] ?? 0
	const source = document.state(sheet, sourceRow, sourceCol)

	if (!source.present) {
		// An empty source still gives its format (its row's or column's).
		const style = document.startingStyle(model, sourceRow, sourceCol)

		return style === 0 ? ABSENT : { present: true, kind: CellKind.EMPTY, value: 0, style, formula: -1, shown: undefined }
	}

	if (source.formula < 0)
		return { present: true, kind: source.kind, value: source.value, style: source.style, formula: -1, shown: undefined }

	const formulas = model.store.formulas
	const text = formulaTranslator(document.formulaText(model, source.formula, sourceRow, sourceCol))(row - sourceRow, col - sourceCol)
	const formula = formulas.add({ kind: FormulaKind.NORMAL, text, sharedIndex: -1, ref: null, flags: 0, cellMetadata: -1 })

	return { present: true, kind: CellKind.EMPTY, value: 0, style: source.style, formula, shown: undefined }
}

export async function fillCells(
	document: XlsxDocument,
	sheet: SheetId,
	ranges: readonly SheetRange[],
	direction: FillDirection,
	changes: SheetChangeSink
): Promise<SheetHandled> {
	const model = await document.ready(sheet)
	const plan = fillPlan(ranges, direction)

	if (plan === null) return null
	if (plan === "tooLarge") return { refused: { reason: "tooLarge" } }

	const count = plan.rows.length

	for (let i = 0; i < count; i++)
		if (document.locked(sheet, plan.rows[i] ?? 0, plan.cols[i] ?? 0)) return { refused: { reason: "protectedSheet" } }

	for (let m = 0; m + 3 < model.merges.length; m += 4) {
		const merge = {
			startRow: model.merges[m] ?? 0,
			startCol: model.merges[m + 1] ?? 0,
			endRow: model.merges[m + 2] ?? 0,
			endCol: model.merges[m + 3] ?? 0
		}

		if (plan.ranges.some(range => intersects(merge, range.startRow, range.startCol, range.endRow, range.endCol)))
			return { refused: { reason: "merged" } }
	}

	if (splitsArray(model, plan.rows, plan.cols)) return { refused: { reason: "arrayFormula" } }

	if (model.relationships.all.some(relationship => relationship.type.endsWith("/table"))) {
		const headers = await tableHeaders(document, model)

		for (let i = 0; i < count; i++) {
			const row = plan.rows[i] ?? 0
			const col = plan.cols[i] ?? 0

			if (
				headers.some(range => intersects(range, row, col, row, col)) &&
				document.cellView(sheet, row, col).text !== document.cellView(sheet, plan.sourceRows[i] ?? 0, plan.sourceCols[i] ?? 0).text
			)
				return { refused: { reason: "tableHeader" } }
		}
	}

	const store = model.store as SheetStore
	let formulas = false

	for (let i = 0; i < count; i++) {
		const formula = store.formula(store.locate(plan.sourceRows[i] ?? 0, plan.sourceCols[i] ?? 0))

		if (arraySource(model, formula)) return { refused: { reason: "arrayFormula" } }

		formulas ||= formula >= 0
	}

	// Built from the cells as they stand before this edit, so the edit is what it recalculates.
	if (document.engine === null && document.allReady && (document.hasFormulas || formulas)) {
		await document.loadTables()
		document.buildEngine(changes)
	}

	// Every source is read before any target is written: a target never feeds another.
	const states: CellState[] = []

	for (let i = 0; i < count; i++) states.push(copied(document, model, plan, i, sheet))

	const log = new CellLog()

	for (let i = 0; i < count; i++) {
		const row = plan.rows[i] ?? 0
		const col = plan.cols[i] ?? 0
		const before = document.state(sheet, row, col)

		if (before.formula >= 0 && store.formulas.kind(before.formula) === FormulaKind.SHARED_MASTER)
			unshare(document, sheet, row, col, log)

		log.touch(sheet, row, col, () => document.state(sheet, row, col), true)
		document.write(sheet, row, col, states[i] as CellState)
	}

	recalculateLogged(document, log, changes)
	log.seal((s, r, c) => document.state(s, r, c))
	reportLogged(log, changes, "content")

	return cellStep(document, log, sink => reportLogged(log, sink, "content"), true)
}

export const xlsxFillFeature = defineSheetFeature<XlsxDocument>({
	id: "fill",
	commands: {
		"fill.copy": {
			info: { write: true },
			run: (context, command) => fillCells(context.doc, command.sheet, command.ranges, command.direction, context.changes)
		},
		"fill.entry": {
			info: { write: true },
			run: async (context, command) => {
				const model = await context.doc.ready(command.sheet)
				const cells = entryCells(command.ranges, model.merges, command.entry, command.from)

				if (cells === null) return null
				if (cells === "tooLarge") return { refused: { reason: "tooLarge" } }

				return setCells(context.doc, command.sheet, cells, context.changes)
			}
		}
	}
})
