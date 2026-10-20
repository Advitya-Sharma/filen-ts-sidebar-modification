import type { SheetStore } from "../../xlsx/model/store"
import { CellFlag, type SheetId, type SheetRange } from "../protocol"
import { setCells } from "../../xlsx/edit/cells"
import type { XlsxDocument } from "../../xlsx/edit/document"
import { CellKind, SLOT_BLOCK, type XlsxSheetModel } from "../../xlsx/model/types"
import { defineSheetFeature, type SheetReadContext } from "../registry"
import { findCells, replaceCells, type FindDocument, type FindSheet } from "./find"

// Find and Replace over an xlsx session's cells.

class XlsxFindSheet implements FindSheet {
	private readonly document: XlsxDocument
	private readonly sheet: SheetId
	private readonly model: XlsxSheetModel
	private readonly store: SheetStore
	// The cell the scan stands on: read by its slot, with no lookup.
	private row = -1
	private col = -1
	private slot = -1
	// The last cell read: a match test and a replace read one cell twice.
	private readRow = -1
	private readCol = -1
	private view: { readonly text: string; readonly input: string | null; readonly flags: number } = { text: "", input: null, flags: 0 }

	constructor(document: XlsxDocument, sheet: SheetId, model: XlsxSheetModel) {
		this.document = document
		this.sheet = sheet
		this.model = model
		this.store = model.store as SheetStore
	}

	scan(range: SheetRange, visit: (row: number, col: number) => boolean): boolean {
		const blocks = this.store.blocks

		for (let b = Math.max(0, this.store.blockOf(range.startRow)); b < blocks.length; b++) {
			const block = blocks[b]

			if (block === undefined || block.firstRow > range.endRow) break

			const formulas = block.formulas

			for (let i = 0; i < block.rowCount; i++) {
				const row = block.firstRow + (block.rows[i] ?? 0)

				if (row < range.startRow) continue
				if (row > range.endRow) break

				const end = block.rowStarts[i + 1] ?? 0

				for (let k = block.rowStarts[i] ?? 0; k < end; k++) {
					const col = block.cols[k] ?? 0

					if (col < range.startCol) continue
					if (col > range.endCol) break
					// A cell kept only for its format holds nothing to find.
					if (block.kinds[k] === CellKind.EMPTY && (formulas?.[k] ?? -1) < 0) continue

					this.row = row
					this.col = col
					this.slot = b * SLOT_BLOCK + k

					if (visit(row, col)) return true
				}
			}
		}

		return false
	}

	private read(row: number, col: number): { readonly text: string; readonly input: string | null; readonly flags: number } {
		if (row !== this.readRow || col !== this.readCol) {
			this.view =
				row === this.row && col === this.col
					? this.document.shownAt(this.model, row, col, this.slot)
					: this.document.cellView(this.sheet, row, col)
			this.readRow = row
			this.readCol = col
		}

		return this.view
	}

	shown(row: number, col: number): string {
		return this.read(row, col).text
	}

	typed(row: number, col: number): string | null {
		const view = this.read(row, col)

		if ((view.flags & CellFlag.FORMULA_HIDDEN) !== 0) return null

		return view.input ?? view.text
	}
}

function findDocument(document: XlsxDocument): FindDocument {
	return {
		searched: () =>
			document
				.worksheets()
				.filter(entry => entry.state === "visible")
				.map(entry => entry.id),
		sheet: async (sheet: SheetId, context: SheetReadContext<unknown>) => {
			const entry = document.workbook.sheet(sheet)

			if (entry === undefined || entry.kind !== "worksheet") return null

			const model = await document.ready(sheet, context.slicing, context.signal)

			return new XlsxFindSheet(document, sheet, model)
		},
		write: (sheet, cells, changes) => setCells(document, sheet, cells, changes)
	}
}

export const xlsxFindFeature = defineSheetFeature<XlsxDocument>({
	id: "find",
	commands: {
		"find.replace": {
			info: { write: true },
			run: (context, command) => replaceCells(findDocument(context.doc), command, context)
		}
	},
	queries: {
		find: { run: (context, request) => findCells(findDocument(context.doc), request, context) }
	}
})
