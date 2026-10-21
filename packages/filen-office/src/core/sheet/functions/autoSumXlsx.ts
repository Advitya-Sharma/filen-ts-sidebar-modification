import type { SheetStore } from "../../xlsx/model/store"
import { CellKind, SLOT_BLOCK } from "../../xlsx/model/types"
import type { XlsxDocument } from "../../xlsx/edit/document"
import { defineSheetFeature } from "../registry"
import { autoSum, type AutoSumCells } from "./autoSum"

// AutoSum over an xlsx sheet's store.

function cellsOf(store: SheetStore): AutoSumCells {
	const at = (row: number, col: number): { kind: number; formula: boolean } | null => {
		const slot = store.locate(row, col)

		if (slot < 0) return null

		const block = store.blocks[Math.floor(slot / SLOT_BLOCK)]
		const k = slot % SLOT_BLOCK

		return block === undefined ? null : { kind: block.kinds[k] ?? CellKind.EMPTY, formula: (block.formulas?.[k] ?? -1) >= 0 }
	}

	return {
		number: (row, col) => at(row, col)?.kind === CellKind.NUMBER,
		filled: (row, col) => {
			const cell = at(row, col)

			return cell !== null && (cell.kind !== CellKind.EMPTY || cell.formula)
		}
	}
}

export const xlsxAutoSumFeature = defineSheetFeature<XlsxDocument>({
	id: "autoSum",
	queries: {
		autoSum: {
			run: async (context, request) => {
				const model = await context.doc.ready(request.sheet, context.slicing, context.signal)

				return autoSum(cellsOf(model.store as SheetStore), request.range, request.active)
			}
		}
	}
})
