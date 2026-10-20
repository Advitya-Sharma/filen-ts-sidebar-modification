import type { SheetRange } from "../../sheet/protocol"
import { sameSheet, type FormulaRef } from "../../sheet/refs/formulaRefs"

// A reference after cells are cut and pasted (a move):
// - a cell or range wholly inside the cut range follows its cells (support.microsoft.com "Paste options": cell
//   references are adjusted when formulas are cut and pasted);
// - one wholly to cells the paste writes over (and the cut does not take) becomes #REF! (support.microsoft.com
//   "How to correct a #REF! error": a formula refers to a cell that is not valid once it is pasted over);
// - anything else stays as written: a range only partly over either, whole rows and columns, 3-D spans
//   (LibreOffice's moveRange, which Microsoft's documents leave this to).

export interface CellMove {
	// The cut range's sheet, and the paste's.
	readonly sheet: string
	readonly range: SheetRange
	readonly to: string
	readonly destination: SheetRange
}

// A changed reference, null for #REF!, undefined to keep it.
export type MovedRef = FormulaRef | null | undefined

export function intersects(a: SheetRange, b: SheetRange): boolean {
	return a.startRow <= b.endRow && b.startRow <= a.endRow && a.startCol <= b.endCol && b.startCol <= a.endCol
}

export function contains(outer: SheetRange, inner: SheetRange): boolean {
	return (
		outer.startRow <= inner.startRow && outer.startCol <= inner.startCol && outer.endRow >= inner.endRow && outer.endCol >= inner.endCol
	)
}

// `own`: the sheet of the formula (unqualified references name it), null for one without.
export function moveReference(ref: FormulaRef, own: string | null, move: CellMove): MovedRef {
	const area = ref.area

	if (ref.external || ref.lastSheet !== null || area === null || area.kind !== "cell") return undefined

	const sheet = ref.sheet ?? own

	if (sheet === null) return undefined

	const box = {
		startRow: Math.min(area.startRow, area.endRow),
		startCol: Math.min(area.startCol, area.endCol),
		endRow: Math.max(area.startRow, area.endRow),
		endCol: Math.max(area.startCol, area.endCol)
	}
	const onSource = sameSheet(sheet, move.sheet)

	if (onSource && contains(move.range, box)) {
		const rows = move.destination.startRow - move.range.startRow
		const cols = move.destination.startCol - move.range.startCol

		return {
			...ref,
			sheet: ref.sheet === null && own !== null && sameSheet(own, move.to) ? null : move.to,
			area: {
				...area,
				startRow: area.startRow + rows,
				endRow: area.endRow + rows,
				startCol: area.startCol + cols,
				endCol: area.endCol + cols
			}
		}
	}

	if (sameSheet(sheet, move.to) && contains(move.destination, box) && !(onSource && intersects(move.range, box))) return null

	return undefined
}
