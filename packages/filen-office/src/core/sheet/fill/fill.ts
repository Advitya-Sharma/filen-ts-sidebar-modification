import { SHEET_MAX_COLS, SHEET_MAX_EDIT_CELLS, SHEET_MAX_ROWS, type SheetCellInputs, type SheetRange } from "../protocol"
import { formulaTranslator } from "../refs/formulaRefs"

// Fill Down and Fill Right (Ctrl+D, Ctrl+R in Excel's "Keyboard shortcuts in Excel"), and Fill Up and Left: the
// first row (column) of each range is copied into the rest of it, contents and formats, a formula's relative
// references moving with each copy as a paste moves them. A range one row high (one column wide) takes the row above
// it (the column left of it) as its source, as Excel's Ctrl+D and Ctrl+R do.

declare module "../protocol" {
	interface SheetCommandMap {
		"fill.copy": {
			readonly sheet: SheetId
			readonly ranges: readonly SheetRange[]
			readonly direction: "down" | "right" | "up" | "left"
		}
		// Ctrl+Enter: `entry`, as typed in the active cell `from`, goes into every cell of the ranges but those a merge
		// covers (its first cell takes it), a formula's relative references moving from `from` to each cell, as
		// Excel fills a selection.
		"fill.entry": {
			readonly sheet: SheetId
			readonly ranges: readonly SheetRange[]
			readonly entry: string
			readonly from: { readonly row: number; readonly col: number }
		}
	}
}

export type FillDirection = "down" | "right" | "up" | "left"

// One copy: the target cell and the source cell it takes.
export interface FillPlan {
	readonly rows: Uint32Array
	readonly cols: Uint32Array
	readonly sourceRows: Uint32Array
	readonly sourceCols: Uint32Array
	// The ranges as filled, sources included.
	readonly ranges: readonly SheetRange[]
}

// What a fill writes, or "tooLarge" past the cells one edit takes; null when it writes nothing.
export function fillPlan(ranges: readonly SheetRange[], direction: FillDirection): FillPlan | "tooLarge" | null {
	const vertical = direction === "down" || direction === "up"
	const filled: SheetRange[] = []
	let count = 0

	for (const range of ranges) {
		let { startRow, startCol, endRow, endCol } = range

		if (endRow < startRow || endCol < startCol) continue

		// One line deep: the line before it is the source.
		if (vertical && startRow === endRow) {
			if (direction === "down" ? startRow === 0 : endRow === SHEET_MAX_ROWS - 1) continue
			if (direction === "down") startRow--
			else endRow++
		} else if (!vertical && startCol === endCol) {
			if (direction === "right" ? startCol === 0 : endCol === SHEET_MAX_COLS - 1) continue
			if (direction === "right") startCol--
			else endCol++
		}

		const cells = (endRow - startRow + 1) * (endCol - startCol + 1) - (vertical ? endCol - startCol + 1 : endRow - startRow + 1)

		count += cells

		if (count > SHEET_MAX_EDIT_CELLS) return "tooLarge"

		filled.push({ startRow, startCol, endRow, endCol })
	}

	if (count === 0) return null

	const rows = new Uint32Array(count)
	const cols = new Uint32Array(count)
	const sourceRows = new Uint32Array(count)
	const sourceCols = new Uint32Array(count)
	let at = 0

	for (const range of filled) {
		for (let row = range.startRow; row <= range.endRow; row++) {
			for (let col = range.startCol; col <= range.endCol; col++) {
				const sourceRow = direction === "down" ? range.startRow : direction === "up" ? range.endRow : row
				const sourceCol = direction === "right" ? range.startCol : direction === "left" ? range.endCol : col

				if (sourceRow === row && sourceCol === col) continue

				rows[at] = row
				cols[at] = col
				sourceRows[at] = sourceRow
				sourceCols[at] = sourceCol
				at++
			}
		}
	}

	return { rows, cols, sourceRows, sourceCols, ranges: filled }
}

interface Merge {
	readonly startRow: number
	readonly startCol: number
	readonly endRow: number
	readonly endCol: number
}

// The cells fill.entry writes: `merges` as the sheet holds them (startRow, startCol, endRow, endCol each); "tooLarge"
// past the cells one edit takes, decided before any is listed; null for none.
export function entryCells(
	ranges: readonly SheetRange[],
	merges: Uint32Array,
	entry: string,
	from: { readonly row: number; readonly col: number }
): SheetCellInputs | "tooLarge" | null {
	const within: Merge[][] = []
	let count = 0

	for (const range of ranges) {
		const inside: Merge[] = []
		let cells = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)

		for (let m = 0; m + 3 < merges.length; m += 4) {
			const merge = { startRow: merges[m] ?? 0, startCol: merges[m + 1] ?? 0, endRow: merges[m + 2] ?? 0, endCol: merges[m + 3] ?? 0 }
			const startRow = Math.max(merge.startRow, range.startRow)
			const startCol = Math.max(merge.startCol, range.startCol)
			const endRow = Math.min(merge.endRow, range.endRow)
			const endCol = Math.min(merge.endCol, range.endCol)

			if (startRow > endRow || startCol > endCol) continue

			// Merges never overlap: each covers its part of the range but its first cell.
			cells -=
				(endRow - startRow + 1) * (endCol - startCol + 1) - (startRow === merge.startRow && startCol === merge.startCol ? 1 : 0)
			inside.push(merge)
		}

		count += Math.max(0, cells)
		within.push(inside)
	}

	if (count > SHEET_MAX_EDIT_CELLS) return "tooLarge"
	if (count === 0) return null

	const formula = entry.length > 1 && entry.startsWith("=") ? formulaTranslator(entry.slice(1)) : null
	const rows = new Uint32Array(count)
	const cols = new Uint16Array(count)
	const inputs: string[] = []
	let at = 0

	ranges.forEach((range, index) => {
		const inside = within[index] ?? []

		for (let row = range.startRow; row <= range.endRow; row++) {
			const crossing = inside.length === 0 ? inside : inside.filter(merge => row >= merge.startRow && row <= merge.endRow)

			for (let col = range.startCol; col <= range.endCol; col++) {
				const merge = crossing.length === 0 ? undefined : crossing.find(found => col >= found.startCol && col <= found.endCol)

				// Past the cells the merge covers in this row, keeping its first.
				if (merge !== undefined && (row !== merge.startRow || col !== merge.startCol)) {
					col = merge.endCol

					continue
				}

				rows[at] = row
				cols[at] = col
				inputs.push(formula === null ? entry : `=${formula(row - from.row, col - from.col)}`)
				at++
			}
		}
	})

	return { rows, cols, inputs }
}
