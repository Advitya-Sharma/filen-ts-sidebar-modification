import type { SheetId } from "@filen/office/core/sheet"

// Where a find stands in its matches (triples sheet, row, col, in the order the session walked them: the sheets in
// tab order, each row by row).

export interface SheetFindCell {
	readonly sheet: SheetId
	readonly row: number
	readonly col: number
}

// The most matches one find lists ("1,000+" past it).
export const SHEET_FIND_LIMIT = 1_000

export function matchAt(cells: Uint32Array, index: number): SheetFindCell | null {
	const at = index * 3

	return at + 2 < cells.length ? { sheet: cells[at] ?? 0, row: cells[at + 1] ?? 0, col: cells[at + 2] ?? 0 } : null
}

// Whether `a` comes before `b` in the walk, given the sheets' order.
function before(a: SheetFindCell, b: SheetFindCell, order: readonly SheetId[]): boolean {
	const sheetA = order.indexOf(a.sheet)
	const sheetB = order.indexOf(b.sheet)

	if (sheetA !== sheetB) {
		return sheetA < sheetB
	}

	return a.row !== b.row ? a.row < b.row : a.col < b.col
}

// Matches found from a cell on (the walk once round from it) put in walk order from the start: the list as a find
// from the first cell lists it, and `shift`, the index in `cells` of the first match in walk order (where the walk
// came round), so a match at `index` of `cells` is at (index - shift + count) % count of the list.
export function walkOrder(cells: Uint32Array, order: readonly SheetId[]): { readonly cells: Uint32Array; readonly shift: number } {
	const count = Math.floor(cells.length / 3)

	for (let index = 1; index < count; index++) {
		const cell = matchAt(cells, index)
		const previous = matchAt(cells, index - 1)

		if (cell !== null && previous !== null && before(cell, previous, order)) {
			const rotated = new Uint32Array(count * 3)

			rotated.set(cells.subarray(index * 3, count * 3), 0)
			rotated.set(cells.subarray(0, index * 3), (count - index) * 3)

			return { cells: rotated, shift: index }
		}
	}

	return { cells, shift: 0 }
}

// The match nearest `here` the way a step goes: the first after it in walk order (direction 1) or the last before it
// (-1), coming round past either end; -1 for none.
export function matchNear(cells: Uint32Array, order: readonly SheetId[], here: SheetFindCell, direction: 1 | -1): number {
	const count = Math.floor(cells.length / 3)

	for (let step = 0; step < count; step++) {
		const index = direction > 0 ? step : count - 1 - step
		const cell = matchAt(cells, index)

		if (cell !== null && (direction > 0 ? before(here, cell, order) : before(cell, here, order))) {
			return index
		}
	}

	return count === 0 ? -1 : direction > 0 ? 0 : count - 1
}
