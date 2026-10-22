import { describe, expect, it } from "vitest"
import { gridMove, type GridBounds } from "../core/sheet/index"

describe("gridMove over hidden rows, columns and merges", () => {
	const at = (row: number, col: number) => ({ anchor: { row, col }, focus: { row, col } })
	const key = (name: string, extra: Partial<{ shiftKey: boolean; ctrlKey: boolean; metaKey: boolean; altKey: boolean }> = {}) => ({
		key: name,
		shiftKey: false,
		ctrlKey: false,
		metaKey: false,
		altKey: false,
		...extra
	})
	// Rows 2 and 3 and columns 0 and 4 hidden, B6:D7 merged.
	const merge = { startRow: 5, startCol: 1, endRow: 6, endCol: 3 }
	const bounds: GridBounds = {
		rowCount: 110,
		colCount: 26,
		lastUsedRow: 9,
		lastUsedCol: 4,
		pageRows: 3,
		hiddenRow: row => row === 2 || row === 3,
		hiddenCol: col => col === 0 || col === 4,
		mergeAt: (row, col) =>
			row >= merge.startRow && row <= merge.endRow && col >= merge.startCol && col <= merge.endCol ? merge : undefined
	}

	it("passes over hidden rows and columns", () => {
		expect(gridMove(key("ArrowDown"), at(1, 1), bounds)).toEqual(at(4, 1))
		expect(gridMove(key("ArrowUp"), at(4, 1), bounds)).toEqual(at(1, 1))
		expect(gridMove(key("ArrowRight"), at(0, 3), bounds)).toEqual(at(0, 5))
		// Nothing shown to the left: stays.
		expect(gridMove(key("ArrowLeft"), at(0, 1), bounds)).toEqual(at(0, 1))
		expect(gridMove(key("Home"), at(0, 3), bounds)).toEqual(at(0, 1))
	})

	it("lands on a merge's anchor, and leaves it from its far edge", () => {
		expect(gridMove(key("ArrowDown"), at(4, 2), bounds)).toEqual(at(5, 1))
		expect(gridMove(key("ArrowLeft"), at(6, 5), bounds)).toEqual(at(5, 1))
		expect(gridMove(key("ArrowRight"), at(5, 1), bounds)).toEqual(at(5, 5))
		expect(gridMove(key("ArrowDown"), at(5, 1), bounds)).toEqual(at(7, 1))
		expect(gridMove(key("ArrowDown", { shiftKey: true }), at(4, 1), bounds)).toEqual({
			anchor: { row: 4, col: 1 },
			focus: { row: 5, col: 1 }
		})
	})
})
