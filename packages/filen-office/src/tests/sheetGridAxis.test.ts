import { describe, expect, it } from "vitest"
import { AxisFlag, SHEET_MAX_ROWS, type SheetAxisLayout, type SheetLayout } from "../core/sheet/protocol"
import { createGridAxis, shownIndex } from "../react/sheet/axis"
import { MAX_SCROLL_PX, domExtent, toDom, toVirtual } from "../react/sheet/scroll"
import { bodyTotal, edgeAt, gridView, hitAxis, revealOffset, screenOf } from "../react/sheet/view"

// The grid's geometry: axes over sparse size exceptions, scaled scrolling past the browsers' element-size caps, and
// the panes of one scroll position.

function axisLayout(defaultSize: number, entries: readonly [index: number, size: number, flags?: number][]): SheetAxisLayout {
	return {
		defaultSize,
		indices: Uint32Array.from(entries, entry => entry[0]),
		sizes: Float32Array.from(entries, entry => entry[1]),
		styles: new Uint32Array(entries.length),
		flags: Uint16Array.from(entries, entry => entry[2] ?? 0)
	}
}

// The same answers by walking every index.
function bruteOffset(sizes: readonly number[], index: number): number {
	let sum = 0

	for (let k = 0; k < index; k++) sum += sizes[k] ?? 0

	return sum
}

function bruteIndexAt(sizes: readonly number[], position: number): number {
	let start = 0

	for (let k = 0; k < sizes.length; k++) {
		const size = sizes[k] ?? 0

		if (position < start + size) return k

		start += size
	}

	return sizes.length - 1
}

function layout(rows: SheetAxisLayout, cols: SheetAxisLayout, pane: SheetLayout["pane"]): SheetLayout {
	return {
		sheet: 0,
		version: 1,
		rowCount: 10,
		colCount: 5,
		rows,
		cols,
		merges: new Uint32Array(0),
		pane,
		zoom: 100,
		gridLines: true,
		headings: true,
		rightToLeft: false,
		protection: null
	}
}

describe("grid axes", () => {
	it("answers offsets, sizes and hit positions as walking every index does, hidden and resized ones included", () => {
		let seed = 7
		const random = (): number => {
			seed = (seed * 1103515245 + 12345) % 2 ** 31

			return seed / 2 ** 31
		}

		for (let round = 0; round < 40; round++) {
			const count = 1 + Math.floor(random() * 300)
			const entries: [number, number, number][] = []

			for (let index = 0; index < count; index++) {
				if (random() < 0.15) entries.push([index, Math.round(random() * 60), random() < 0.3 ? AxisFlag.HIDDEN : 0])
			}

			const axis = createGridAxis(axisLayout(20, entries), count, 1)
			const sizes = Array.from({ length: count }, (_, index) => {
				const entry = entries.find(item => item[0] === index)

				return entry === undefined ? 20 : entry[2] === AxisFlag.HIDDEN ? 0 : entry[1]
			})

			expect(axis.total).toBe(bruteOffset(sizes, count))

			for (let index = 0; index <= count; index++) expect(axis.offset(index)).toBe(bruteOffset(sizes, index))
			for (let index = 0; index < count; index++) expect(axis.size(index)).toBe(sizes[index])
			for (let position = 0; position < axis.total; position += 3) expect(axis.indexAt(position)).toBe(bruteIndexAt(sizes, position))
		}
	})

	it("scales sizes by the zoom and takes a drag's overrides (NaN back to the default)", () => {
		const base = axisLayout(64, [[2, 100]])
		const zoomed = createGridAxis(base, 10, 1.5)

		expect(zoomed.size(0)).toBe(96)
		expect(zoomed.size(2)).toBe(150)

		const dragged = createGridAxis(
			base,
			10,
			1,
			new Map([
				[2, Number.NaN],
				[5, 30]
			])
		)

		expect(dragged.size(2)).toBe(64)
		expect(dragged.size(5)).toBe(30)
		expect(dragged.offset(6)).toBe(5 * 64 + 30)
	})

	it("holds a million rows in memory for its exceptions alone and finds the last row", () => {
		const axis = createGridAxis(axisLayout(20, [[5, 0, AxisFlag.HIDDEN]]), SHEET_MAX_ROWS, 1)

		expect(axis.total).toBe((SHEET_MAX_ROWS - 1) * 20)
		expect(axis.indexAt(axis.total - 1)).toBe(SHEET_MAX_ROWS - 1)
		expect(axis.hidden(5)).toBe(true)
		expect(shownIndex(axis, 5, 1)).toBe(6)
		expect(shownIndex(axis, 5, -1)).toBe(4)
	})
})

describe("scaled scrolling", () => {
	it("keeps the scrolled element under the browsers' caps and reaches both ends", () => {
		const virtual = SHEET_MAX_ROWS * 20 * 4
		const view = 800

		expect(domExtent(virtual)).toBe(MAX_SCROLL_PX)
		expect(toVirtual(0, virtual, view)).toBe(0)
		expect(toVirtual(MAX_SCROLL_PX - view, virtual, view)).toBe(virtual - view)

		for (const offset of [0, 1234567, virtual / 2, virtual - view]) {
			const back = toVirtual(toDom(offset, virtual, view), virtual, view)

			// One browser pixel moves at most the scale factor's pixels of sheet.
			expect(Math.abs(back - offset)).toBeLessThanOrEqual(Math.ceil(virtual / MAX_SCROLL_PX))
		}
	})

	it("maps one to one under the cap", () => {
		expect(domExtent(5000)).toBe(5000)
		expect(toVirtual(1234, 5000, 600)).toBe(1234)
		expect(toDom(1234, 5000, 600)).toBe(1234)
		expect(toVirtual(9999, 5000, 600)).toBe(4400)
	})
})

describe("grid view", () => {
	const rows = axisLayout(20, [])
	const cols = axisLayout(64, [[1, 0, AxisFlag.HIDDEN]])

	it("places frozen rows and columns before the scrolled ones and hit-tests each part", () => {
		const sheet = layout(rows, cols, { rows: 2, cols: 1, topRow: 2, leftCol: 1 })
		const axes = { rows: createGridAxis(rows, 200, 1), cols: createGridAxis(cols, 40, 1), scale: 1 }
		const view = gridView(axes, sheet, { top: 100, left: 0 }, 600, 400)

		expect(view.rows.header).toBe(20)
		expect(view.rows.frozen).toBe(2)
		expect(view.rows.first).toBe(7)
		expect(screenOf(view.rows, 0)).toBe(20)
		expect(screenOf(view.rows, 7)).toBe(20 + 40)
		expect(hitAxis(view.rows, 10)).toEqual({ part: "header" })
		expect(hitAxis(view.rows, 25)).toEqual({ part: "frozen", index: 0 })
		expect(hitAxis(view.rows, 61)).toEqual({ part: "body", index: 7 })
		expect(bodyTotal(view.rows)).toBe(200 * 20 - 40)
		// Column 1 is hidden: the scrolled part starts at column 2.
		expect(view.cols.first).toBe(2)
	})

	it("scrolls the least that reveals a cell and leaves frozen ones alone", () => {
		const sheet = layout(rows, cols, { rows: 2, cols: 0, topRow: 2, leftCol: 0 })
		const axes = { rows: createGridAxis(rows, 200, 1), cols: createGridAxis(cols, 40, 1), scale: 1 }
		const view = gridView(axes, sheet, { top: 0, left: 0 }, 600, 400)

		expect(revealOffset(view.rows, 1)).toBe(0)
		expect(revealOffset(view.rows, 5)).toBe(0)
		// The body shows 400 - 20 - 40 = 340 px: row 30 ends at 31 * 20 - 40.
		expect(revealOffset(view.rows, 30)).toBe(31 * 20 - 40 - 340)
	})

	it("finds the header edge a pointer is on, skipping hidden columns", () => {
		const sheet = layout(rows, cols, null)
		const axes = { rows: createGridAxis(rows, 200, 1), cols: createGridAxis(cols, 40, 1), scale: 1 }
		const view = gridView(axes, sheet, { top: 0, left: 0 }, 600, 400)
		const header = view.cols.header

		// Column 0 ends at header + 64; column 1 is hidden, so the edge before column 2 is column 0's.
		expect(edgeAt(view.cols, header + 63, 4)).toBe(0)
		expect(edgeAt(view.cols, header + 66, 4)).toBe(0)
		expect(edgeAt(view.cols, header + 100, 4)).toBeNull()
	})
})
