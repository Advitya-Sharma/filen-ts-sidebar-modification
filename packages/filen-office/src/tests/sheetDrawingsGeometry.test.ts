import { describe, expect, it } from "vitest"
import { createGridAxis } from "../react/sheet/axis"
import { anchorBox, anchorOf, draggedBox, placeOnScreen } from "../react/sheet/drawings/geometry"
import { gridView } from "../react/sheet/view"
import type { DrawingAnchor } from "../core/xlsx/drawings/anchor"
import type { SheetAxisLayout, SheetLayout } from "../core/sheet/protocol"

// Drawing objects on the grid: an anchor's box through the axes the grid draws with (zoom, custom and hidden sizes,
// past the extent), the anchor a dragged box gets back, and where a box shows with frozen panes.

function axisLayout(defaultSize: number, entries: readonly (readonly [number, number, number?])[] = []): SheetAxisLayout {
	return {
		defaultSize,
		indices: Uint32Array.from(entries.map(entry => entry[0])),
		sizes: Float32Array.from(entries.map(entry => entry[1])),
		styles: new Uint32Array(entries.length),
		flags: Uint16Array.from(entries.map(entry => entry[2] ?? 0))
	}
}

const LIMITS = { rows: 1_048_576, cols: 16_384 }

function axes(scale = 1) {
	return {
		// Column B is 100 px wide; row 3 is hidden.
		cols: createGridAxis(axisLayout(64, [[1, 100]]), 50, scale),
		rows: createGridAxis(axisLayout(20, [[2, 20, 1]]), 100, scale),
		scale
	}
}

describe("drawing geometry", () => {
	it("puts a two-cell anchor's box on its cells, offsets in EMU, a hidden row taking no height", () => {
		const anchor: DrawingAnchor = {
			mode: "twoCell",
			editAs: "twoCell",
			from: { row: 1, col: 1, rowOff: 95250, colOff: 0 },
			to: { row: 4, col: 2, rowOff: 0, colOff: 9525 * 32 }
		}

		expect(anchorBox(anchor, axes())).toEqual({ x: 64, y: 30, width: 132, height: 30 })
		// At 200 % everything doubles.
		expect(anchorBox(anchor, axes(2))).toEqual({ x: 128, y: 60, width: 264, height: 60 })
	})

	it("ends an offset past its cell at the cell's far edge, and reaches past the axes' extent at the default size", () => {
		const anchor: DrawingAnchor = {
			mode: "oneCell",
			from: { row: 200, col: 0, rowOff: 9525 * 500, colOff: 0 },
			cx: 9525 * 10,
			cy: 9525 * 5
		}

		// Rows 0-99 hold 99 × 20 px (row 2 hidden); row 200 starts 100 default rows on; the offset stops at 20 px.
		expect(anchorBox(anchor, axes())).toEqual({ x: 0, y: 99 * 20 + 100 * 20 + 20, width: 10, height: 5 })
	})

	it("gives back the anchor of the same kind for a moved box, its far edge on a boundary in the next cell", () => {
		const anchor: DrawingAnchor = {
			mode: "twoCell",
			editAs: "oneCell",
			from: { row: 0, col: 0, rowOff: 0, colOff: 0 },
			to: { row: 1, col: 1, rowOff: 0, colOff: 0 }
		}

		expect(anchorOf(anchor, { x: 64, y: 40, width: 100, height: 20 }, axes(), LIMITS)).toEqual({
			mode: "twoCell",
			editAs: "oneCell",
			// y 40 is row 3's top (row 2 is hidden at 40 px).
			from: { row: 3, col: 1, rowOff: 0, colOff: 0 },
			to: { row: 4, col: 2, rowOff: 0, colOff: 0 }
		})
		expect(anchorOf({ mode: "absolute", x: 0, y: 0, cx: 1, cy: 1 }, { x: 10, y: 5, width: 20, height: 10 }, axes(2), LIMITS)).toEqual({
			mode: "absolute",
			x: 9525 * 5,
			y: Math.round(9525 * 2.5),
			cx: 9525 * 10,
			cy: 9525 * 5
		})
	})

	it("keeps a picture's proportions on a corner handle, and moves the far edge on an edge handle", () => {
		const start = { x: 10, y: 10, width: 100, height: 50 }

		expect(draggedBox(start, "se", 50, 0, true)).toEqual({ x: 10, y: 10, width: 150, height: 75 })
		expect(draggedBox(start, "nw", -50, 0, true)).toEqual({ x: -40, y: -15, width: 150, height: 75 })
		expect(draggedBox(start, "e", 30, 99, false)).toEqual({ x: 10, y: 10, width: 130, height: 50 })
		expect(draggedBox(start, null, 5, -5, false)).toEqual({ x: 15, y: 5, width: 100, height: 50 })
		// Never smaller than a few pixels.
		expect(draggedBox(start, "w", 500, 0, false).width).toBe(4)
	})

	it("places a box in the scrolled pane below the frozen rows, and one on a frozen row over them", () => {
		const layout = {
			headings: true,
			pane: { rows: 2, cols: 0, firstRow: 0, firstCol: 0 }
		} as unknown as SheetLayout
		const view = gridView(axes(), layout, { top: 100, left: 0 }, 800, 600)
		const scrolled = placeOnScreen(view, { x: 64, y: 200, width: 10, height: 10 })
		const frozen = placeOnScreen(view, { x: 64, y: 5, width: 10, height: 10 })

		// Headers: 20 px of column header, 32 px of row header; two frozen rows of 20 px.
		expect(scrolled.box.y).toBe(20 + 40 + 200 - 40 - 100)
		expect(scrolled.clip.y).toBe(60)
		expect(frozen.box.y).toBe(25)
		expect(frozen.clip.y).toBe(20)
	})
})
