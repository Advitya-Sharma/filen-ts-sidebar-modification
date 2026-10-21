import { describe, expect, it } from "vitest"
import { CellFlag } from "../core/sheet/protocol"
import type { GridDecoration, GridPaintScene } from "../react/sheet/layers"
import { ANNOTATION_DECORATORS } from "../react/sheet/overlays/notes"
import type { GridCanvas } from "../react/sheet/paint/canvas"

// The corner marks of notes and threaded comments: a triangle in the cell's top trailing corner, inside its gridline,
// red for a note, purple for a comment (which wins on a cell with both), scaled with the zoom.

function recorder(): { canvas: GridCanvas; rects: { color: string; x: number; y: number; width: number; height: number }[] } {
	const rects: { color: string; x: number; y: number; width: number; height: number }[] = []
	const canvas = {
		fillStyle: "" as string | object,
		fillRect(x: number, y: number, width: number, height: number) {
			rects.push({ color: String(canvas.fillStyle), x, y, width, height })
		}
	}

	return { canvas: canvas as unknown as GridCanvas, rects }
}

function cell(flags: number): GridDecoration {
	return { row: 1, col: 1, flags, tile: {} as GridDecoration["tile"], entry: 0, box: { x: 100, y: 40, width: 64, height: 20 } }
}

function paint(flags: number, scale = 1): { color: string; x: number; y: number; width: number; height: number }[] {
	const { canvas, rects } = recorder()
	const scene = { scale } as GridPaintScene

	for (const decorator of ANNOTATION_DECORATORS) if ((flags & decorator.mask) !== 0) decorator.paint(canvas, cell(flags), scene)

	return rects
}

describe("note and comment corner marks", () => {
	it("draws a note's red triangle in the top right corner, inside the gridline", () => {
		const rects = paint(CellFlag.NOTE)

		expect(rects).toHaveLength(6)
		expect(new Set(rects.map(rect => rect.color))).toEqual(new Set(["#e00000"]))
		// Rows of the triangle: the widest at the top, each ending at the cell's last pixel inside its gridline.
		expect(rects[0]).toMatchObject({ x: 157, y: 40, width: 6, height: 1 })
		expect(rects[5]).toMatchObject({ x: 162, y: 45, width: 1, height: 1 })
	})

	it("draws a comment's purple mark instead of a note's on a cell with both", () => {
		const rects = paint(CellFlag.NOTE | CellFlag.THREADED_COMMENT)

		expect(new Set(rects.map(rect => rect.color))).toEqual(new Set(["#7b3fbf"]))
	})

	it("scales with the zoom, never past half the cell", () => {
		expect(paint(CellFlag.NOTE, 2)).toHaveLength(10)
		expect(paint(CellFlag.NOTE, 0.25)).toHaveLength(4)
	})
})
