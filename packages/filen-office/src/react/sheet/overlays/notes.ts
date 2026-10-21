import { CellFlag } from "../../../core/sheet/protocol"
import type { GridDecoration, GridPaintScene, SheetCellDecorator } from "../layers"
import type { GridCanvas } from "../paint/canvas"

// The corner marks Excel draws on annotated cells: a red triangle in the top trailing corner of a cell with a note,
// a purple one for a threaded comment (laid out left to right: the grid's mirroring puts it on the left of a
// right-to-left sheet, as Excel does). A comment's mark wins where a cell has both.

const NOTE_COLOR = "#e00000"
const COMMENT_COLOR = "#7b3fbf"
// The triangle's legs at 100 %, in CSS px.
const SIZE = 6

// A right triangle on the box's top right corner (inside its gridline), from rows of rectangles: GridCanvas has no
// paths.
function corner(canvas: GridCanvas, cell: GridDecoration, scene: GridPaintScene, color: string): void {
	const { box } = cell
	const size = Math.max(4, Math.min(Math.round(SIZE * scene.scale), Math.floor(Math.min(box.width, box.height) / 2)))
	const right = box.x + box.width - 1

	canvas.fillStyle = color

	for (let k = 0; k < size; k++) canvas.fillRect(right - size + k, box.y + k, size - k, 1)
}

export const ANNOTATION_DECORATORS: readonly SheetCellDecorator[] = [
	{
		id: "note",
		mask: CellFlag.NOTE,
		phase: "over",
		paint: (canvas, cell, scene) => {
			if ((cell.flags & CellFlag.THREADED_COMMENT) === 0) corner(canvas, cell, scene, NOTE_COLOR)
		}
	},
	{
		id: "comment",
		mask: CellFlag.THREADED_COMMENT,
		phase: "over",
		paint: (canvas, cell, scene) => corner(canvas, cell, scene, COMMENT_COLOR)
	}
]
