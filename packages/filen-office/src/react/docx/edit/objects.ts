import type { DocxDrawingContext } from "../../../core/docx/edit/snapshot/drawing"
import type { DocxEditorCommand } from "../../../core/docx/edit/index"
import { DRAWING_HANDLES, HANDLE_PLACES, type DrawingHandle, type PageBox } from "../../../core/docx/objects/hits"

// Moving and resizing a drawing with a pointer, headless: the selection frame's handles, the box a drag gives
// (shown meanwhile as a CSS transform of the painted object, so nothing is laid out until the pointer is
// released) and the one command the drag ends with. Page px throughout (a view divides pointer deltas by its
// zoom); EMUs only in the commands. The frame and handles are drawn by the app (data-fo-handle, --fo-*
// variables in docx-objects.css).

export { DRAWING_HANDLES, type DrawingHandle, type PageBox }

// Layout px are 1/96 inch.
export const EMU_PER_PAGE_PX = 914400 / 96
// The smallest a drag makes a drawing (page px).
export const MIN_DRAWING_PX = 4

// Where each handle of a box is (page px), for drawing the frame.
export function handlePoints(box: PageBox): { readonly handle: DrawingHandle; readonly x: number; readonly y: number }[] {
	return DRAWING_HANDLES.map(handle => {
		const [fx, fy] = HANDLE_PLACES[handle]

		return { handle, x: box.x + box.width * fx, y: box.y + box.height * fy }
	})
}

// The CSS cursor a handle shows.
export function handleCursor(handle: DrawingHandle): string {
	return `${handle}-resize`
}

// The box after dragging `handle` of `box` by (dx, dy): the opposite edge or corner stays; with `keepRatio`
// (a corner dragged with Shift, or a picture whose proportions are locked) the box keeps its proportions,
// following the larger of the two moves.
export function resizedBox(box: PageBox, handle: DrawingHandle, dx: number, dy: number, keepRatio: boolean): PageBox {
	const [fx, fy] = HANDLE_PLACES[handle]
	const horizontal = fx !== 0.5
	const vertical = fy !== 0.5
	let width = horizontal ? box.width + (fx === 1 ? dx : -dx) : box.width
	let height = vertical ? box.height + (fy === 1 ? dy : -dy) : box.height

	width = Math.max(MIN_DRAWING_PX, width)
	height = Math.max(MIN_DRAWING_PX, height)

	if (keepRatio && box.width > 0 && box.height > 0) {
		const scale =
			horizontal && vertical ? Math.max(width / box.width, height / box.height) : horizontal ? width / box.width : height / box.height

		width = Math.max(MIN_DRAWING_PX, box.width * scale)
		height = Math.max(MIN_DRAWING_PX, box.height * scale)
	}

	// The handle moves; the edges it does not hold stay (a side handle keeps the box centred across).
	const x = fx === 0 ? box.x + box.width - width : fx === 1 ? box.x : box.x + (box.width - width) / 2
	const y = fy === 0 ? box.y + box.height - height : fy === 1 ? box.y : box.y + (box.height - height) / 2

	return { x, y, width, height }
}

// The transform that shows the object painted in `from` as `to` (transform-origin 0 0 on the object).
export function previewTransform(from: PageBox, to: PageBox): string {
	const sx = from.width > 0 ? to.width / from.width : 1
	const sy = from.height > 0 ? to.height / from.height : 1

	return `translate(${to.x - from.x}px, ${to.y - from.y}px) scale(${sx}, ${sy})`
}

// The resize a box drag ends with: the drawing's extent scaled as its box was (the box may include effects).
export function resizeCommand(drawing: DocxDrawingContext, from: PageBox, to: PageBox): DocxEditorCommand | undefined {
	if (!drawing.resizable || from.width <= 0 || from.height <= 0) return undefined

	const cx = Math.max(1, Math.round((drawing.cx * to.width) / from.width))
	const cy = Math.max(1, Math.round((drawing.cy * to.height) / from.height))

	return cx === drawing.cx && cy === drawing.cy ? undefined : { type: "drawing.resize", cx, cy }
}

// The move a drag of a floating drawing by (dx, dy) page px ends with: its offsets from what it is placed
// against, moved as far.
export function moveCommand(drawing: DocxDrawingContext, dx: number, dy: number): DocxEditorCommand | undefined {
	if (!drawing.placeable || !drawing.anchored || (dx === 0 && dy === 0)) return undefined

	return {
		type: "drawing.position",
		x: Math.round((drawing.x ?? 0) + dx * EMU_PER_PAGE_PX),
		y: Math.round((drawing.y ?? 0) + dy * EMU_PER_PAGE_PX)
	}
}

// One pointer drag on a selected drawing: from the press (a handle, or the body of a floating drawing) to the
// release. move() gives the box to show; end() the command to run (undefined: nothing changed, or the drag
// cannot change the drawing).
export class DrawingDrag {
	readonly drawing: DocxDrawingContext
	readonly from: PageBox
	readonly handle: DrawingHandle | undefined
	private box: PageBox
	private dx = 0
	private dy = 0

	constructor(drawing: DocxDrawingContext, from: PageBox, handle: DrawingHandle | undefined) {
		this.drawing = drawing
		this.from = from
		this.handle = handle
		this.box = from
	}

	// Whether the drag can do anything: a handle of a resizable drawing, or a floating one moved.
	get active(): boolean {
		return this.handle === undefined ? this.drawing.placeable && this.drawing.anchored : this.drawing.resizable
	}

	// The pointer moved by (dx, dy) page px from the press; `free`: proportions not kept even where they are
	// locked (Word's Shift inverts the lock at corners).
	move(dx: number, dy: number, free = false): PageBox {
		this.dx = dx
		this.dy = dy

		if (!this.active) return this.box

		if (this.handle === undefined) {
			this.box = { ...this.from, x: this.from.x + dx, y: this.from.y + dy }
		} else {
			const corner = HANDLE_PLACES[this.handle][0] !== 0.5 && HANDLE_PLACES[this.handle][1] !== 0.5
			const keep = corner && this.drawing.lockAspect !== free

			this.box = resizedBox(this.from, this.handle, dx, dy, keep)
		}

		return this.box
	}

	get transform(): string {
		return previewTransform(this.from, this.box)
	}

	end(): DocxEditorCommand | undefined {
		if (!this.active) return undefined

		return this.handle === undefined ? moveCommand(this.drawing, this.dx, this.dy) : resizeCommand(this.drawing, this.from, this.box)
	}
}
