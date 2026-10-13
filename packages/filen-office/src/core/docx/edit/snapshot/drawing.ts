import type { Node } from "prosemirror-model"
import { drawingState, selectedDrawing } from "../../objects/edit"
import type { DocxCrop, DocxWrapMode } from "../commands/drawings"
import type { DocxEditorSession } from "../session"
import type { SnapshotInput } from "./index"

// The object the selection is (a picture, shape, chart or text box), as it is now (its source with what
// commands changed). The same object while the selected drawing node is the same.

export interface DocxDrawingContext {
	readonly kind: "picture" | "shape" | "chart" | "group" | "textbox" | "object"
	readonly anchored: boolean
	readonly wrap: DocxWrapMode
	// EMUs.
	readonly cx: number
	readonly cy: number
	// A floating drawing's offsets (EMUs) from relH / relV; null while it is inline or aligned.
	readonly x: number | null
	readonly y: number | null
	readonly relH: string
	readonly relV: string
	// Degrees, clockwise; mirrored (applied before the rotation).
	readonly rotation: number
	readonly flipH: boolean
	readonly flipV: boolean
	// A picture's crop (drawing.crop), what each side of the picture has cut off (negative: padded).
	readonly crop: DocxCrop
	readonly altText: string
	// What drawing commands can change (drawing.resize; drawing.rotate; drawing.wrap, .position, .altText;
	// drawing.crop).
	readonly resizable: boolean
	readonly rotatable: boolean
	readonly placeable: boolean
	readonly croppable: boolean
	// Resizing keeps the proportions unless asked otherwise (Word's lock for pictures).
	readonly lockAspect: boolean
	// The drawing's position in the story's document.
	readonly pos: number
}

const previous = new WeakMap<
	DocxEditorSession,
	{ readonly node: Node; readonly pos: number; readonly readOnly: boolean; readonly value: DocxDrawingContext }
>()

export function drawingOf(input: SnapshotInput): DocxDrawingContext | null {
	const found = selectedDrawing(input.state)

	if (found === undefined) return null

	const cached = previous.get(input.session)

	if (cached?.node === found.node && cached.pos === found.pos && cached.readOnly === input.readOnly) return cached.value

	const state = drawingState(input.context, found.node)
	const writable = !input.readOnly
	const value: DocxDrawingContext = Object.freeze({
		kind: state.kind,
		anchored: state.wrap !== "inline",
		wrap: state.wrap,
		cx: state.cx,
		cy: state.cy,
		x: state.posX ?? null,
		y: state.posY ?? null,
		relH: state.relH,
		relV: state.relV,
		rotation: state.rot,
		flipH: state.flipH,
		flipV: state.flipV,
		crop: state.crop,
		altText: state.descr,
		resizable: writable && state.resizable,
		rotatable: writable && state.rotatable,
		placeable: writable && state.placeable,
		croppable: writable && state.croppable,
		lockAspect: state.lockAspect,
		pos: found.pos
	})

	previous.set(input.session, { node: found.node, pos: found.pos, readOnly: input.readOnly, value })

	return value
}
