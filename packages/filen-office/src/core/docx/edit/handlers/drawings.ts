import type { Node } from "prosemirror-model"
import {
	canInsertPicture,
	changeDrawing,
	croppedExtent,
	drawingState,
	floatingPosition,
	insertPicture,
	pictureRefusal,
	resizedExtent,
	sameCrop,
	selectedDrawing,
	settleMedia,
	sourceState,
	type DrawingChanges,
	type DrawingState
} from "../../objects/edit"
import type { DocumentCommandContext, DocumentHandlers, FeatureHandlers } from "../registry"

// Pictures, shapes, charts and text boxes as objects: insert a picture, and resize, move, wrap, rotate, crop,
// describe or delete the drawing the selection is (a node selection of it). Document commands: a drawing in
// a header or text box changes in that story's session, one undo step each.
//
// Changes to a drawing's size, position, wrapping, rotation, crop and alt text are never tracked, as in Word:
// WordprocessingML has no revision markup for them (ISO/IEC 29500-1 §17.13: the property revisions are
// rPrChange, pPrChange, sectPrChange and the table ones; wp:inline and wp:anchor have no change element),
// and Word's revision types (WdRevisionType) have none for a picture or shape. Inserting and deleting one is
// a run's insertion or deletion and is tracked as such.

interface Selected {
	readonly pos: number
	readonly node: Node
	readonly state: DrawingState
}

function selected(context: DocumentCommandContext): Selected | undefined {
	const session = context.session

	if (session.readOnly) return undefined

	const found = selectedDrawing(session.state)

	return found === undefined ? undefined : { ...found, state: drawingState(context.context, found.node) }
}

function apply(context: DocumentCommandContext, found: Selected, changes: DrawingChanges): boolean {
	const session = context.session
	const tr = changeDrawing(session.state, context.context, found.pos, found.node, changes)

	return tr !== null && session.dispatch(tr)
}

const document: DocumentHandlers = {
	"image.insert": {
		enabled: context => !context.session.readOnly && canInsertPicture(context.session.state),
		run: (context, command) => {
			// Used or not, the pictures are no longer on their way in: the next save may release what nothing uses.
			settleMedia(context.context, [command.media, command.fallback])

			if (pictureRefusal(context.context, command.media, command.fallback) !== undefined) return Promise.resolve(false)

			const tr = insertPicture(context, command)

			return Promise.resolve(tr !== null && context.session.dispatch(tr))
		}
	},
	"drawing.resize": {
		enabled: context => selected(context)?.state.resizable === true,
		run: (context, command) => {
			const found = selected(context)

			if (found?.state.resizable !== true) return false

			return apply(context, found, resizedExtent(found.state, command.cx, command.cy, command.keepRatio === true))
		}
	},
	"drawing.position": {
		enabled: context => {
			const state = selected(context)?.state

			return state !== undefined && state.placeable && state.wrap !== "inline"
		},
		run: (context, command) => {
			const found = selected(context)

			if (found === undefined || !found.state.placeable || found.state.wrap === "inline") return false

			return apply(context, found, {
				posX: command.x,
				posY: command.y,
				relH: command.relH ?? found.state.relH,
				relV: command.relV ?? found.state.relV
			})
		}
	},
	"drawing.wrap": {
		enabled: context => selected(context)?.state.placeable === true,
		run: (context, command) => {
			const found = selected(context)

			if (found?.state.placeable !== true || found.state.wrap === command.mode) return false
			if (command.mode === "inline")
				return apply(context, found, { wrap: "inline", posX: undefined, posY: undefined, relH: undefined, relV: undefined })

			// Floating from inline: where it stands now, from its column and paragraph (as Word places it).
			const position = found.state.wrap === "inline" ? floatingPosition(context, found.pos) : {}

			return apply(context, found, { wrap: command.mode, ...position })
		}
	},
	"drawing.altText": {
		enabled: context => selected(context)?.state.placeable === true,
		run: (context, command) => {
			const found = selected(context)

			return found?.state.placeable === true && apply(context, found, { descr: command.text })
		}
	},
	"drawing.rotate": {
		enabled: context => selected(context)?.state.rotatable === true,
		run: (context, command) => {
			const found = selected(context)

			return found?.state.rotatable === true && apply(context, found, { rot: (found.state.rot + command.by) % 360 })
		}
	},
	"drawing.crop": {
		enabled: context => selected(context)?.state.croppable === true,
		run: (context, command) => {
			const found = selected(context)

			if (found?.state.croppable !== true) return false

			const crop = { left: command.left, top: command.top, right: command.right, bottom: command.bottom }

			if (sameCrop(crop, found.state.crop)) return false

			const extent = croppedExtent(found.state, sourceState(context.context, found.node), crop)

			return (
				extent !== undefined &&
				apply(context, found, {
					...extent,
					cropLeft: crop.left,
					cropTop: crop.top,
					cropRight: crop.right,
					cropBottom: crop.bottom
				})
			)
		}
	},
	"drawing.delete": {
		enabled: context => selected(context) !== undefined,
		run: context => {
			if (selected(context) === undefined) return false

			const session = context.session

			return session.dispatch(session.state.tr.deleteSelection())
		}
	}
}

export const DRAWING_HANDLERS: FeatureHandlers = { document }
