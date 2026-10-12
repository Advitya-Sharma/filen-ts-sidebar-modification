import type { Node } from "prosemirror-model"
import { NodeSelection } from "prosemirror-state"
import {
	FRAGMENT_FLOATS,
	FRAGMENT_INTS,
	FRAGMENT_KIND_MASK,
	FragmentKind,
	LINE_FLOATS,
	LINE_INTS,
	LineFloat,
	LineInt
} from "../../text/index"
import { docxSchema } from "../edit/schema"
import { Layer, OBJECT_FLOATS, OBJECT_INTS, ObjectKind, type ObjectTable, type PageDisplayList } from "../layout/index"
import { Item, type StoryPart } from "../model/index"
import type { PagedEditor } from "../paged/editor"
import type { StoryEdits } from "../paged/edits"
import type { ObjectHit, ObjectHitProvider } from "../paged/objects"

// Pictures, shapes, charts and text boxes under a page point, front to back: the drawing (its position in the
// story being edited, or in the body when the story being edited is a text box of it), its box on the page,
// and, on the drawing the selection is, the resize handle the point is on ("handle:nw" ... "handle:w"). A
// shape holding a text box is hit on its edge only (inside, a press goes to its text, as in Word); a drawing
// behind the text only where no line of text is.

const nodes = docxSchema.nodes
// How far from a handle's centre, or a text box's edge, a press still counts (screen px).
export const HANDLE_GRAB = 10
export const DRAWING_HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const

export type DrawingHandle = (typeof DRAWING_HANDLES)[number]

// A handle's place on a box, as fractions of its width and height.
export const HANDLE_PLACES: Readonly<Record<DrawingHandle, readonly [number, number]>> = {
	nw: [0, 0],
	n: [0.5, 0],
	ne: [1, 0],
	e: [1, 0.5],
	se: [1, 1],
	s: [0.5, 1],
	sw: [0, 1],
	w: [0, 0.5]
}

export interface PageBox {
	readonly x: number
	readonly y: number
	readonly width: number
	readonly height: number
}

// The paragraph item holding item `item` of `part`, or -1.
function paragraphOf(part: StoryPart, item: number): number {
	for (let at = item; at >= 0; at = part.parent(at)) if (part.kinds[at] === Item.Paragraph) return at

	return -1
}

// The drawing nodes of a paragraph with their offsets in it, in order.
function drawingsIn(paragraph: Node): { node: Node; offset: number }[] {
	const out: { node: Node; offset: number }[] = []

	paragraph.forEach((child, offset) => {
		if (child.type === nodes.drawing) out.push({ node: child, offset })
	})

	return out
}

// Where item `item` of `part` (a laid-out drawing) stands in a story's editing document, or undefined.
function positionIn(story: StoryEdits, part: StoryPart, item: number): number | undefined {
	const paragraph = paragraphOf(part, item)

	if (paragraph < 0) return undefined

	const found = story.paragraphOf(part.index, paragraph)

	if (found === undefined) return undefined

	// The paragraph as laid out (`shadow`, built from the part laid out) and as edited have their drawings in
	// the same order.
	const laid = drawingsIn(found.shadow)
	const edited = drawingsIn(found.node)
	const index = laid.findIndex(entry => {
		const own = entry.node.attrs["src"] as number

		return own >= 0 && own <= item && item < part.end(own)
	})
	const match = edited[index]

	return index < 0 || match === undefined || laid.length !== edited.length ? undefined : found.pos + 1 + match.offset
}

function onText(page: PageDisplayList, x: number, y: number): boolean {
	const text = page.text

	for (let line = 0; line < text.lineCount; line++) {
		const f = line * LINE_FLOATS
		const left = text.lineFloats[f + LineFloat.X] ?? 0
		const top = text.lineFloats[f + LineFloat.TOP] ?? 0

		if (
			x >= left &&
			x <= left + (text.lineFloats[f + LineFloat.WIDTH] ?? 0) &&
			y >= top &&
			y <= top + (text.lineFloats[f + LineFloat.HEIGHT] ?? 0)
		)
			return true
	}

	return false
}

interface Placed {
	readonly id: number
	readonly layer: number
	readonly box: PageBox
}

// The objects on a page: those placed apart (floating), then those on its lines (inline: object fragments, the
// object's ascent above the baseline), in drawing order.
function placedObjects(page: PageDisplayList, objects: ObjectTable): Placed[] {
	const out: Placed[] = []

	for (let o = 0; o < page.objectCount; o++) {
		const f = o * OBJECT_FLOATS

		out.push({
			id: page.objectInts[o * OBJECT_INTS] ?? -1,
			layer: page.objectInts[o * OBJECT_INTS + 1] ?? Layer.TEXT,
			box: {
				x: page.objectFloats[f] ?? 0,
				y: page.objectFloats[f + 1] ?? 0,
				width: page.objectFloats[f + 2] ?? 0,
				height: page.objectFloats[f + 3] ?? 0
			}
		})
	}

	const text = page.text

	for (let line = 0; line < text.lineCount; line++) {
		const baseline = text.lineFloats[line * LINE_FLOATS + LineFloat.BASELINE] ?? 0
		const first = text.lineInts[line * LINE_INTS + LineInt.FIRST_FRAGMENT] ?? 0
		const count = text.lineInts[line * LINE_INTS + LineInt.FRAGMENT_COUNT] ?? 0

		for (let k = first; k < first + count; k++) {
			if (((text.fragmentInts[k * FRAGMENT_INTS + 3] ?? 0) & FRAGMENT_KIND_MASK) !== FragmentKind.OBJECT) continue

			const id = text.fragmentInts[k * FRAGMENT_INTS + 4] ?? -1
			const height = objects.heights[id] ?? 0

			if (id < 0 || !(height > 0)) continue

			out.push({
				id,
				layer: Layer.TEXT,
				box: {
					x: text.fragmentFloats[k * FRAGMENT_FLOATS] ?? 0,
					y: baseline - (objects.ascents[id] ?? height),
					width: text.fragmentFloats[k * FRAGMENT_FLOATS + 1] ?? objects.widths[id] ?? 0,
					height
				}
			})
		}
	}

	return out
}

// The handle of a box at (x, y) within `grab` (page px), or undefined.
export function handleAt(box: PageBox, x: number, y: number, grab: number): DrawingHandle | undefined {
	for (const handle of DRAWING_HANDLES) {
		const [fx, fy] = HANDLE_PLACES[handle]

		if (Math.abs(box.x + box.width * fx - x) <= grab && Math.abs(box.y + box.height * fy - y) <= grab) return handle
	}

	return undefined
}

function drawingHit(editor: PagedEditor, pageIndex: number, x: number, y: number, zoom = 1): ObjectHit | undefined {
	const page = editor.page(pageIndex)

	if (page === undefined) return undefined

	const objects = editor.layout.objects
	const parts = editor.edits.model.parts
	const grab = HANDLE_GRAB / Math.max(zoom, 0.01)
	const selection = editor.session.state.selection
	const selected = selection instanceof NodeSelection && selection.node.type === nodes.drawing ? selection.from : -1
	const stories: readonly StoryEdits[] = editor.story === editor.body.edits ? [editor.story] : [editor.story, editor.body.edits]
	let best: { hit: ObjectHit; rank: number; order: number } | undefined

	const placed = placedObjects(page, objects)

	for (let o = 0; o < placed.length; o++) {
		const { id, layer, box } = placed[o] as Placed
		const kind = objects.kinds[id]

		if (kind !== ObjectKind.DRAWING && kind !== ObjectKind.VML) continue
		if (x < box.x - grab || x > box.x + box.width + grab || y < box.y - grab || y > box.y + box.height + grab) continue

		const part = parts[objects.parts[id] ?? -1]
		const item = objects.items[id] ?? -1

		if (part === undefined || item < 0) continue

		let story: StoryEdits | undefined
		let pos: number | undefined

		for (const candidate of stories) {
			pos = positionIn(candidate, part, item)

			if (pos !== undefined) {
				story = candidate

				break
			}
		}

		if (story === undefined || pos === undefined) continue

		const handle = story === editor.story && pos === selected ? handleAt(box, x, y, grab) : undefined

		if (handle === undefined) {
			if (x < box.x || x > box.x + box.width || y < box.y || y > box.y + box.height) continue

			// A text box is entered inside, grabbed on its edge.
			const boxed = (objects.textboxes[id]?.length ?? 0) > 0
			const edge = Math.min(x - box.x, box.x + box.width - x, y - box.y, box.y + box.height - y) <= grab

			if (boxed && !edge) continue
			if (layer === Layer.BEHIND && onText(page, x, y)) continue
		}

		const hit: ObjectHit = {
			kind: "drawing",
			story: story.story.story.index,
			pos,
			box,
			...(handle === undefined ? {} : { part: `handle:${handle}` })
		}
		// Front objects first, later ones on the same layer over earlier ones; a handle over everything.
		const rank = handle !== undefined ? Number.MAX_SAFE_INTEGER : layer

		if (best === undefined || rank > best.rank || (rank === best.rank && o > best.order)) best = { hit, rank, order: o }
	}

	return best?.hit
}

export const DRAWING_HITS: readonly ObjectHitProvider[] = [{ kind: "drawing", hit: drawingHit }]
