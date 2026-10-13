import type { Node } from "prosemirror-model"
import { docxSchema } from "../edit/schema"
import type { PagedEditor } from "../paged/editor"
import type { PageRect } from "../paged/geometry"
import type { ObjectHit, ObjectHitProvider } from "../paged/objects"
import { mayHoldToc, tocRegions } from "./toc"

// Fields under a page point, for what Word shows on a press there: a field's gray shading (Word's default
// field shading "When selected") and, on a table of contents, its Update Table control. `part` is "toc" for a
// table of contents (the fields inside it go with it), whose box is its lines on the page; a field's box is its
// text on the line pressed. Asked after the other providers: a form field's glyph or a picture wins.

const nodes = docxSchema.nodes

function contains(rect: PageRect, x: number, y: number): boolean {
	return x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height
}

// The field atom next to `pos` whose text holds the point.
function fieldAt(editor: PagedEditor, doc: Node, pos: number, page: number, x: number, y: number): ObjectHit | undefined {
	const $pos = doc.resolve(pos)

	for (const [node, at] of [
		[$pos.nodeAfter, pos],
		[$pos.nodeBefore, pos - ($pos.nodeBefore?.nodeSize ?? 0)]
	] as const) {
		if (node?.type !== nodes.field) continue

		const rect = editor.rangeRects(at, at + node.nodeSize, page).find(candidate => contains(candidate, x, y))

		if (rect !== undefined)
			return {
				kind: "field",
				story: editor.activeStory,
				pos: at,
				box: { x: rect.x, y: rect.y, width: rect.width, height: rect.height }
			}
	}

	return undefined
}

function fieldHit(editor: PagedEditor, page: number, x: number, y: number): ObjectHit | undefined {
	const session = editor.session
	const hit = editor.hitTest(page, x, y)

	if (hit === undefined) return undefined

	const doc = session.state.doc
	const pos = Math.max(0, Math.min(hit.pos, doc.content.size))
	// A table of contents is shaded whole, the fields inside it with it.
	const region = mayHoldToc(session) ? tocRegions(session).find(candidate => candidate.from <= pos && pos <= candidate.to) : undefined
	const rects = region === undefined ? [] : editor.rangeRects(region.from, region.to, page)

	if (region === undefined || !rects.some(rect => contains(rect, x, y))) return fieldAt(editor, doc, pos, page, x, y)

	let left = Infinity
	let top = Infinity
	let right = -Infinity
	let bottom = -Infinity

	for (const rect of rects) {
		left = Math.min(left, rect.x)
		top = Math.min(top, rect.y)
		right = Math.max(right, rect.x + rect.width)
		bottom = Math.max(bottom, rect.y + rect.height)
	}

	return {
		kind: "field",
		story: editor.activeStory,
		pos: region.from,
		box: { x: left, y: top, width: right - left, height: bottom - top },
		part: "toc"
	}
}

export const FIELD_HITS: readonly ObjectHitProvider[] = [{ kind: "field", hit: fieldHit }]
