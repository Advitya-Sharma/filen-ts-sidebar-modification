import type { Node } from "prosemirror-model"
import { docxSchema } from "../edit/schema"
import type { PagedEditor } from "../paged/editor"
import type { PageRect } from "../paged/geometry"
import type { ObjectHit, ObjectHitProvider } from "../paged/objects"

// Equations under a page point, for the equation editor: the equation next to the position the point hits
// whose box on the page holds the point, with that box.

const nodes = docxSchema.nodes

function contains(rect: PageRect, x: number, y: number): boolean {
	return x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height
}

function equationAt(editor: PagedEditor, page: number, x: number, y: number): ObjectHit | undefined {
	const hit = editor.hitTest(page, x, y)

	if (hit === undefined) return undefined

	const doc = editor.session.state.doc
	const $pos = doc.resolve(Math.max(0, Math.min(hit.pos, doc.content.size)))
	const before: Node | null = $pos.nodeBefore

	for (const [node, at] of [
		[$pos.nodeAfter, $pos.pos],
		[before, $pos.pos - (before?.nodeSize ?? 0)]
	] as const) {
		if (node?.type !== nodes.math) continue

		const rect = editor.rangeRects(at, at + node.nodeSize, page).find(candidate => contains(candidate, x, y))

		if (rect !== undefined)
			return {
				kind: "equation",
				story: editor.activeStory,
				pos: at,
				box: { x: rect.x, y: rect.y, width: rect.width, height: rect.height }
			}
	}

	return undefined
}

export const EQUATION_HITS: readonly ObjectHitProvider[] = [{ kind: "equation", hit: equationAt }]
