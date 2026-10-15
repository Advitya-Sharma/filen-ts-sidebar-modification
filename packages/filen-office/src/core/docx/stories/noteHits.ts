import type { Node } from "prosemirror-model"
import { docxSchema } from "../edit/schema"
import type { PagedEditor } from "../paged/editor"
import type { ObjectHit, ObjectHitProvider } from "../paged/objects"

// Note references under a page point, for what a press on one offers (Go to note): a reference in the text, or
// the note's own mark at its start (Word's w:footnoteRef / w:endnoteRef) while the note is edited. A mark is a
// small superscript: a press this close to its box (screen px) still finds it.

const nodes = docxSchema.nodes
const GRAB = 3
const NOTE_MARKS: ReadonlySet<string> = new Set(["FootnoteMark", "EndnoteMark"])

function noteMark(node: Node | null | undefined): node is Node {
	return node?.type === nodes.note_ref || (node?.type === nodes.run_object && NOTE_MARKS.has(node.attrs["kind"] as string))
}

function noteRefHit(editor: PagedEditor, page: number, x: number, y: number, zoom = 1): ObjectHit | undefined {
	const hit = editor.hitTest(page, x, y)

	if (hit === undefined) return undefined

	const doc = editor.session.state.doc
	const $pos = doc.resolve(Math.max(0, Math.min(hit.pos, doc.content.size)))
	const grab = GRAB / Math.max(zoom, 0.01)

	for (const [node, at] of [
		[$pos.nodeAfter, $pos.pos],
		[$pos.nodeBefore, $pos.pos - ($pos.nodeBefore?.nodeSize ?? 0)]
	] as const) {
		if (!noteMark(node)) continue

		const rect = editor
			.rangeRects(at, at + node.nodeSize, page)
			.find(box => x >= box.x - grab && x <= box.x + box.width + grab && y >= box.y - grab && y <= box.y + box.height + grab)

		if (rect !== undefined)
			return {
				kind: "noteRef",
				story: editor.activeStory,
				pos: at,
				box: { x: rect.x, y: rect.y, width: rect.width, height: rect.height }
			}
	}

	return undefined
}

export const NOTE_REF_HITS: readonly ObjectHitProvider[] = [{ kind: "noteRef", hit: noteRefHit }]
