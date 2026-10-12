import { DRAWING_HITS } from "../objects/hits"
import { FORM_HITS } from "../forms/hits"
import type { PagedEditor } from "./editor"

// What sits on the pages besides text, for pointer input: a feature registers how it finds its objects at a
// page point (pictures and shapes, content controls, comment anchors, fields), and a view hands a press to
// the feature's handler before placing the caret (DocxEditor `interactions`).

export type ObjectHitKind = "drawing" | "control" | "commentMarker" | "field" | "link" | "noteRef" | "equation"

export interface ObjectHit {
	readonly kind: ObjectHitKind
	// The story (index) and position of the object in that story's editing document.
	readonly story: number
	readonly pos: number
	// The object's box on the page (page coordinates), for selection frames and handles.
	readonly box?: { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
	// A part of the object (a resize handle, a checkbox glyph), as the provider names it.
	readonly part?: string
}

export interface ObjectHitProvider {
	readonly kind: ObjectHitKind
	// The object at page point (x, y) of page `page`, or undefined. Called on presses only.
	hit(editor: PagedEditor, page: number, x: number, y: number): ObjectHit | undefined
}

// In the order they are asked: the first hit wins.
export const OBJECT_HIT_PROVIDERS: readonly ObjectHitProvider[] = [...FORM_HITS, ...DRAWING_HITS]

export function hitObject(editor: PagedEditor, page: number, x: number, y: number): ObjectHit | undefined {
	for (const provider of OBJECT_HIT_PROVIDERS) {
		const hit = provider.hit(editor, page, x, y)

		if (hit !== undefined) return hit
	}

	return undefined
}
