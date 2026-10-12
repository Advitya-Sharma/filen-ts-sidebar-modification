import type { Node } from "prosemirror-model"
import { NO_SOURCE, type MarkerAttrs, type MarkerKind } from "../edit/schema"
import { Item } from "../model/index"
import type { Out } from "./markup"
import { src } from "./source"
import type { StoryWriting } from "./story"
import { DocxWriteError } from "./values"

// Source items that are range markers (bookmarks, comment ranges, move ranges, permissions).
export const MARKER_ITEMS = new Set<number>([
	Item.BookmarkStart,
	Item.BookmarkEnd,
	Item.CommentRangeStart,
	Item.CommentRangeEnd,
	Item.MoveFromRangeStart,
	Item.MoveFromRangeEnd,
	Item.MoveToRangeStart,
	Item.MoveToRangeEnd,
	Item.PermissionStart,
	Item.PermissionEnd,
	Item.CustomXmlRangeMarker
])

// Writes a range marker without a source: one of a block's marker lists (`node` undefined) or an inline
// marker node.
export interface MarkerWriter {
	readonly kind: MarkerKind
	write(w: StoryWriting, out: Out, marker: MarkerAttrs, node: Node | undefined): void
}

// Writers of new markers by kind; a new marker of any other kind is refused.
const NEW_MARKERS: ReadonlyMap<MarkerKind, MarkerWriter> = new Map()

// A marker standing before or after blocks (MarkerList).
export function writeListMarker(w: StoryWriting, out: Out, marker: MarkerAttrs): void {
	if (marker.src !== NO_SOURCE) {
		w.copyWhole(out, w.part.node(marker.src))

		return
	}

	const writer = NEW_MARKERS.get(marker.kind)

	if (writer === undefined) throw new DocxWriteError("missingSource", "A new range marker")

	writer.write(w, out, marker, undefined)
}

// A marker node inside a paragraph.
export function writeMarkerAtom(w: StoryWriting, out: Out, node: Node): void {
	const own = src(node)

	if (own !== NO_SOURCE) {
		w.copyNode(out, w.part.node(own))

		return
	}

	const marker: MarkerAttrs = { src: own, kind: node.attrs["kind"] as MarkerKind, id: node.attrs["id"] as string | null }
	const writer = NEW_MARKERS.get(marker.kind)

	if (writer === undefined) throw new DocxWriteError("missingSource", `A new ${node.type.name}`)

	writer.write(w, out, marker, node)
}
