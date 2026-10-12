import type { Node } from "prosemirror-model"
import { docxSchema } from "../edit/index"
import { SOURCE_MAP_INTS } from "../layout/index"
import { Item, type StoryPart } from "../model/index"
import { Ns } from "../model/namespaces"

// Exact maps between a paragraph of the editing document and the text its source laid out: every inline
// node of the paragraph comes from items of the part the layout read (the paragraph's own source part, or
// the overlay its edited version was written to), and the layout records where each item's content landed.
// Generated text (list markers, field results) maps to the node it belongs to; nodes the layout shows
// nothing for (range markers, hidden text, anchored drawings) map to where they stand.

// A source's item map: [start, end) of a page's map pool (no pool: the layout recorded none).
export interface SourceMap {
	readonly pool: Int32Array | undefined
	readonly start: number
	readonly end: number
}

export interface OffsetMap {
	// Paragraph content offset (0..content size) -> text offset of the gap before it.
	readonly toLayout: Int32Array
	// Text offset (0..text length) -> paragraph content offset.
	readonly toPm: Int32Array
}

const nodes = docxSchema.nodes
// Zero-width nodes a writer may leave out next to edited text (Word regenerates them).
const DROPPABLE = new Set([nodes.proof_err, nodes.rendered_break])
const EDGE_WHITESPACE = /^[ \t\r\n]|[ \t\r\n]$/
const LEADING_WHITESPACE = /^[ \t\r\n]+/
const EDGE_TRIM = /^[ \t\r\n]+|[ \t\r\n]+$/g

// Where the content of each item of the paragraph `item` of `part` landed in the source's text: the
// source map entries belonging to it.
class ItemSpans {
	readonly items: number[] = []
	readonly starts: number[] = []
	readonly ends: number[] = []
	// The offset where the paragraph's own content starts (after a list marker).
	readonly contentStart: number
	// What range() found.
	rangeStart = 0
	rangeEnd = 0

	constructor(part: StoryPart, item: number, map: SourceMap, textLength: number) {
		const end = part.end(item)

		if (map.pool !== undefined) {
			for (let k = map.start; k + SOURCE_MAP_INTS <= map.end; k += SOURCE_MAP_INTS) {
				const pool = map.pool
				const at = pool[k + 1] ?? -1

				if (pool[k] !== part.index || at < item || at >= end) continue

				this.items.push(at)
				this.starts.push(pool[k + 2] ?? 0)
				this.ends.push(pool[k + 3] ?? 0)
			}
		}

		this.contentStart = this.starts[0] ?? textLength
	}

	// The entry of `item`, or the last before it (-1 none).
	at(item: number): number {
		let low = 0
		let high = this.items.length - 1
		let found = -1

		while (low <= high) {
			const middle = (low + high) >> 1

			if ((this.items[middle] ?? 0) <= item) {
				found = middle
				low = middle + 1
			} else {
				high = middle - 1
			}
		}

		return found
	}

	// Whether items [from, to) laid out text, and the range they laid out ([rangeStart, rangeEnd), set here:
	// no pair is made per character).
	range(from: number, to: number): boolean {
		let k = this.at(to - 1)

		if (k < 0 || (this.items[k] ?? 0) < from) return false

		this.rangeEnd = this.ends[k] ?? 0

		while (k > 0 && (this.items[k - 1] ?? 0) >= from) k--

		this.rangeStart = this.starts[k] ?? 0

		return true
	}

	// Where an item that laid out nothing stands: after what the items before it laid out.
	position(item: number): number {
		const k = this.at(item - 1)

		return k < 0 ? this.contentStart : (this.ends[k] ?? this.contentStart)
	}
}

// Text items of a run in order, with the number of characters the editor reads from each.
function runTexts(part: StoryPart, run: number): { items: number[]; lengths: number[]; leads: number[] } {
	const items: number[] = []
	const lengths: number[] = []
	const leads: number[] = []

	for (let child = run + 1; child < part.end(run); child = part.end(child)) {
		const kind = part.kinds[child]

		if (kind !== Item.Text && kind !== Item.DeletedText) continue

		const text = part.text(child)
		const trimmed = EDGE_WHITESPACE.test(text) && !part.preservesSpace(child) && !inheritsPreserve(part, part.node(child))
		const lead = trimmed ? (LEADING_WHITESPACE.exec(text)?.[0].length ?? 0) : 0

		items.push(child)
		lengths.push(trimmed ? text.replace(EDGE_TRIM, "").length : text.length)
		leads.push(lead)
	}

	return { items, lengths, leads }
}

function inheritsPreserve(part: StoryPart, node: number): boolean {
	const document = part.document

	for (let at = document.parent(node); at > 0; at = document.parent(at)) {
		const space = part.reader.attr(at, Ns.Xml, "space")

		if (space !== undefined) return space === "preserve"
	}

	return false
}

// Text offsets of the gaps of `paragraph` (built from `part`, so its sources are items of `part`).
function layoutOffsets(paragraph: Node, part: StoryPart, item: number, map: SourceMap, textLength: number): Int32Array {
	const spans = new ItemSpans(part, item, map, textLength)
	const out = new Int32Array(paragraph.content.size + 1)
	const runs = new Map<number, { texts: ReturnType<typeof runTexts>; index: number; used: number }>()
	let at = spans.contentStart

	paragraph.forEach((child, offset) => {
		if (child.isText) {
			const run = child.marks.find(mark => mark.type === docxSchema.marks.run)
			const src = run === undefined ? -1 : (run.attrs["src"] as number)
			let cursor = src < 0 ? undefined : runs.get(src)

			if (cursor === undefined && src >= 0) {
				cursor = { texts: runTexts(part, src), index: 0, used: 0 }
				runs.set(src, cursor)
			}

			const length = child.text?.length ?? 0

			for (let c = 0; c < length; c++) {
				// The character's item: the run's text items consumed in order.
				while (
					cursor !== undefined &&
					cursor.index < cursor.texts.items.length &&
					cursor.used >= (cursor.texts.lengths[cursor.index] ?? 0)
				) {
					cursor.index++
					cursor.used = 0
				}

				const textItem = cursor?.texts.items[cursor.index]

				if (cursor === undefined || textItem === undefined) {
					out[offset + c] = at

					continue
				}

				const found = spans.range(textItem, textItem + 1)
				const lead = cursor.texts.leads[cursor.index] ?? 0
				const gap = found ? Math.min(spans.rangeStart + lead + cursor.used, spans.rangeEnd) : spans.position(textItem)

				out[offset + c] = gap
				at = found ? Math.min(gap + 1, spans.rangeEnd) : gap
				cursor.used++
			}

			return
		}

		const src = child.attrs["src"]

		if (typeof src !== "number" || src < 0 || child.type === nodes.proof_err) {
			out[offset] = at

			return
		}

		// A field covers its items up to its end run; any other atom its element.
		const last = child.type === nodes.field ? ((child.attrs["end"] as number | undefined) ?? src) : src
		if (spans.range(src, part.end(last))) {
			out[offset] = spans.rangeStart
			at = spans.rangeEnd
		} else {
			out[offset] = Math.max(at, spans.position(src))
			at = out[offset] ?? at
		}
	})

	out[paragraph.content.size] = at

	return out
}

// Content offsets of `node` (the editing document's paragraph) -> offsets of `shadow` (the same paragraph
// built from the part the layout read): the same nodes, except zero-width ones a writer may drop.
function shadowOffsets(node: Node, shadow: Node): Int32Array {
	const out = new Int32Array(node.content.size + 1)

	if (node === shadow) {
		for (let k = 0; k < out.length; k++) out[k] = k

		return out
	}

	let s = 0
	let sIndex = 0
	// Characters of the shadow's current child already matched.
	let sUsed = 0
	const current = (): Node | undefined => (sIndex < shadow.childCount ? shadow.child(sIndex) : undefined)
	const advance = (size: number): void => {
		s += size
		sIndex++
		sUsed = 0
	}

	node.forEach((child, offset) => {
		if (child.isText) {
			const length = child.text?.length ?? 0

			for (let c = 0; c < length; c++) {
				for (;;) {
					const other = current()

					if (other !== undefined && !other.isText && DROPPABLE.has(other.type)) {
						advance(1)

						continue
					}

					break
				}

				const other = current()

				out[offset + c] = s

				if (other?.isText === true) {
					s++
					sUsed++

					if (sUsed >= (other.text?.length ?? 0)) {
						sIndex++
						sUsed = 0
					}
				}
			}

			return
		}

		for (;;) {
			const other = current()

			if (other !== undefined && other.type !== child.type && !other.isText && DROPPABLE.has(other.type)) {
				advance(1)

				continue
			}

			break
		}

		const other = current()

		out[offset] = s

		if (other !== undefined && other.type === child.type && sUsed === 0) advance(other.nodeSize)
	})

	out[node.content.size] = shadow.content.size

	return out
}

// The maps between `node` and the text `text` its source laid out (`map`: the source's item map), `shadow`
// being `node` as built from `part` (the part the layout read) and `item` that paragraph's item there.
export function offsetMap(node: Node, shadow: Node, part: StoryPart, item: number, text: string, map: SourceMap): OffsetMap {
	const layout = layoutOffsets(shadow, part, item, map, text.length)
	const toShadow = shadowOffsets(node, shadow)
	const size = node.content.size
	const toLayout = new Int32Array(size + 1)

	for (let k = 0; k <= size; k++) toLayout[k] = Math.max(0, Math.min(text.length, layout[toShadow[k] ?? 0] ?? 0))

	const toPm = new Int32Array(text.length + 1).fill(-1)

	for (let k = 0; k <= size; k++) {
		const at = toLayout[k] ?? 0

		if ((toPm[at] ?? -1) < 0) toPm[at] = k
	}

	let last = 0

	for (let k = 0; k <= text.length; k++) {
		const value = toPm[k] ?? -1

		if (value < 0) toPm[k] = last
		else last = value
	}

	return { toLayout, toPm }
}
