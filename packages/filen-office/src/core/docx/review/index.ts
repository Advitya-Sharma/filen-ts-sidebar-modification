import type { Mark, Node } from "prosemirror-model"
import type { Transaction } from "prosemirror-state"
import { DocAttrStep, ReplaceAroundStep, ReplaceStep } from "prosemirror-transform"
import { changedRanges } from "../edit/invariants/index"
import { docxSchema, survivorsOf, type MarkerList } from "../edit/schema"

// Where a story's comments and tracked changes are, kept up to date edit by edit: positions are mapped
// through each transaction and only the paragraphs it changed are read again (as spelling does), so no
// feature walks the document per keystroke. Comments, tracked changes and their balloons read this.

export const ReviewKind = {
	// Comment anchors: w:commentRangeStart, w:commentRangeEnd, w:commentReference.
	COMMENT_START: 0,
	COMMENT_END: 1,
	COMMENT_REFERENCE: 2,
	// Text under a revision mark.
	INSERTION: 3,
	DELETION: 4,
	MOVE_FROM: 5,
	MOVE_TO: 6,
	FORMAT: 7,
	// A paragraph mark, row or cell inserted or deleted (RevisionFlag bits on the node).
	PARAGRAPH: 8,
	ROW: 9,
	CELL: 10,
	// A paragraph's properties changed while tracking (the paragraph_change node mark).
	PROPERTIES: 11
} as const

export type ReviewKind = (typeof ReviewKind)[keyof typeof ReviewKind]

export interface ReviewEntry {
	// [from, to) in the story's document; anchors and block revisions take one position.
	readonly from: number
	readonly to: number
	readonly kind: ReviewKind
	// The comment's w:id; a revision's w:id, "#" and its source item when it has none, or for one the editor
	// recorded (no id until saved) "new:" and its author and date.
	readonly id: string
}

const nodes = docxSchema.nodes
const marks = docxSchema.marks
const MARK_KINDS = new Map<unknown, ReviewKind>([
	[marks.insertion, ReviewKind.INSERTION],
	[marks.deletion, ReviewKind.DELETION],
	[marks.move_from, ReviewKind.MOVE_FROM],
	[marks.move_to, ReviewKind.MOVE_TO],
	[marks.format_change, ReviewKind.FORMAT]
])
const PROPERTIES_MARK = marks.paragraph_change
const MARKER_KINDS: Readonly<Record<string, ReviewKind | undefined>> = {
	commentRangeStart: ReviewKind.COMMENT_START,
	commentRangeEnd: ReviewKind.COMMENT_END,
	commentReference: ReviewKind.COMMENT_REFERENCE
}
const BLOCK_KINDS = new Map<unknown, ReviewKind>([
	[nodes.paragraph, ReviewKind.PARAGRAPH],
	[nodes.table_row, ReviewKind.ROW],
	[nodes.table_cell, ReviewKind.CELL]
])

function isComment(kind: number): boolean {
	return kind <= ReviewKind.COMMENT_REFERENCE
}

function markId(mark: Mark): string {
	const id: unknown = mark.attrs["id"]
	const src: unknown = mark.attrs["src"]

	if (typeof id === "string" || typeof id === "number") return String(id)
	if (typeof src === "number" && src >= 0) return `#${String(src)}`

	// Recorded by the editor: one typing burst of one author is one change.
	return `new:${String(mark.attrs["author"] ?? "")}:${String(mark.attrs["date"] ?? "")}`
}

// A change with neither id nor source (a block the editor recorded) counts on its own.
const ANONYMOUS = "#-1"

// Entries found in a range, in document order.
class Found {
	readonly from: number[] = []
	readonly to: number[] = []
	readonly kinds: number[] = []
	readonly ids: string[] = []

	add(from: number, to: number, kind: number, id: string): void {
		const last = this.from.length - 1

		// Pieces of one revision next to each other are one range.
		if (last >= 0 && !isComment(kind) && this.kinds[last] === kind && this.ids[last] === id && this.to[last] === from) {
			this.to[last] = to

			return
		}

		this.from.push(from)
		this.to.push(to)
		this.kinds.push(kind)
		this.ids.push(id)
	}
}

export class ReviewIndex {
	private from = new Int32Array(16)
	private to = new Int32Array(16)
	private kinds = new Uint8Array(16)
	private ids: string[] = []
	// What each entry counts under (a comment's id, a revision's kind and id).
	private keys: string[] = []
	private serial = 0
	private count = 0
	// Entries per comment id and per revision (kind and id): what the counts count.
	private readonly comments = new Map<string, number>()
	private readonly revisions = new Map<string, number>()
	private versionValue = 0

	constructor(doc: Node) {
		const found = new Found()

		scan(doc, 0, doc.content.size, found)
		this.insert(0, found)
	}

	// Changes with every transaction that changed an entry.
	get version(): number {
		return this.versionValue
	}

	get size(): number {
		return this.count
	}

	// Comments with an anchor or reference in the story.
	get commentCount(): number {
		return this.comments.size
	}

	// Tracked changes in the story (an insertion or deletion as one, whatever its pieces).
	get revisionCount(): number {
		return this.revisions.size
	}

	entry(index: number): ReviewEntry | undefined {
		if (index < 0 || index >= this.count) return undefined

		return {
			from: this.from[index] ?? 0,
			to: this.to[index] ?? 0,
			kind: (this.kinds[index] ?? 0) as ReviewKind,
			id: this.ids[index] ?? ""
		}
	}

	// Entries overlapping [from, to) (an anchor at `from` included), in document order.
	between(from: number, to: number): ReviewEntry[] {
		const out: ReviewEntry[] = []
		const end = Math.max(to, from + 1)

		for (let index = 0; index < this.count && (this.from[index] ?? 0) < end; index++) {
			if (Math.max(this.to[index] ?? 0, (this.from[index] ?? 0) + 1) > from) {
				const entry = this.entry(index)

				if (entry !== undefined) out.push(entry)
			}
		}

		return out
	}

	entries(): ReviewEntry[] {
		const out: ReviewEntry[] = []

		for (let index = 0; index < this.count; index++) {
			const entry = this.entry(index)

			if (entry !== undefined) out.push(entry)
		}

		return out
	}

	applyAll(transactions: readonly Transaction[]): void {
		for (const tr of transactions) this.apply(tr)
	}

	apply(tr: Transaction): void {
		if (!tr.docChanged) return

		const doc = tr.doc
		const mapping = tr.mapping
		const regions = changedRegions(doc, tr)

		// Entries outside the changed paragraphs move with the edit; those inside are read again.
		let kept = 0
		let changed = false

		for (let index = 0; index < this.count; index++) {
			const from = mapping.map(this.from[index] ?? 0, 1)
			const to = Math.max(from, mapping.map(this.to[index] ?? 0, -1))
			const end = Math.max(to, from + 1)
			let inside = false

			for (let k = 0; k < regions.length && !inside; k += 2) inside = from < (regions[k + 1] ?? 0) && end > (regions[k] ?? 0)

			if (inside) {
				this.forget(index)
				changed = true

				continue
			}

			this.from[kept] = from
			this.to[kept] = to
			this.kinds[kept] = this.kinds[index] ?? 0
			this.ids[kept] = this.ids[index] ?? ""
			this.keys[kept] = this.keys[index] ?? ""
			kept++
		}

		this.count = kept
		this.ids.length = kept
		this.keys.length = kept

		for (let k = 0; k < regions.length; k += 2) {
			const start = regions[k] ?? 0
			const end = regions[k + 1] ?? 0
			const found = new Found()

			scan(doc, start, end, found)

			if (found.from.length === 0) continue

			changed = true
			this.insert(this.lowerBound(start), found)
		}

		if (changed) this.versionValue++
	}

	// ---- internals --------------------------------------------------------------------------------------

	// The first entry starting at or after `pos`.
	private lowerBound(pos: number): number {
		let low = 0
		let high = this.count

		while (low < high) {
			const middle = (low + high) >> 1

			if ((this.from[middle] ?? 0) < pos) low = middle + 1
			else high = middle
		}

		return low
	}

	private forget(index: number): void {
		const map = isComment(this.kinds[index] ?? 0) ? this.comments : this.revisions
		const key = this.keys[index] ?? ""
		const left = (map.get(key) ?? 1) - 1

		if (left <= 0) map.delete(key)
		else map.set(key, left)
	}

	// Inserts `found` at `at` (entries stay ordered by their start).
	private insert(at: number, found: Found): void {
		const added = found.from.length
		const size = this.count + added

		if (size > this.from.length) {
			const capacity = Math.max(size, this.from.length * 2)
			const from = new Int32Array(capacity)
			const to = new Int32Array(capacity)
			const kinds = new Uint8Array(capacity)

			from.set(this.from.subarray(0, this.count))
			to.set(this.to.subarray(0, this.count))
			kinds.set(this.kinds.subarray(0, this.count))
			this.from = from
			this.to = to
			this.kinds = kinds
		}

		this.from.copyWithin(at + added, at, this.count)
		this.to.copyWithin(at + added, at, this.count)
		this.kinds.copyWithin(at + added, at, this.count)
		this.ids.splice(at, 0, ...found.ids)

		const keys: string[] = []

		for (let k = 0; k < added; k++) {
			const kind = found.kinds[k] ?? 0
			const id = found.ids[k] ?? ""
			const map = isComment(kind) ? this.comments : this.revisions
			const key = isComment(kind)
				? id
				: id === ANONYMOUS
					? `${String(kind)}:anonymous${String(++this.serial)}`
					: `${String(kind)}:${id}`

			keys.push(key)

			this.from[at + k] = found.from[k] ?? 0
			this.to[at + k] = found.to[k] ?? 0
			this.kinds[at + k] = kind
			map.set(key, (map.get(key) ?? 0) + 1)
		}

		this.keys.splice(at, 0, ...keys)
		this.count = size
	}
}

// What a transaction changed, widened to what is read again, merged, as [from, to) pairs: the paragraphs
// (textblocks) it touched, every block it made or changed whole, and the story's end (its markers) when a range
// reaches it or the document's attributes changed.
function changedRegions(doc: Node, tr: Transaction): number[] {
	const size = doc.content.size
	const ranges = changedRanges(tr)
	const pieces: [number, number][] = []

	for (let k = 0; k < ranges.length; k += 2) {
		const $from = doc.resolve(Math.max(0, Math.min(size, ranges[k] ?? 0)))
		const $to = doc.resolve(Math.max(0, Math.min(size, ranges[k + 1] ?? 0)))
		const from = $from.parent.isTextblock && $from.depth > 0 ? $from.before() : $from.pos
		let to = $to.parent.isTextblock && $to.depth > 0 ? $to.after() : Math.min(size, $to.pos + 1)

		// Blocks the edit made (a split's half, rows and cells a paste opened, a node whose attributes changed)
		// are read whole: their end markers and revisions lie past the range.
		for (let depth = $to.depth; depth > 0; depth--) if ($to.before(depth) >= $from.pos) to = Math.max(to, $to.after(depth))

		pieces.push([from, Math.max(from + 1, to)])
	}

	// A deletion across block boundaries joins the blocks on each side: the blocks it ends in now close with
	// the end (markers, revision) of those it took out, so they are read whole.
	tr.steps.forEach((step, index) => {
		const before = tr.docs[index]

		if (step instanceof DocAttrStep) pieces.push([size, size + 1])
		if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep) || before === undefined || step.to <= step.from) return

		const shared = before.resolve(step.from).sharedDepth(step.to)
		const pos = tr.mapping.slice(index).map(step.from, -1)
		const $pos = doc.resolve(Math.max(0, Math.min(size, pos)))

		for (let depth = $pos.depth; depth > shared; depth--) pieces.push([pos, $pos.after(depth)])
	})

	pieces.sort((a, b) => a[0] - b[0])

	const out: number[] = []

	for (const [from, piece] of pieces) {
		// The story's end markers sit at its size: a region reaching it takes them in.
		const to = piece >= size ? size + 1 : piece
		const last = out.length - 1

		if (last > 0 && from <= (out[last] ?? 0)) out[last] = Math.max(out[last] ?? 0, to)
		else out.push(from, to)
	}

	return out
}

function markers(list: unknown, pos: number, found: Found): void {
	if (!Array.isArray(list)) return

	for (const marker of list as MarkerList) {
		const kind = MARKER_KINDS[marker.kind]

		if (kind !== undefined && marker.id !== null) found.add(pos, pos + 1, kind, marker.id)
	}
}

// Entries in [from, to) of `doc`, in document order (block-level ones only where they sit inside it).
function scan(doc: Node, from: number, to: number, found: Found): void {
	const within = (pos: number): boolean => pos >= from && pos < to
	const visit = (node: Node, pos: number): void => {
		if (node.isInline) {
			if (!within(pos)) return

			markers(survivorsOf(node), pos, found)

			for (const mark of node.marks) {
				const kind = MARK_KINDS.get(mark.type)

				if (kind !== undefined) found.add(pos, pos + node.nodeSize, kind, markId(mark))
			}

			return
		}

		const end = pos + node.nodeSize - 1

		if (within(pos)) {
			markers(node.attrs["before"], pos, found)

			// A block kept as bytes holds its anchors in them.
			if (node.isLeaf) markers(node.attrs["markers"], pos, found)
		}

		children(node, pos + 1)

		const kind = BLOCK_KINDS.get(node.type)
		const revision: unknown = node.attrs["revision"]

		if (!within(end)) return
		if (kind !== undefined && typeof revision === "number" && revision !== 0)
			found.add(end, end + 1, kind, `#${String(node.attrs["src"])}`)
		if (PROPERTIES_MARK.isInSet(node.marks) !== undefined)
			found.add(end, end + 1, ReviewKind.PROPERTIES, `#${String(node.attrs["src"])}`)

		markers(node.attrs["end"], end, found)
	}
	const children = (parent: Node, start: number): void => {
		parent.forEach((child, offset) => {
			const at = start + offset

			if (at < to && at + child.nodeSize > from) visit(child, at)
		})
	}

	children(doc, 0)

	if (to >= doc.content.size) markers(doc.attrs["end"], doc.content.size, found)
}
