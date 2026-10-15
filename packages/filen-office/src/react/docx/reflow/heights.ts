import type { Node } from "prosemirror-model"

// The heights of a story's top-level blocks in the reflowed view: measured for blocks that were shown, estimated
// from their size for the others, summed in a Fenwick tree so a block's offset and the block at an offset are
// found in O(log n). An edit splices the entries it replaced; nothing is walked per keystroke.

// Estimates at the default 11 pt text: a line's height, a character's width and a paragraph's spacing (px).
const LINE = 19
const CHARACTER = 7
const SPACING = 11
const ROW = 28

// A block's estimated height at `width` px.
export function estimateHeight(node: Node, width: number): number {
	const perLine = Math.max(12, Math.floor(Math.max(1, width) / CHARACTER))

	if (node.isTextblock) return Math.ceil(Math.max(1, node.content.size) / perLine) * LINE + SPACING

	if (node.type.spec["tableRole"] === "table") {
		// Rows of text in narrow columns: a row's share of the text, a cell's width a column of the grid.
		const columns = Math.max(1, node.firstChild?.childCount ?? 1)
		const lines = Math.ceil((node.content.size * columns) / perLine)

		return Math.max(node.childCount * ROW, (lines / columns) * LINE) + SPACING
	}

	return Math.max(1, Math.ceil(node.content.size / perLine)) * LINE + SPACING
}

export class BlockHeights {
	private values = new Float64Array(0)
	private measured = new Uint8Array(0)
	private tree = new Float64Array(1)
	private length = 0

	get count(): number {
		return this.length
	}

	// Every block of `doc` estimated anew (the view's width changed).
	reset(doc: Node, width: number): void {
		const count = doc.childCount

		this.length = count
		this.values = new Float64Array(count)
		this.measured = new Uint8Array(count)

		for (let k = 0; k < count; k++) this.values[k] = estimateHeight(doc.child(k), width)

		this.build()
	}

	// Blocks [from, from + removed) were replaced by `inserted` (an edit): those in place of an old one keep what
	// it was measured or taken to be (its element is the old one's until it is measured again), the others are
	// estimated at `width`.
	splice(from: number, removed: number, inserted: readonly Node[], width: number): void {
		// Blocks edited in place: each keeps what it was taken to be (nothing else moved).
		if (removed === inserted.length) return

		const count = this.length - removed + inserted.length
		const kept = Math.min(removed, inserted.length)
		const values = new Float64Array(count)
		const measured = new Uint8Array(count)

		values.set(this.values.subarray(0, from + kept))
		measured.set(this.measured.subarray(0, from + kept))
		values.set(this.values.subarray(from + removed, this.length), from + inserted.length)
		measured.set(this.measured.subarray(from + removed, this.length), from + inserted.length)

		for (let k = kept; k < inserted.length; k++) {
			const node = inserted[k]

			values[from + k] = node === undefined ? LINE : estimateHeight(node, width)
		}

		this.values = values
		this.measured = measured
		this.length = count
		this.build()
	}

	// Block `index` measured `height` px high; true when that changed what it was taken to be.
	set(index: number, height: number): boolean {
		if (index < 0 || index >= this.length || !Number.isFinite(height) || height < 0) return false

		const delta = height - (this.values[index] ?? 0)

		this.measured[index] = 1

		if (Math.abs(delta) < 0.5) return false

		this.values[index] = height

		for (let at = index + 1; at < this.tree.length; at += at & -at) this.tree[at] = (this.tree[at] ?? 0) + delta

		return true
	}

	isMeasured(index: number): boolean {
		return this.measured[index] === 1
	}

	height(index: number): number {
		return this.values[index] ?? 0
	}

	// The top of block `index` (the sum of the heights before it).
	top(index: number): number {
		let sum = 0

		for (let at = Math.min(index, this.length); at > 0; at -= at & -at) sum += this.tree[at] ?? 0

		return sum
	}

	get total(): number {
		return this.top(this.length)
	}

	// The block at offset `y`: the first whose bottom is below it (the last for an offset past the end).
	indexAt(y: number): number {
		if (this.length === 0) return 0
		if (y <= 0) return 0

		let index = 0
		let rest = y
		let step = 1

		while (step * 2 <= this.length) step *= 2

		for (; step > 0; step >>= 1) {
			const next = index + step

			if (next <= this.length && (this.tree[next] ?? 0) < rest) {
				index = next
				rest -= this.tree[next] ?? 0
			}
		}

		return Math.min(index, this.length - 1)
	}

	private build(): void {
		const tree = new Float64Array(this.length + 1)

		for (let k = 0; k < this.length; k++) {
			const at = k + 1

			tree[at] = (tree[at] ?? 0) + (this.values[k] ?? 0)

			const parent = at + (at & -at)

			if (parent <= this.length) tree[parent] = (tree[parent] ?? 0) + (tree[at] ?? 0)
		}

		this.tree = tree
	}
}
