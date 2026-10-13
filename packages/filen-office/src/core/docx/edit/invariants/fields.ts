import type { Node, Slice } from "prosemirror-model"
import { isHistoryTransaction } from "prosemirror-history"
import type { EditorState, Transaction } from "prosemirror-state"
import { AddMarkStep, ReplaceAroundStep, ReplaceStep } from "prosemirror-transform"
import { SetMarkSetsStep } from "../marks"
import { NO_SOURCE, docxSchema } from "../schema"
import { changedRanges, isInternal, markInternal } from "./core"
import type { InvariantStage } from "./index"

// Fields spanning paragraphs (a table of contents: field_char begin, its instruction texts when it is the
// file's, separate and end in different paragraphs) stay whole: an edit that deletes one of their characters
// deletes the field from its begin to its end, as Word deletes a field a selection takes a boundary of;
// characters an edit leaves without their partners otherwise (a paste of part of a table) go, and the rest is
// plain text. A paragraph split inside a field's code (between its begin and separate, where nothing shows)
// puts the code after the split: a code over a paragraph mark would hide what follows it. A tracked deletion
// takes a field's characters all or none: those it took without the rest are kept (what the field holds stays
// deleted). Only edits whose steps remove or insert a field character, split paragraphs or mark content
// deleted are looked at. A move that would take such a field apart is refused (see brokenMove).

const nodes = docxSchema.nodes

function sliceHas(slice: Slice): boolean {
	let found = false

	slice.content.descendants(node => {
		if (node.type === nodes.field_char) found = true

		return !found && !node.isLeaf
	})

	return found
}

function rangeHas(doc: Node, from: number, to: number): boolean {
	let found = false

	if (to > from)
		doc.nodesBetween(from, to, node => {
			if (node.type === nodes.field_char) found = true

			return !found && !node.isLeaf
		})

	return found
}

// Where a transaction removed field characters (positions in its own final document), or undefined when none
// of its steps removed or inserted one.
function touched(tr: Transaction): number[] | undefined {
	const removed: number[] = []
	let any = false

	tr.steps.forEach((step, index) => {
		if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep)) return

		const doc = tr.docs[index] as Node
		const gone =
			step instanceof ReplaceAroundStep
				? rangeHas(doc, step.from, step.gapFrom) || rangeHas(doc, step.gapTo, step.to)
				: rangeHas(doc, step.from, step.to)

		if (gone) removed.push(tr.mapping.slice(index).map(step.from, -1))
		if (gone || sliceHas(step.slice)) any = true
	})

	return any ? removed : undefined
}

interface Char {
	readonly pos: number
	readonly kind: string
}

// The field characters of `doc` without partners.
function unmatched(doc: Node): Char[] {
	const open: { begin: Char; separate: Char | undefined; instructions: Char[] }[] = []
	const stray: Char[] = []

	doc.descendants((node, pos) => {
		if (node.type !== nodes.field_char) return !node.isLeaf

		const char = { pos, kind: node.attrs["kind"] as string }
		const top = open[open.length - 1]

		if (char.kind === "begin") open.push({ begin: char, separate: undefined, instructions: [] })
		else if ((char.kind === "instruction" || char.kind === "deletedInstruction") && top !== undefined && top.separate === undefined)
			top.instructions.push(char)
		else if (char.kind === "separate" && top !== undefined && top.separate === undefined) top.separate = char
		else if (char.kind === "end" && top !== undefined) open.pop()
		else stray.push(char)

		return false
	})

	for (const field of open) stray.push(field.begin, ...field.instructions, ...(field.separate === undefined ? [] : [field.separate]))

	return stray.sort((a, b) => a.pos - b.pos)
}

// Whether [from, to) can go in one deletion: paragraphs of one container, no section break or locked block.
function deletable(doc: Node, from: number, to: number): boolean {
	const $from = doc.resolve(from)
	const $to = doc.resolve(to)

	if (
		!$from.parent.isTextblock ||
		!$to.parent.isTextblock ||
		$from.depth !== $to.depth ||
		$from.node($from.depth - 1) !== $to.node($to.depth - 1)
	)
		return false

	let safe = true

	doc.nodesBetween(from, to, (node, pos) => {
		if (!safe) return false

		if (node.type === nodes.locked_block || node.type === nodes.table || node.type === nodes.sdt) safe = false
		else if (node.type === nodes.paragraph && (node.attrs["sectPr"] as number) !== NO_SOURCE && pos + node.nodeSize - 1 < to)
			safe = false

		return safe && !node.isTextblock
	})

	return safe
}

// Whether a step of `tr` put a block boundary into the document (a split, paragraphs pasted).
function splits(tr: Transaction): boolean {
	return tr.steps.some(step => (step instanceof ReplaceStep || step instanceof ReplaceAroundStep) && step.slice.openStart > 0)
}

// Where a paragraph ends inside a field's code: the offset of its last begin with no separate after it (only
// instruction texts and zero-width markers follow), or -1.
function openCode(paragraph: Node): number {
	let open = -1

	paragraph.forEach((child, offset) => {
		if (child.type === nodes.field_char) {
			const kind = child.attrs["kind"] as string

			if (kind === "begin") open = offset
			else if (kind !== "instruction" && kind !== "deletedInstruction") open = -1
		} else if (child.type !== nodes.marker && child.type !== nodes.proof_err) {
			open = -1
		}
	})

	return open
}

// Codes split from their separates by the transactions' splits, moved after the split (onto `tr`, whose steps
// follow them).
function rejoinCodes(transactions: readonly Transaction[], tr: Transaction): void {
	const ranges: number[] = []

	for (const [index, step] of transactions.entries()) {
		if (!step.docChanged || isHistoryTransaction(step) || isInternal(step) || !splits(step)) continue

		const later = transactions.slice(index + 1)

		for (const at of changedRanges(step, false)) ranges.push(tr.mapping.map(later.reduce((pos, next) => next.mapping.map(pos), at)))
	}

	if (ranges.length === 0) return

	const doc = tr.doc
	const moves: { from: number; to: number; into: number }[] = []
	const seen = new Set<number>()

	for (let k = 0; k < ranges.length; k += 2) {
		const from = Math.max(0, Math.min(ranges[k] ?? 0, ranges[k + 1] ?? 0) - 1)
		const to = Math.min(doc.content.size, Math.max(ranges[k] ?? 0, ranges[k + 1] ?? 0) + 1)

		doc.nodesBetween(from, to, (node, pos, parent, index) => {
			if (!node.isTextblock) return !node.isLeaf
			if (seen.has(pos) || node.type !== nodes.paragraph) return false

			seen.add(pos)

			const offset = openCode(node)
			const next = parent !== null && index + 1 < parent.childCount ? parent.child(index + 1) : null

			if (offset >= 0 && next?.type === nodes.paragraph)
				moves.push({ from: pos + 1 + offset, to: pos + node.nodeSize - 1, into: pos + node.nodeSize + 1 })

			return false
		})
	}

	for (const { from, to, into } of moves.sort((a, b) => b.from - a.from)) {
		const code = tr.doc.slice(from, to).content

		tr.insert(into, code)
		tr.delete(from, to)
	}
}

const DELETING = [docxSchema.marks.deletion, docxSchema.marks.move_from]

// Field characters the transactions' tracked deletions marked without the rest of their field: unmarked
// (onto `tr`, whose steps follow them).
function wholeDeletions(transactions: readonly Transaction[], tr: Transaction): void {
	const ranges: number[] = []

	transactions.forEach((step, index) => {
		const later = transactions.slice(index + 1)

		step.steps.forEach((one, k) => {
			if (
				!(one instanceof AddMarkStep) ||
				!(DELETING.includes(one.mark.type) || (one instanceof SetMarkSetsStep && one.types.some(type => DELETING.includes(type))))
			)
				return

			// From after the step to the end of the transactions, then through `tr`.
			const map = (pos: number, assoc: number): number =>
				tr.mapping.map(
					later.reduce((at, next) => next.mapping.map(at, assoc), step.mapping.slice(k + 1).map(pos, assoc)),
					assoc
				)

			ranges.push(map(one.from, -1), map(one.to, 1))
		})
	})

	if (ranges.length === 0) return

	const doc = tr.doc
	let any = false

	for (let k = 0; k < ranges.length && !any; k += 2) {
		const from = Math.max(0, Math.min(ranges[k] ?? 0, doc.content.size))
		const to = Math.max(from, Math.min(ranges[k + 1] ?? 0, doc.content.size))

		doc.nodesBetween(from, to, node => {
			if (node.type === nodes.field_char) any = true

			return !any && !node.isLeaf
		})
	}

	if (!any) return

	const deleted = (node: Node): boolean => DELETING.some(type => type.isInSet(node.marks) !== undefined)
	const open: { pos: number; deleted: boolean }[][] = []
	const mixed: { pos: number; deleted: boolean }[][] = []

	doc.descendants((node, pos) => {
		if (node.type !== nodes.field_char) return !node.isLeaf

		const kind = node.attrs["kind"] as string
		const char = { pos, deleted: deleted(node) }

		if (kind === "begin") open.push([char])
		else open[open.length - 1]?.push(char)

		if (kind === "end") {
			const group = open.pop()

			if (group !== undefined && group.some(entry => entry.deleted) && group.some(entry => !entry.deleted)) mixed.push(group)
		}

		return false
	})

	for (const group of mixed)
		for (const char of group) if (char.deleted) for (const type of DELETING) tr.removeMark(char.pos, char.pos + 1, type)
}

function repair(transactions: readonly Transaction[], state: EditorState): Transaction | null {
	const tr = repairCharacters(transactions, state)

	rejoinCodes(transactions, tr)
	wholeDeletions(transactions, tr)

	return tr.docChanged ? markInternal(tr) : null
}

// Characters an edit left without partners, and the rest of fields it took one of (see the top).
function repairCharacters(transactions: readonly Transaction[], state: EditorState): Transaction {
	const removed: number[] = []
	let any = false

	transactions.forEach((tr, index) => {
		if (!tr.docChanged || isHistoryTransaction(tr) || isInternal(tr)) return

		const points = touched(tr)

		if (points === undefined) return

		any = true

		// Carried through the transactions after it.
		const later = transactions.slice(index + 1)

		for (const point of points) removed.push(later.reduce((at, next) => next.mapping.map(at, -1), point))
	})

	if (!any) return state.tr

	const doc = state.doc
	// Characters without partners before these edits (a file's own) stay as they are.
	const first = transactions[0]
	const kept = new Set<number>()

	if (first !== undefined)
		for (const char of unmatched(first.before)) {
			let pos: number | null = char.pos

			for (const tr of transactions) {
				const mapped: { deleted: boolean; pos: number } | null = pos === null ? null : tr.mapping.mapResult(pos, 1)

				pos = mapped === null || mapped.deleted ? null : mapped.pos
			}

			if (pos !== null) kept.add(pos)
		}

	const stray = unmatched(doc).filter(char => !kept.has(char.pos))
	const tr = state.tr

	if (stray.length === 0) return tr

	// The rest of a field one of whose characters an edit deleted: from its begin to where its end was, or from
	// where its begin was to its end.
	const open = stray.filter(char => char.kind === "begin" || char.kind === "end")
	const char = open.length === 1 ? open[0] : undefined
	let range: [number, number] | undefined

	if (char !== undefined && removed.length > 0) {
		const after = Math.min(...removed.filter(point => point > char.pos))
		const before = Math.max(...removed.filter(point => point <= char.pos))

		range =
			char.kind === "begin"
				? Number.isFinite(after)
					? [char.pos, after]
					: undefined
				: Number.isFinite(before)
					? [before, char.pos + 1]
					: undefined

		if (range !== undefined && !deletable(doc, range[0], range[1])) range = undefined
	}

	if (range !== undefined) tr.delete(range[0], range[1])

	// What is still without partners goes alone.
	for (const left of unmatched(tr.doc).reverse()) {
		const before = tr.mapping.invert().map(left.pos, 1)

		if (!kept.has(before)) tr.delete(left.pos, left.pos + 1)
	}

	return tr
}

// Field characters in [from, to) of `doc`, added to `out`.
function charsIn(doc: Node, from: number, to: number, out: Set<Node>): void {
	if (to > from)
		doc.nodesBetween(from, to, node => {
			if (node.type === nodes.field_char) out.add(node)

			return !node.isLeaf
		})
}

// A move (field characters taken out and put back, the same nodes: a drag, a block moved) that leaves a field
// spanning paragraphs without its partners is refused rather than repaired: Word moves a field whole (a
// selection taking one of its boundaries takes all of it), and dropping its characters would lose the field.
function brokenMove(tr: Transaction): string | undefined {
	if (!tr.docChanged || isHistoryTransaction(tr) || isInternal(tr)) return undefined

	const gone = new Set<Node>()
	const put = new Set<Node>()

	tr.steps.forEach((step, index) => {
		if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep)) return

		const doc = tr.docs[index] as Node

		if (step instanceof ReplaceAroundStep) {
			charsIn(doc, step.from, step.gapFrom, gone)
			charsIn(doc, step.gapTo, step.to, gone)
		} else {
			charsIn(doc, step.from, step.to, gone)
		}

		step.slice.content.descendants(node => {
			if (node.type === nodes.field_char) put.add(node)

			return !node.isLeaf
		})
	})

	if (![...put].some(node => gone.has(node)) || unmatched(tr.doc).length <= unmatched(tr.before).length) return undefined

	return "A move would take a field spanning paragraphs apart"
}

export const FIELDS_STAGE: InvariantStage | null = { name: "fields", violation: brokenMove, repair }
