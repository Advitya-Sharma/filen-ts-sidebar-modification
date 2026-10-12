import type { Node, Slice } from "prosemirror-model"
import { isHistoryTransaction } from "prosemirror-history"
import type { EditorState, Transaction } from "prosemirror-state"
import { ReplaceAroundStep, ReplaceStep } from "prosemirror-transform"
import { NO_SOURCE, docxSchema } from "../schema"
import { isInternal, markInternal } from "./core"
import type { InvariantStage } from "./index"

// Fields spanning paragraphs (a table of contents the editor made: field_char begin, separate and end in
// different paragraphs) stay whole: an edit that deletes one of their characters deletes the field from its
// begin to its end, as Word deletes a field a selection takes a boundary of; characters an edit leaves without
// their partners otherwise (a paste of part of a table) go, and the rest is plain text. Only edits whose steps
// remove or insert a field character are looked at.

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
	const open: { begin: Char; separate: Char | undefined }[] = []
	const stray: Char[] = []

	doc.descendants((node, pos) => {
		if (node.type !== nodes.field_char) return !node.isLeaf

		const char = { pos, kind: node.attrs["kind"] as string }
		const top = open[open.length - 1]

		if (char.kind === "begin") open.push({ begin: char, separate: undefined })
		else if (char.kind === "separate" && top !== undefined && top.separate === undefined) top.separate = char
		else if (char.kind === "end" && top !== undefined) open.pop()
		else stray.push(char)

		return false
	})

	for (const field of open) stray.push(field.begin, ...(field.separate === undefined ? [] : [field.separate]))

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

function repair(transactions: readonly Transaction[], state: EditorState): Transaction | null {
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

	if (!any) return null

	const doc = state.doc
	const stray = unmatched(doc)

	if (stray.length === 0) return null

	// The rest of a field one of whose characters an edit deleted: from its begin to where its end was, or from
	// where its begin was to its end.
	const tr = state.tr
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
	for (const left of unmatched(tr.doc).reverse()) tr.delete(left.pos, left.pos + 1)

	return tr.docChanged ? markInternal(tr) : null
}

export const FIELDS_STAGE: InvariantStage | null = { name: "fields", repair }
