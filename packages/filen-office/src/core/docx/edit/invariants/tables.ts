import type { Node, Slice } from "prosemirror-model"
import { isHistoryTransaction } from "prosemirror-history"
import type { EditorState, Transaction } from "prosemirror-state"
import { AttrStep, ReplaceAroundStep, ReplaceStep } from "prosemirror-transform"
import { RevisionFlag } from "../../model/index"
import { SetAttrsStep } from "../attrs"
import { NO_SOURCE, docxSchema } from "../schema"
import { cellOverrides, rowOverrides } from "../tables/map"
import { changedRanges, isCommand, isInternal, markInternal } from "./core"
import type { InvariantStage } from "./index"

// Tables keep a grid every row fits and their vertical merges whole: spans, merges and grid widths change
// only through table commands, and a host's own step may not take rows or cells out of a table that stays, or
// put them into one or join two, where a vertical merge runs (table.row.delete moves the merge's start; a step
// cannot). A row wider than the grid that other edits leave (two rows a deletion joined) the writer squares
// away: the grid grows with it. A copy of a row or cell the editor recorded as a tracked change carries no
// record (its stamp stays with the original): it is no tracked change. A pasted copy of part of a table has its
// merges start where its copy starts.

const nodes = docxSchema.nodes
const STRUCTURE = ["colspan", "vMerge", "grid"]
const containers = new Set([nodes.table, nodes.table_cell])

function merged(row: Node | null | undefined, which: "any" | "continue"): boolean {
	if (row === null || row === undefined || row.type !== nodes.table_row) return false

	let found = false

	row.forEach(cell => {
		const vMerge = cell.attrs["vMerge"] as string | null

		if (vMerge !== null && (which === "any" || vMerge === "continue")) found = true
	})

	return found
}

// Whether a row or cell (`node`, at `pos` of `doc`) whose start or end token an edit takes out belongs to a
// table that stays and sits where a merge runs: in its row (or the row itself), or continuing below it.
function breaksMerge(doc: Node, node: Node, pos: number, from: number, to: number): boolean {
	const $node = doc.resolve(pos)
	let depth = $node.depth

	while (depth > 0 && $node.node(depth).type !== nodes.table) depth--

	const table = $node.node(depth)
	const tablePos = depth > 0 ? $node.before(depth) : -1

	if (tablePos < 0 || (tablePos >= from && tablePos + table.nodeSize <= to)) return false

	const row = node.type === nodes.table_row ? node : $node.parent
	const rowEnd = node.type === nodes.table_row ? pos + node.nodeSize : $node.after()

	return merged(row, "any") || merged(doc.nodeAt(rowEnd), "continue")
}

// Rows or cells of a table that stays taken out of [from, to) of `doc`, or joined with a neighbour (a start or
// end token in the range), where a merge runs.
function removesMergedRows(doc: Node, from: number, to: number): boolean {
	// Inside one paragraph (a key deleting text): no row or cell (the step resolved both ends already).
	const $from = doc.resolve(from)

	if ($from.parent.inlineContent && $from.sameParent(doc.resolve(to))) return false

	let found = false

	doc.nodesBetween(from, to, (node, pos) => {
		if (found) return false
		if (node.type !== nodes.table_row && node.type !== nodes.table_cell) return !node.isInline

		const end = pos + node.nodeSize
		const opens = pos >= from && pos < to
		const closes = end - 1 >= from && end - 1 < to

		if ((opens || closes) && breaksMerge(doc, node, pos, from, to)) found = true

		return !found && !(opens && closes)
	})

	return found
}

// Rows or cells a slice puts into a table (not inside a whole table of its own, or a whole row for cells),
// where a merge runs: in what it puts in, in the row it goes into or continuing below.
function insertsMergedRows(slice: Slice, doc: Node, at: number): boolean {
	let rows = false
	let cells = false
	let found = false
	const visit = (node: Node): boolean => {
		if (node.type === nodes.table) return false
		if (node.type === nodes.table_row) {
			rows = true
			found ||= merged(node, "any")

			return false
		}

		if (node.type !== nodes.table_cell) return !node.isInline

		cells = true
		found ||= node.attrs["vMerge"] !== null

		return false
	}

	slice.content.forEach(node => {
		if (visit(node)) node.descendants(visit)
	})

	if (found) return true

	if (!rows && !cells) return false

	const $at = doc.resolve(at)

	// Into a row (rows put there split it): the row and the one below. Between rows: the one below.
	for (let depth = $at.depth; depth > 0; depth--) {
		const node = $at.node(depth)

		if (node.type === nodes.table_row) return merged(node, "any") || merged(doc.nodeAt($at.after(depth)), "continue")
		if (node.type === nodes.table) return rows && merged($at.nodeAfter, "continue")
	}

	return false
}

function violation(tr: Transaction): string | undefined {
	if (isCommand(tr) || isHistoryTransaction(tr)) return undefined

	for (let index = 0; index < tr.steps.length; index++) {
		const step = tr.steps[index]
		const doc = tr.docs[index]

		if (step === undefined || doc === undefined) continue
		if (step instanceof AttrStep) {
			if (STRUCTURE.includes(step.attr)) return "tableStructure"

			continue
		}

		if (step instanceof SetAttrsStep) {
			if (step.entries.some(entry => Object.keys(entry.attrs).some(name => STRUCTURE.includes(name)))) return "tableStructure"

			continue
		}

		if (step instanceof ReplaceStep) {
			if (step.to > step.from && removesMergedRows(doc, step.from, step.to)) return "tableStructure"
			if (step.slice.size > 0 && insertsMergedRows(step.slice, doc, step.to)) return "tableStructure"

			continue
		}

		if (!(step instanceof ReplaceAroundStep)) continue

		// setNodeMarkup: the same node around the same content with other attributes.
		if (step.insert === 1 && step.gapFrom === step.from + 1 && step.gapTo === step.to - 1) {
			const before = doc.nodeAt(step.from)
			const after = step.slice.content.firstChild

			if (
				before !== null &&
				after !== null &&
				containers.has(before.type) &&
				STRUCTURE.some(name => name in before.attrs && before.attrs[name] !== after.attrs[name])
			)
				return "tableStructure"

			continue
		}

		if (removesMergedRows(doc, step.from, step.gapFrom) || removesMergedRows(doc, step.gapTo, step.to)) return "tableStructure"
	}

	return undefined
}

const TRACKED: readonly [number, "ins" | "del"][] = [
	[RevisionFlag.Inserted, "ins"],
	[RevisionFlag.Deleted, "del"]
]

// The vertical merges of new rows (a pasted copy of part of a table, a table of its own or rows that went into
// another: new rows have no grid columns before their cells) that continue under no merge: such a cell starts
// the merge when the cell below continues it, and is no merge when nothing does.
function healMerges(table: Node, tablePos: number, fixes: { pos: number; vMerge: string | null }[]): void {
	const rows: { start: number; span: number; vMerge: string | null; pos: number }[][] = []
	const fresh: boolean[] = []

	table.forEach((row, rowOffset) => {
		fresh.push((row.attrs["src"] as number) === NO_SOURCE)

		const cells: { start: number; span: number; vMerge: string | null; pos: number }[] = []
		let column = 0

		row.forEach((cell, cellOffset) => {
			const span = Math.max(1, cell.attrs["colspan"] as number)

			cells.push({
				start: column,
				span,
				vMerge: cell.attrs["vMerge"] as string | null,
				pos: tablePos + 1 + rowOffset + 1 + cellOffset
			})
			column += span
		})
		rows.push(cells)
	})

	const at = (r: number, start: number, span: number) => rows[r]?.find(cell => cell.start === start && cell.span === span)

	rows.forEach((cells, r) => {
		if (fresh[r] !== true) return

		for (const cell of cells) {
			if (cell.vMerge !== "continue") continue

			const above = r === 0 ? undefined : at(r - 1, cell.start, cell.span)

			if (above !== undefined && above.vMerge !== null) continue

			cell.vMerge = at(r + 1, cell.start, cell.span)?.vMerge === "continue" ? "restart" : null
			fixes.push({ pos: cell.pos, vMerge: cell.vMerge })
		}
	})
}

// Whether every step of `tr` replaced content inside one textblock with content holding no table (typing, a
// split): no table structure changes. Positions a step applied at are resolved already (cached).
function withinTextblocks(tr: Transaction): boolean {
	return tr.steps.every((step, index) => {
		const doc = tr.docs[index]

		if (!(step instanceof ReplaceStep) || doc === undefined) return false

		const $from = doc.resolve(step.from)

		if (!$from.parent.inlineContent || (step.to > step.from && !$from.sameParent(doc.resolve(step.to)))) return false

		let tables = false

		step.slice.content.descendants(node => {
			if (node.type === nodes.table || node.type === nodes.table_row || node.type === nodes.table_cell) tables = true

			return !tables && !node.isTextblock
		})

		return !tables
	})
}

function repair(transactions: readonly Transaction[], state: EditorState): Transaction | null {
	let fix: Transaction | null = null
	const size = state.doc.content.size
	const merges: { pos: number; vMerge: string | null }[] = []
	const healed = new Set<number>()

	// What the transactions changed, mapped to the state's document, in order and merged; an edit of many blocks
	// (a property of every paragraph) has a range per block, so past a few their whole span is read once.
	const spans: [number, number][] = []

	transactions.forEach((tr, index) => {
		if (!tr.docChanged || isInternal(tr) || isCommand(tr) || isHistoryTransaction(tr) || withinTextblocks(tr)) return

		// Inline marks change no table structure.
		const ranges = changedRanges(tr, false)

		// Outward: a later transaction that rebuilds whole blocks (the tracking stage's bulk marking) takes the
		// range in, where inward mapping would collapse it.
		for (let k = 0; k < ranges.length; k += 2) {
			let from = ranges[k] ?? 0
			let to = ranges[k + 1] ?? 0

			for (let later = index + 1; later < transactions.length; later++) {
				from = transactions[later]?.mapping.map(from, -1) ?? from
				to = transactions[later]?.mapping.map(to, 1) ?? to
			}

			if (to > from) spans.push([Math.max(0, from), Math.min(size, to)])
		}
	})

	spans.sort((a, b) => a[0] - b[0])

	const merged: [number, number][] = []

	for (const span of spans) {
		const last = merged[merged.length - 1]

		if (last !== undefined && span[0] <= last[1]) last[1] = Math.max(last[1], span[1])
		else merged.push([span[0], span[1]])
	}

	const first = merged[0]
	const last = merged[merged.length - 1]
	const walks = merged.length > 64 && first !== undefined && last !== undefined ? [[first[0], last[1]]] : merged

	for (const [from = 0, to = 0] of walks) {
		state.doc.nodesBetween(from, to, (node, pos) => {
			if (node.isTextblock || node.isInline) return false

			if (node.type === nodes.table && pos >= from && (node.attrs["src"] as number) === NO_SOURCE && !healed.has(pos)) {
				healed.add(pos)
				healMerges(node, pos, merges)
			}

			if (pos < from || (node.type !== nodes.table_row && node.type !== nodes.table_cell)) return true

			// A new row in a table that stays (rows pasted into it): its merges are healed too.
			if (node.type === nodes.table_row && (node.attrs["src"] as number) === NO_SOURCE) {
				const $row = state.doc.resolve(pos)
				const tablePos = $row.before()

				if (!healed.has(tablePos)) {
					healed.add(tablePos)
					healMerges($row.parent, tablePos, merges)
				}
			}

			const bits = node.attrs["revision"] as number

			if (bits === 0 || (node.attrs["src"] as number) !== NO_SOURCE) return true

			const overrides = node.type === nodes.table_row ? rowOverrides(node) : cellOverrides(node)
			let kept = bits

			for (const [flag, key] of TRACKED) if ((bits & flag) !== 0 && overrides[key] === undefined) kept &= ~flag

			if (kept !== bits) (fix ??= state.tr).setNodeAttribute(pos, "revision", kept)

			return true
		})
	}

	for (const merge of merges) (fix ??= state.tr).setNodeAttribute(merge.pos, "vMerge", merge.vMerge)

	return fix === null ? null : markInternal(fix)
}

export const TABLES_STAGE: InvariantStage | null = { name: "tables", violation, repair }
