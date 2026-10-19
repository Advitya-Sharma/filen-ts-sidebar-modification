import type { Mark, Node, NodeType, Slice } from "prosemirror-model"
import { isHistoryTransaction } from "prosemirror-history"
import type { EditorState, Transaction } from "prosemirror-state"
import { AttrStep, ReplaceAroundStep, ReplaceStep } from "prosemirror-transform"
import { RevisionFlag } from "../../model/index"
import { SetAttrsStep } from "../attrs"
import { NO_SOURCE, docxSchema, tableEdge } from "../schema"
import { cellOverrides, rowOverrides } from "../tables/map"
import { planOf, type RevisionStamp } from "../../review/track"
import { changedRanges, isCommand, isInternal, markInternal, orderedResolve, positionBefore, replacesOf } from "./core"
import type { InvariantStage } from "./index"

// Tables keep a grid every row fits and their vertical merges whole: spans, merges and grid widths change
// only through table commands, and a host's own step may not take rows or cells out of a table that stays, or
// put them into one or join two, where a vertical merge runs (table.row.delete moves the merge's start; a step
// cannot). A row wider than the grid that other edits leave (two rows a deletion joined) the writer squares
// away: the grid grows with it. A copy of a row or cell the editor recorded as a tracked change carries no
// record (its stamp stays with the original): it is no tracked change. A pasted copy of part of a table has its
// merges start where its copy starts.

const nodes = docxSchema.nodes
const marks = docxSchema.marks
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

// Whether a row whose start or end token an edit taking out [from, to) takes, or that holds such a cell, belongs
// to a table that stays (`table` at `tablePos`) and sits where a merge runs: in the row, or continuing below it.
function breaksMerge(table: Node, tablePos: number, row: Node, rowIndex: number, from: number, to: number): boolean {
	if (tablePos >= from && tablePos + table.nodeSize <= to) return false

	return merged(row, "any") || merged(table.maybeChild(rowIndex + 1), "continue")
}

// Rows or cells of a table that stays taken out of [from, to) of `doc`, or joined with a neighbour (a start or
// end token in the range), where a merge runs. The walk keeps the tables and rows around what it visits, so no
// position is resolved (a step over many rows of a long table stays linear).
function removesMergedRows(doc: Node, from: number, to: number): boolean {
	// Inside one paragraph (a key deleting text, a match Replace All replaces): no row or cell.
	const start = orderedResolve(doc, from)
	const end = orderedResolve(doc, to)

	if (start.$pos.parent.inlineContent && start.start + start.$pos.start() === end.start + end.$pos.start()) return false

	let found = false
	const tables = new Map<Node, number>()
	const rows = new Map<Node, { readonly table: Node; readonly index: number }>()

	doc.nodesBetween(from, to, (node, pos, parent, index) => {
		if (found) return false

		if (node.type === nodes.table) {
			// A table taken out whole (Replace All rebuilding it) breaks none of its merges.
			if (pos >= from && pos + node.nodeSize <= to) return false

			tables.set(node, pos)

			return true
		}

		if (node.type !== nodes.table_row && node.type !== nodes.table_cell) return !node.isInline

		const end = pos + node.nodeSize
		const opens = pos >= from && pos < to
		const closes = end - 1 >= from && end - 1 < to
		const row = node.type === nodes.table_row ? { table: parent, index } : parent === null ? undefined : rows.get(parent)

		if (node.type === nodes.table_row && parent !== null) rows.set(node, { table: parent, index })

		if (opens || closes) {
			const table = row?.table ?? null
			const tablePos = table === null ? undefined : tables.get(table)
			const rowNode = node.type === nodes.table_row ? node : parent

			if (table !== null && tablePos !== undefined && rowNode !== null && row !== undefined)
				found = breaksMerge(table, tablePos, rowNode, row.index, from, to)
		}

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
			for (const replace of replacesOf(step)) {
				if (replace.to > replace.from && removesMergedRows(doc, replace.from, replace.to)) return "tableStructure"
				if (replace.slice.size > 0 && insertsMergedRows(replace.slice, doc, replace.to)) return "tableStructure"
			}

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
	const rows: {
		start: number
		span: number
		vMerge: string | null
		pos: number
	}[][] = []
	const fresh: boolean[] = []

	table.forEach((row, rowOffset) => {
		fresh.push((row.attrs["src"] as number) === NO_SOURCE)

		const cells: {
			start: number
			span: number
			vMerge: string | null
			pos: number
		}[] = []
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

// A row or cell the end of a step's deletion lay in, which ProseMirror's fitting opens again from its type and
// attributes alone (Fitter.openFrontierNode): its node marks (properties commands set, tracked-change stamps)
// are left behind. Where it is in the state's document (inside it), and what it had.
interface Reopened {
	pos: number
	readonly type: NodeType
	readonly attrs: Node["attrs"]
	readonly marks: readonly Mark[]
}

function reopened(tr: Transaction, out: Reopened[]): void {
	tr.steps.forEach((step, index) => {
		const doc = tr.docs[index]

		if (!(step instanceof ReplaceStep) || step.to <= step.from || doc === undefined) return

		const $to = doc.resolve(step.to)

		for (let depth = 1; depth <= $to.depth; depth++) {
			const node = $to.node(depth)

			// Only one the step opened (it starts inside what the step replaces).
			if (
				(node.type !== nodes.table_row && node.type !== nodes.table_cell) ||
				node.marks.length === 0 ||
				$to.before(depth) < step.from
			)
				continue

			out.push({
				pos: tr.mapping.slice(index).map(step.to, 1),
				type: node.type,
				attrs: node.attrs,
				marks: node.marks
			})
		}
	})
}

// Junctions of tables side by side in `node`'s blocks (content from `start`) within [from, to], into `out`.
function sideBySide(node: Node, start: number, from: number, to: number, out: number[]): void {
	let previous: Node | null = null

	node.forEach((child, offset) => {
		const at = start + offset

		if (previous !== null && at >= from && at <= to && tableEdge(previous, true) && tableEdge(child, false)) out.push(at)
		if (!child.isTextblock && !child.isLeaf && at <= to && at + child.nodeSize >= from) sideBySide(child, at + 1, from, to, out)

		previous = child
	})
}

function sideBySideAt(doc: Node, pos: number): boolean {
	const $pos = doc.resolve(pos)

	return !$pos.parent.isTextblock && tableEdge($pos.nodeBefore, true) && tableEdge($pos.nodeAfter, false)
}

// Where a host's edit put tables side by side that were not (the paragraph between two deleted, a table pasted or
// dropped right after another), in `tr`'s document: Word and LibreOffice show such tables as one table, so a
// paragraph goes back between them.
function joinedTables(tr: Transaction): number[] {
	// Attributes and marks put no blocks side by side; the span every other step changed is walked once.
	const ranges = changedRanges(tr, false, NO_ATTRS)
	let from = Infinity
	let to = -Infinity

	for (let k = 0; k < ranges.length; k += 2) {
		from = Math.min(from, ranges[k] ?? 0)
		to = Math.max(to, ranges[k + 1] ?? 0)
	}

	if (from > to) return []

	const found: number[] = []

	sideBySide(tr.doc, 0, Math.max(0, from), Math.min(tr.doc.content.size, to), found)

	return found.filter(pos => !sideBySideAt(tr.before, positionBefore(tr, pos)))
}

const NO_ATTRS: ReadonlySet<string> = new Set()

// A paragraph a repair puts back at `pos` of `doc` for an edit recorded with `stamp` while tracking: tracked as
// inserted (its mark carrying the insertion, as Word records it, ECMA-376 Part 1 §17.13.5.15) by the change of the
// table the edit inserted next to it when there is one, so it goes when that change is rejected.
function separator(doc: Node, pos: number, stamp: RevisionStamp | undefined): Node {
	if (stamp === undefined) return nodes.paragraph.create()

	const $pos = doc.resolve(pos)
	const inserted = insertedBy($pos.nodeAfter, true) ?? insertedBy($pos.nodeBefore, false) ?? stamp

	return nodes.paragraph.create({ revision: RevisionFlag.Inserted }, null, [marks.paragraph_change.create({ value: { inserted } })])
}

// The tracked insertion of a table's first (or last) row.
function insertedBy(table: Node | null | undefined, first: boolean): RevisionStamp | undefined {
	const row = table?.type === nodes.table ? (first ? table.firstChild : table.lastChild) : null

	if (row === null || ((row.attrs["revision"] as number) & RevisionFlag.Inserted) === 0) return undefined

	return rowOverrides(row).ins
}

function repair(transactions: readonly Transaction[], state: EditorState): Transaction | null {
	let fix: Transaction | null = null
	const size = state.doc.content.size
	const lost: Reopened[] = []
	const merges: { pos: number; vMerge: string | null }[] = []
	const healed = new Set<number>()

	// What the transactions changed, mapped to the state's document, in order and merged; an edit of many blocks
	// (a property of every paragraph) has a range per block, so past a few their whole span is read once.
	const spans: [number, number][] = []

	const joins = new Map<number, RevisionStamp | undefined>()
	let endsWithTable: { stamp: RevisionStamp | undefined } | undefined

	transactions.forEach((tr, index) => {
		if (!tr.docChanged || isInternal(tr) || isCommand(tr) || isHistoryTransaction(tr) || withinTextblocks(tr)) return

		const stamp = planOf(tr)?.stamp

		// The story's last paragraph taken from after its last table: Word always keeps a paragraph there.
		if (tableEdge(tr.doc.lastChild, true) && !tableEdge(tr.before.lastChild, true)) endsWithTable = { stamp }

		for (let pos of joinedTables(tr)) {
			for (let later = index + 1; later < transactions.length; later++) pos = transactions[later]?.mapping.map(pos, 1) ?? pos

			joins.set(pos, stamp)
		}

		const from = lost.length

		reopened(tr, lost)

		for (let k = from; k < lost.length; k++) {
			const entry = lost[k]

			if (entry !== undefined)
				for (let later = index + 1; later < transactions.length; later++)
					entry.pos = transactions[later]?.mapping.map(entry.pos, 1) ?? entry.pos
		}

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

	// Their marks back, on the node they opened again (its type and attributes as they were, no marks).
	const restored = new Set<number>()

	for (const entry of lost) {
		if (entry.pos < 0 || entry.pos > size) continue

		const $pos = state.doc.resolve(entry.pos)

		for (let depth = $pos.depth; depth > 0; depth--) {
			const node = $pos.node(depth)

			if (!node.hasMarkup(entry.type, entry.attrs)) continue

			const at = $pos.before(depth)

			fix ??= state.tr

			for (const mark of entry.marks) fix.addNodeMark(at, mark)

			restored.add(at)

			break
		}
	}

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

			if (bits === 0 || (node.attrs["src"] as number) !== NO_SOURCE || restored.has(pos)) return true

			const overrides = node.type === nodes.table_row ? rowOverrides(node) : cellOverrides(node)
			let kept = bits

			for (const [flag, key] of TRACKED) if ((bits & flag) !== 0 && overrides[key] === undefined) kept &= ~flag

			if (kept !== bits) (fix ??= state.tr).setNodeAttribute(pos, "revision", kept)

			return true
		})
	}

	for (const merge of merges) (fix ??= state.tr).setNodeAttribute(merge.pos, "vMerge", merge.vMerge)

	// Last, from the end: a paragraph between tables an edit put side by side.
	for (const [pos, stamp] of [...joins].sort((a, b) => b[0] - a[0])) {
		if (pos < 0 || pos > size || !sideBySideAt(state.doc, pos)) continue

		fix ??= state.tr
		fix.insert(fix.mapping.map(pos), separator(state.doc, pos, stamp))
	}

	// A story an edit left ending with a table gets its last paragraph back, as Word always keeps one there
	// (LibreOffice 26.8's text export of such a body never ends).
	if (endsWithTable !== undefined && tableEdge(state.doc.lastChild, true)) {
		fix ??= state.tr
		fix.insert(fix.mapping.map(size), separator(state.doc, size, endsWithTable.stamp))
	}

	return fix === null ? null : markInternal(fix)
}

export const TABLES_STAGE: InvariantStage | null = {
	name: "tables",
	violation,
	repair
}
