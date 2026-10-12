import { Fragment, type Mark, type MarkType, type Node } from "prosemirror-model"
import { closeHistory } from "prosemirror-history"
import type { EditorState, Transaction } from "prosemirror-state"
import type { DocxEditContext } from "../edit/context"
import { tableMap } from "../edit/tables/map"
import { clearCellRevision, clearRowRevision, removeCells, removeRows } from "../edit/tables/ops"
import type { PropsStore } from "../props/index"
import { formattingMarks } from "../edit/formatting"
import { markInternal, reinsert } from "../edit/invariants/core"
import { NO_MARKERS, NO_SOURCE, docxSchema, survivorsOf, type MarkerAttrs, type MarkerList } from "../edit/schema"
import type { DocxEditorSession } from "../edit/session"
import { Ns } from "../model/namespaces"
import { partReader } from "../model/reader"
import type { XmlDocument } from "../../xml/index"
import { Revision, RevisionFlag, type StoryPart } from "../model/index"
import { ReviewKind, revisionMarkId, type ReviewEntry, type ReviewIndex } from "./index"
import { paragraphChange, propertyMarks, setParagraphChange, withParagraphChange, type ParagraphAttrs, type ParagraphChange } from "./track"

// Accepting and rejecting tracked changes, as one internal transaction of the story (one undo step): an
// accepted insertion or rejected deletion keeps its content and loses the mark, an accepted deletion or
// rejected insertion removes its content (range markers and comment anchors in it stay), a formatting or
// paragraph property change keeps (accept) or restores (reject) the properties, an inserted or deleted
// paragraph mark stays or joins its paragraph with the next one. Moves go as pairs with their range markers.
// Rows and cells inserted or deleted go or stay the same way. Changes inside locked blocks, property changes of
// tables, rows, cells and sections, and paragraph marks that cannot join (a section break, no paragraph after)
// are skipped and counted.

const nodes = docxSchema.nodes
const marks = docxSchema.marks

// What to resolve: every change of the story, those overlapping a range (an empty range: the change at that
// position), or those with an id (as the review index has it).
export type ReviewScope = { readonly all: true } | { readonly from: number; readonly to: number } | { readonly id: string }

export interface ReviewOutcome {
	readonly tr: Transaction | null
	// Changes resolved, and those left as they are (inside locked blocks, property changes of tables and sections,
	// cells a vertical merge holds, marks that cannot join).
	readonly resolved: number
	readonly skipped: number
}

const INSERTED: ReadonlySet<MarkType> = new Set([marks.insertion, marks.move_to])
const DELETED: ReadonlySet<MarkType> = new Set([marks.deletion, marks.move_from])
const MOVE_MARKERS: ReadonlySet<string> = new Set(["moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"])

interface Plan {
	readonly removeMarks: { from: number; to: number; mark: Mark }[]
	readonly addMarks: { from: number; to: number; mark: Mark }[]
	readonly deletes: [number, number][]
	readonly attrs: { pos: number; name: string; value: unknown }[]
	readonly changes: { pos: number; change: ParagraphChange }[]
	// Paragraphs (by position) to join with the next one.
	readonly joins: number[]
	// Move range markers to take out: inline ones by position, listed ones by holder.
	readonly markers: number[]
	readonly lists: { pos: number; name: string; list: MarkerList }[]
	// Table rows (by the table's position and the row's index) whose tracked insertion or deletion goes:
	// cleared (the row stays) or removed.
	readonly rows: { table: number; row: number; flag: number; remove: boolean }[]
	// The same for cells (by their index in the row).
	readonly cells: { table: number; row: number; index: number; flag: number; remove: boolean; key: string }[]
	// What outlives the content it is in (range markers, comment anchors), and the markers before paragraphs
	// joined to the one before them: comments all of whose anchors are among them go (accepting a deletion).
	readonly outliving: { from: number; to: number; markers: MarkerList }[]
	readonly joined: MarkerAttrs[]
	readonly dropped: Set<string>
	resolved: Set<string>
	skipped: Set<string>
}

// The ranges a scope covers and the test a change must pass.
interface Target {
	readonly ranges: readonly [number, number][]
	readonly id: string | undefined
}

function targetOf(state: EditorState, index: ReviewIndex, scope: ReviewScope): Target | undefined {
	const size = state.doc.content.size

	if ("all" in scope) return { ranges: [[0, size]], id: undefined }

	const byId = (id: string): Target | undefined => {
		const ranges = bridged(state.doc, index.changeRanges(id))

		return ranges.length === 0 ? undefined : { ranges, id }
	}

	if ("id" in scope) return byId(scope.id)
	if (scope.to > scope.from) return { ranges: [[scope.from, scope.to]], id: undefined }

	// The change at a position: the whole change (every piece with its id, in any paragraph).
	const entry = index
		.between(Math.max(0, scope.from - 1), scope.from + 1)
		.find(
			candidate =>
				candidate.kind > ReviewKind.COMMENT_REFERENCE &&
				candidate.from <= scope.from &&
				Math.max(candidate.to, candidate.from + 1) >= scope.from
		)

	if (entry === undefined) return undefined

	// A change without an id of its own (neither the file's nor the editor's) is the range itself.
	return entry.id === "#-1" || entry.id.startsWith("new:")
		? { ranges: [[entry.from, Math.max(entry.to, entry.from + 1)]], id: entry.id }
		: byId(entry.id)
}

// Ranges of a change in document order, those with only anchors (range markers, comment references) between
// them in one paragraph joined: what stands between pieces of one change is inside it. One walk.
function bridged(doc: Node, ranges: [number, number][]): [number, number][] {
	const first = ranges[0]
	const last = ranges[ranges.length - 1]

	if (first === undefined || last === undefined || ranges.length < 2) return ranges

	// The gap (by the range before it) holding `pos`, or -1.
	const gapAt = (pos: number): number => {
		let low = 0
		let high = ranges.length

		while (low < high) {
			const middle = (low + high) >> 1

			if ((ranges[middle]?.[0] ?? 0) <= pos) low = middle + 1
			else high = middle
		}

		const k = low - 1

		return k >= 0 && k + 1 < ranges.length && pos >= (ranges[k]?.[1] ?? 0) ? k : -1
	}
	// Gaps holding anything but anchors: other content, or a block's start or end.
	const broken = new Set<number>()

	doc.nodesBetween(first[0], last[1], (node, pos) => {
		if (node.isInline) {
			if (!anchorOnly(node)) broken.add(gapAt(pos))

			return false
		}

		broken.add(gapAt(pos))
		broken.add(gapAt(pos + node.nodeSize - 1))

		return true
	})

	const out: [number, number][] = [[first[0], first[1]]]

	for (let index = 1; index < ranges.length; index++) {
		const range = ranges[index] as [number, number]
		const previous = out[out.length - 1] as [number, number]

		if (!broken.has(index - 1) && range[0] >= previous[1]) previous[1] = Math.max(previous[1], range[1])
		else out.push([range[0], range[1]])
	}

	return out
}

const skippedBy = new WeakMap<DocxEditorSession, number>()

// What the session's last resolution left (see ReviewOutcome.skipped).
export function lastSkipped(session: DocxEditorSession): number {
	return skippedBy.get(session) ?? 0
}

// The transaction resolving the changes of `scope` (null when there are none).
export function resolveChanges(session: DocxEditorSession, index: ReviewIndex, scope: ReviewScope, accept: boolean): ReviewOutcome {
	const state = session.state
	const target = targetOf(state, index, scope)

	if (target === undefined) return { tr: null, resolved: 0, skipped: 0 }

	const plan: Plan = {
		removeMarks: [],
		addMarks: [],
		deletes: [],
		attrs: [],
		changes: [],
		joins: [],
		markers: [],
		lists: [],
		rows: [],
		cells: [],
		outliving: [],
		joined: [],
		dropped: new Set(),
		resolved: new Set(),
		skipped: new Set()
	}
	const moves = "all" in scope ? [] : movePairs(session, state.doc, target.ranges)
	const seen = new Set<number>()

	for (const [from, to] of target.ranges) collect(session, state.doc, from, to, target.id, false, accept, plan, seen)

	// The other side of each move the scope touches goes with it, and the range markers of both.
	for (const [from, to] of moves) collect(session, state.doc, from, to, undefined, true, accept, plan, seen)

	// The story's final section properties (w:sectPrChange) are the section writer's.
	if ("all" in scope && sourcePropertyChanges(session.story.part).others.has(state.doc.attrs["sectPr"] as number))
		plan.skipped.add("props:story")

	if ("all" in scope) collectMoveMarkers(state.doc, 0, state.doc.content.size, plan)
	else for (const [from, to] of moves) collectMoveMarkers(state.doc, Math.max(0, from - 1), to + 1, plan)

	dropComments(plan)

	const tr = apply(state, plan, session.story.part, session.context.store)

	skippedBy.set(session, plan.skipped.size)

	return { tr, resolved: plan.resolved.size, skipped: plan.skipped.size }
}

// Whether a mark is one the scope resolves: any, the one with `id`, or (`moves`) a move's.
function matches(mark: Mark, id: string | undefined, moves = false): boolean {
	if (moves) return mark.type === marks.move_from || mark.type === marks.move_to

	return id === undefined || revisionMarkId(mark) === id
}

function collect(
	session: DocxEditorSession,
	doc: Node,
	from: number,
	to: number,
	id: string | undefined,
	moves: boolean,
	accept: boolean,
	plan: Plan,
	seen: Set<number>
): void {
	const context = session.context
	const others = sourcePropertyChanges(session.story.part).others
	// The blocks around the node visited (position and index in their parent): a row's table, a cell's row and
	// table without resolving positions (linear in the blocks before them).
	const stack: { pos: number; end: number; index: number }[] = []
	// Anchors (markers, comment references) next to removed content in the paragraph walked, with nothing else
	// between: they go with it (a comment on exactly an insertion goes when it is rejected, as in Word).
	let anchors: { from: number; to: number; markers: MarkerList }[] = []
	let afterRemoved = false
	const flush = (): void => {
		if (afterRemoved) plan.outliving.push(...anchors)

		anchors = []
		afterRemoved = false
	}
	const [start, end] = withAdjacentAnchors(doc, from, to)

	doc.nodesBetween(start, end, (node, pos, parent, index) => {
		if (!node.isInline) flush()

		while (stack.length > 0 && (stack[stack.length - 1]?.end ?? 0) <= pos) stack.pop()

		if (!node.isInline && !node.isLeaf) stack.push({ pos, end: pos + node.nodeSize, index })

		if (node.isInline || node.type === nodes.locked_block) {
			if (seen.has(pos)) return false

			seen.add(pos)

			if (node.isInline) {
				const removed = collectInline(context, node, pos, id, moves, accept, plan)

				if (anchorOnly(node) && !removed) {
					anchors.push({ from: pos, to: pos + node.nodeSize, markers: survivorsOf(node) })
				} else if (removed) {
					plan.outliving.push(...anchors)
					anchors = []
					afterRemoved = true
				} else {
					flush()
				}
			} else if (lockedChanges(session.story.part, node)) plan.skipped.add(`locked:${String(pos)}`)

			return false
		}

		const close = pos + node.nodeSize - 1
		const ends = close >= from && close < Math.max(to, from + 1)

		// A block's own changes are read where its range takes its end (once: ranges of one change may share it).
		if (!ends || seen.has(pos)) return true

		seen.add(pos)

		// Property changes of tables, rows, cells, sections and numbering: left to their writers.
		if (ends && !moves && (id === undefined || id === `#${String(node.attrs["src"])}`) && others.has(node.attrs["src"] as number))
			plan.skipped.add(`props:${String(pos)}`)

		if ((node.type === nodes.table_row || node.type === nodes.table_cell) && (node.attrs["revision"] as number) !== 0) {
			if (close >= from && close < Math.max(to, from + 1)) {
				if (moves) {
					// Rows and cells are no part of a move.
				} else if (node.type === nodes.table_row) {
					if (id === undefined || id === tableChangeId(node))
						collectRow(node, stack[stack.length - 2]?.pos ?? -1, index, accept, plan)
				} else if (id === undefined || id === tableChangeId(node)) {
					const row = stack[stack.length - 2]

					collectCell(node, stack[stack.length - 3]?.pos ?? -1, row?.index ?? -1, index, accept, plan)
				}
			}
		}

		if (node.type === nodes.paragraph && close >= from && close < Math.max(to, from + 1)) {
			// A table the resolution takes out whole stands in no join: the paragraph joins the block after it.
			let after = index + 1

			while (parent !== null && after < parent.childCount && goesWhole(parent.child(after), id, accept)) after++

			const next = parent !== null && after < parent.childCount ? parent.child(after) : null

			if (!moves || ((node.attrs["revision"] as number) & (RevisionFlag.MovedFrom | RevisionFlag.MovedTo)) !== 0)
				collectParagraph(session, node, pos, next, moves ? undefined : id, accept, plan)
		}

		return true
	})

	flush()
}

function collectInline(
	context: DocxEditContext,
	node: Node,
	pos: number,
	id: string | undefined,
	moves: boolean,
	accept: boolean,
	plan: Plan
): boolean {
	const end = pos + node.nodeSize
	let remove = false
	const removed: Mark[] = []

	for (const mark of node.marks) {
		if (!matches(mark, id, moves)) continue

		if (INSERTED.has(mark.type)) {
			plan.resolved.add(revisionMarkId(mark))

			if (accept) removed.push(mark)
			else remove = true
		} else if (DELETED.has(mark.type)) {
			plan.resolved.add(revisionMarkId(mark))

			if (accept) remove = true
			else removed.push(mark)
		} else if (mark.type === marks.format_change) {
			plan.resolved.add(revisionMarkId(mark))
			removed.push(mark)

			if (!accept) restoreFormatting(context, node, mark, pos, end, plan)
		}
	}

	// What outlives deletion (range markers, comment anchors) only loses the marks.
	if (remove && survivorsOf(node).length === 0) {
		const last = plan.deletes[plan.deletes.length - 1]

		if (last !== undefined && last[1] === pos) last[1] = end
		else plan.deletes.push([pos, end])

		return true
	}

	if (remove) {
		for (const mark of node.marks) if (INSERTED.has(mark.type) || DELETED.has(mark.type)) removed.push(mark)

		plan.outliving.push({ from: pos, to: end, markers: survivorsOf(node) })
	}

	for (const mark of removed) plan.removeMarks.push({ from: pos, to: end, mark })

	return remove
}

// Rejecting a formatting change: the formatting before it comes back.
function restoreFormatting(context: DocxEditContext, node: Node, change: Mark, from: number, to: number, plan: Plan): void {
	const base = change.attrs["base"] as number | null
	let before: readonly Mark[]

	if (base === null) {
		// Read from the file: the properties its w:rPrChange holds, written from those bytes (a run of their own:
		// properties the editor has no mark for come back too, those of the change go).
		const runMark = marks.run.isInSet(node.marks)
		const runBase = context.bases.get((runMark?.attrs["base"] as number | undefined) ?? -1)
		const inner = runBase === undefined ? -1 : previousProperties(runBase.document, runBase.node)
		const previous = inner < 0 || runBase === undefined ? undefined : context.bases.of(runBase.document, inner)

		const fallback = inner < 0 ? runBase?.props.change?.previous : undefined

		before = (previous?.marks ?? (fallback === undefined ? [] : formattingMarks(fallback))).filter(
			mark => mark.type !== marks.format_change
		)

		if (runBase !== undefined && inner >= 0)
			plan.addMarks.push({ from, to, mark: marks.run.create({ src: NO_SOURCE, base: previous?.id ?? -1 }) })
	} else {
		before = base < 0 ? [] : (context.bases.get(base)?.marks.filter(mark => mark.type !== marks.format_change) ?? [])
	}

	for (const mark of propertyMarks(node.marks)) plan.removeMarks.push({ from, to, mark })
	for (const mark of before) plan.addMarks.push({ from, to, mark })

	// The run the change replaced comes back (its source bytes with it).
	const run = change.attrs["run"] as { readonly src: number; readonly base: number } | false | null

	if (run === null) return

	const current = marks.run.isInSet(node.marks)

	if (run === false) {
		if (current !== undefined) plan.removeMarks.push({ from, to, mark: current })
	} else plan.addMarks.push({ from, to, mark: marks.run.create(run) })
}

// The w:rPr inside a w:rPr's w:rPrChange (the properties before the change), or -1.
function previousProperties(document: XmlDocument, rPr: number): number {
	const reader = partReader(document)

	for (let child = document.firstChildElement(rPr); child >= 0; child = document.nextSiblingElement(child)) {
		if (!reader.is(child, Ns.W, "rPrChange")) continue

		for (let inner = document.firstChildElement(child); inner >= 0; inner = document.nextSiblingElement(inner))
			if (reader.is(inner, Ns.W, "rPr")) return inner
	}

	return -1
}

function collectParagraph(
	session: DocxEditorSession,
	node: Node,
	pos: number,
	next: Node | null,
	id: string | undefined,
	accept: boolean,
	plan: Plan
): void {
	const src = node.attrs["src"] as number
	const own = `#${String(src)}`
	const revision = node.attrs["revision"] as number
	const recorded = paragraphChange(node)
	// By id, only the changes with that id: a recorded one by its own, the source's by its paragraph.
	const named = (stamp: { readonly id?: number } | undefined): boolean =>
		id === undefined || (stamp?.id !== undefined ? String(stamp.id) === id : src !== NO_SOURCE && id === own)
	const insert = (revision & RevisionFlag.Inserted) !== 0 && named(recorded?.inserted)
	const remove = (revision & RevisionFlag.Deleted) !== 0 && named(recorded?.deleted)
	const move = (revision & (RevisionFlag.MovedFrom | RevisionFlag.MovedTo)) !== 0 && named(undefined)
	const props = recorded?.props !== undefined && named(recorded.props)
	const source = recorded?.props === undefined && recorded?.source === undefined && named(undefined)
	let change: ParagraphChange = recorded ?? {}
	let flags = revision
	let join = false
	const key = `p:${String(pos)}`

	if (insert) {
		plan.resolved.add(`${key}:ins`)
		flags &= ~RevisionFlag.Inserted

		const { inserted: _inserted, ...rest } = change

		change = rest

		if (!accept) join = true
	}

	if (remove) {
		plan.resolved.add(`${key}:del`)
		flags &= ~RevisionFlag.Deleted

		const { deleted: _deleted, ...rest } = change

		change = rest

		if (accept) join = true
	}

	if (move) {
		plan.resolved.add(`${key}:move`)
		flags &= ~(RevisionFlag.MovedFrom | RevisionFlag.MovedTo)

		if ((revision & RevisionFlag.MovedFrom) !== 0 ? accept : !accept) join = true
	}

	if (props && recorded?.props !== undefined) {
		plan.resolved.add(`${key}:props`)

		if (!accept) setParagraphAttrs(node, pos, recorded.props.before, plan)

		const { props: _props, ...rest } = change

		change = rest
	} else if (source) {
		const previous = sourcePropertyChange(session, node)

		if (previous !== undefined) {
			plan.resolved.add(`${key}:source`)
			change = { ...change, source: accept ? "accept" : "reject" }

			if (!accept) setParagraphAttrs(node, pos, previous, plan)
		}
	}

	if (join) {
		// A mark holding a section break stays as it was; one with no paragraph after it goes with its paragraph
		// when that is left empty (see joinWithNext).
		const blocked =
			(node.attrs["sectPr"] as number) !== NO_SOURCE ||
			next === null ||
			(next.type === nodes.paragraph
				? (next.attrs["sectPr"] as number) !== NO_SOURCE
				: !("before" in next.attrs) || !emptied(node, id, accept))

		if (blocked) {
			plan.skipped.add(key)
			flags = revision

			if (recorded !== undefined)
				change = {
					...change,
					...(recorded.inserted === undefined ? {} : { inserted: recorded.inserted }),
					...(recorded.deleted === undefined ? {} : { deleted: recorded.deleted })
				}
		} else {
			plan.joins.push(pos)

			if (next.type === nodes.paragraph) plan.joined.push(...(next.attrs["before"] as MarkerList))
		}
	}

	if (flags !== revision) plan.attrs.push({ pos, name: "revision", value: flags })
	if (change !== recorded && !(recorded === undefined && Object.keys(change).length === 0)) plan.changes.push({ pos, change })
}

// [from, to) widened over the anchors right next to it in its paragraphs.
function withAdjacentAnchors(doc: Node, from: number, to: number): [number, number] {
	let start = from
	let end = to

	for (let $pos = doc.resolve(start); $pos.parent.inlineContent && $pos.nodeBefore !== null && anchorOnly($pos.nodeBefore);)
		$pos = doc.resolve((start -= $pos.nodeBefore.nodeSize))

	for (let $pos = doc.resolve(end); $pos.parent.inlineContent && $pos.nodeAfter !== null && anchorOnly($pos.nodeAfter);)
		$pos = doc.resolve((end += $pos.nodeAfter.nodeSize))

	return [start, end]
}

// An inline node that is only anchors (a range marker, a comment reference).
function anchorOnly(node: Node): boolean {
	return node.type === nodes.marker || node.type === nodes.comment_ref
}

const ANCHORS: Readonly<Record<string, number>> = { commentRangeStart: 1, commentRangeEnd: 2, commentReference: 4 }

// The change a session's last review.next or review.prev selected, while that selection stands (the caret a
// paragraph's property change leaves is inside other changes too).
const movedChanges = new WeakMap<
	DocxEditorSession,
	{ readonly doc: Node; readonly from: number; readonly to: number; readonly id: string }
>()

export function rememberMovedChange(session: DocxEditorSession, doc: Node, selection: { from: number; to: number }, id: string): void {
	movedChanges.set(session, { doc, from: selection.from, to: selection.to, id })
}

export function movedChange(
	session: DocxEditorSession,
	state: { readonly doc: Node; readonly selection: { readonly from: number; readonly to: number } } = session.state
): string | undefined {
	const moved = movedChanges.get(session)
	const { doc, selection } = state

	return moved !== undefined && moved.doc === doc && moved.from === selection.from && moved.to === selection.to ? moved.id : undefined
}

// Comments whose range start, range end and reference all lie in what resolving removes (a deletion accepted,
// an insertion rejected) go with it, as in Word (the comment shows until then; the save leaves out a comment
// without anchors): their anchors are deleted too, or left out where joined paragraphs meet.
function dropComments(plan: Plan): void {
	const found = new Map<string, number>()
	const add = (marker: MarkerAttrs): void => {
		const bit = ANCHORS[marker.kind]

		if (bit !== undefined && marker.id !== null) found.set(marker.id, (found.get(marker.id) ?? 0) | bit)
	}

	for (const { markers } of plan.outliving) for (const marker of markers) add(marker)
	for (const marker of plan.joined) add(marker)

	for (const [id, bits] of found) if (bits === 7) plan.dropped.add(id)

	if (plan.dropped.size === 0) return

	for (const { from, to, markers } of plan.outliving)
		if (markers.every(marker => droppedAnchor(plan, marker))) plan.deletes.push([from, to])
}

function droppedAnchor(plan: Plan, marker: MarkerAttrs): boolean {
	return ANCHORS[marker.kind] !== undefined && marker.id !== null && plan.dropped.has(marker.id)
}

// The markers between two paragraphs joined, standing at the join inline (the comment anchors of comments
// that go left out).
function betweenNodes(plan: Plan, list: MarkerList): Node[] {
	return list
		.filter(marker => !droppedAnchor(plan, marker))
		.map(marker =>
			marker.kind === "commentReference"
				? nodes.comment_ref.create({ src: marker.src, id: marker.id })
				: nodes.marker.create({ src: marker.src, kind: marker.kind, id: marker.id })
		)
}

// Whether resolving takes `node`, a table, out whole: every row an insertion it rejects (or a deletion it
// accepts) of the scope.
function goesWhole(node: Node, id: string | undefined, accept: boolean): boolean {
	if (node.type !== nodes.table || node.childCount === 0) return false

	const flag = accept ? RevisionFlag.Deleted : RevisionFlag.Inserted

	for (let k = 0; k < node.childCount; k++) {
		const row = node.child(k)

		if (((row.attrs["revision"] as number) & flag) === 0 || (id !== undefined && id !== tableChangeId(row))) return false
	}

	return true
}

// The id a row's or a cell's tracked insertion or deletion counts under in the review index.
function tableChangeId(node: Node): string {
	const type = node.type === nodes.table_row ? marks.row_props : marks.cell_props
	const value = type.isInSet(node.marks)?.attrs["value"] as
		{ readonly ins?: { readonly id?: number }; readonly del?: { readonly id?: number } } | undefined
	const id = value?.ins?.id ?? value?.del?.id

	return id === undefined ? `#${String(node.attrs["src"])}` : String(id)
}

// A row's tracked insertion (or deletion) accepted keeps the row, rejected takes it out (and the reverse).
function collectRow(node: Node, table: number, row: number, accept: boolean, plan: Plan): void {
	const revision = node.attrs["revision"] as number

	for (const flag of [RevisionFlag.Inserted, RevisionFlag.Deleted]) {
		if ((revision & flag) === 0) continue

		plan.resolved.add(`r:${String(table)}:${String(row)}:${String(flag)}`)
		plan.rows.push({ table, row, flag, remove: (flag === RevisionFlag.Inserted) !== accept })
	}
}

// A cell's likewise: accepted insertion or rejected deletion keeps it, the others take it out.
function collectCell(node: Node, table: number, row: number, index: number, accept: boolean, plan: Plan): void {
	const revision = node.attrs["revision"] as number

	for (const flag of [RevisionFlag.Inserted, RevisionFlag.Deleted]) {
		if ((revision & flag) === 0) continue

		const key = `c:${String(table)}:${String(row)}:${String(index)}:${String(flag)}`

		plan.resolved.add(key)
		plan.cells.push({ table, row, index, flag, remove: (flag === RevisionFlag.Inserted) !== accept, key })
	}
}

// Whether resolving leaves the paragraph without content: every child goes (an accepted deletion, a rejected
// insertion), none outlives deletion.
function emptied(node: Node, id: string | undefined, accept: boolean): boolean {
	let empty = true

	node.forEach(child => {
		if (!empty) return

		const goes = child.marks.some(mark => matches(mark, id) && (accept ? DELETED : INSERTED).has(mark.type))

		if (!goes || survivorsOf(child).length > 0) empty = false
	})

	return empty
}

function setParagraphAttrs(node: Node, pos: number, attrs: ParagraphAttrs, plan: Plan): void {
	for (const [name, value] of Object.entries(attrs)) if (node.attrs[name] !== value) plan.attrs.push({ pos, name, value })
}

// The attributes before a source paragraph's own w:pPrChange, when it has one.
function sourcePropertyChange(session: DocxEditorSession, node: Node): ParagraphAttrs | undefined {
	const src = node.attrs["src"] as number

	if (src === NO_SOURCE) return undefined

	const part = session.story.part
	const row = part.aux[src] ?? -1
	const pPr = row < 0 ? -1 : (part.paragraphs.pPr[row] ?? -1)
	const previous = pPr < 0 ? undefined : session.context.store.paragraph(part.document, pPr).change?.previous

	if (previous === undefined) return undefined

	return {
		styleId: previous.styleId ?? null,
		numId: previous.numbering?.numId ?? null,
		ilvl: previous.numbering?.ilvl ?? null,
		align: previous.alignment ?? null,
		indent: previous.indent ?? null,
		spacing: previous.spacing ?? null,
		borders: previous.borders ?? null
	}
}

// Whether a locked block's source holds tracked changes (they stay as they are).
function lockedChanges(part: StoryPart, node: Node): boolean {
	const src = node.attrs["src"] as number

	if (src === NO_SOURCE) return false

	const start = part.document.start(part.node(src))
	const end = part.document.end(part.node(src))
	const revisions = part.revisions

	for (let k = 0; k < revisions.count; k++) {
		const at = part.document.start(revisions.node[k] ?? -1)

		if (at >= start && at < end) return true
	}

	return false
}

// ---- the index ----------------------------------------------------------------------------------------------

// Source items of a part whose properties hold a tracked change: paragraphs with a w:pPrChange (resolved
// here), and what is left to the table and section writers (w:tblPrChange, w:tblGridChange, w:trPrChange,
// w:tblPrExChange, w:tcPrChange, w:cellMerge, w:sectPrChange, w:numberingChange): counted, and skipped when
// resolving.
interface SourceChanges {
	readonly paragraphs: ReadonlySet<number>
	readonly others: ReadonlySet<number>
	readonly all: ReadonlySet<number>
}

const UNRESOLVED: ReadonlySet<number> = new Set([
	Revision.TableProperties,
	Revision.TableGrid,
	Revision.RowProperties,
	Revision.TablePropertiesException,
	Revision.CellProperties,
	Revision.CellMerge,
	Revision.SectionProperties,
	Revision.Numbering
])

const sourceChangesOf = new WeakMap<StoryPart, SourceChanges>()

function sourcePropertyChanges(part: StoryPart): SourceChanges {
	let found = sourceChangesOf.get(part)

	if (found === undefined) {
		const paragraphs = new Set<number>()
		const others = new Set<number>()
		const revisions = part.revisions

		for (let k = 0; k < revisions.count; k++) {
			const kind = revisions.kind[k] ?? 0

			if (kind === Revision.ParagraphProperties) paragraphs.add(revisions.item[k] ?? -1)
			else if (UNRESOLVED.has(kind)) others.add(revisions.item[k] ?? -1)
		}

		found = { paragraphs, others, all: new Set([...paragraphs, ...others]) }
		sourceChangesOf.set(part, found)
	}

	return found
}

// The session's review index with its source paragraph property changes counted.
export function storyReviewIndex(session: DocxEditorSession): ReviewIndex {
	const index = session.reviewIndex

	if (!index.sourcesAttached) index.attachSources(sourcePropertyChanges(session.story.part).all)

	return index
}

const wanted = new WeakSet<DocxEditorSession>()

// Whether keeping the session's review index pays: changes are or were recorded, the story's part has changes
// or the document comments. Otherwise there is nothing to count (and no reason to read a long document).
export function reviewIndexWanted(session: DocxEditorSession): boolean {
	if (wanted.has(session)) return true

	const context = session.context

	if (!context.review.tracking && session.story.part.revisions.count === 0 && context.model.commentAnchors.count === 0) return false

	wanted.add(session)

	return true
}

// Keeps the session's index from now on (tracking was on, a comment was added).
export function wantReviewIndex(session: DocxEditorSession): void {
	wanted.add(session)
}

// ---- moves --------------------------------------------------------------------------------------------------

interface MoveRange {
	readonly from: number
	readonly to: number
	readonly name: string | undefined
	readonly kind: "from" | "to"
}

// The move ranges (range markers by w:id) of a document, with their names as the source writes them; markers
// stand inline or in a block's lists (before it, at its end).
function moveRanges(session: DocxEditorSession, doc: Node): MoveRange[] {
	const part = session.story.part
	const starts = new Map<string, { pos: number; name: string | undefined }>()
	const out: MoveRange[] = []
	const marker = (kind: string, id: string | null, src: number, own: string | null, pos: number): void => {
		if (id === null) return

		if (kind === "moveFromRangeStart" || kind === "moveToRangeStart") {
			const name = src === NO_SOURCE ? own : (part.reader.attr(part.node(src), Ns.W, "name") ?? null)

			starts.set(`${kind}:${id}`, { pos, name: name ?? undefined })
		} else if (kind === "moveFromRangeEnd" || kind === "moveToRangeEnd") {
			const start = starts.get(`${kind === "moveFromRangeEnd" ? "moveFromRangeStart" : "moveToRangeStart"}:${id}`)

			if (start !== undefined)
				out.push({ from: start.pos, to: pos + 1, name: start.name, kind: kind === "moveFromRangeEnd" ? "from" : "to" })
		}
	}
	const listed = (list: unknown, pos: number): void => {
		if (Array.isArray(list)) for (const entry of list as MarkerList) marker(entry.kind, entry.id, entry.src, null, pos)
	}

	const visit = (node: Node, pos: number): void => {
		if (node.type === nodes.marker) {
			marker(
				node.attrs["kind"] as string,
				node.attrs["id"] as string | null,
				node.attrs["src"] as number,
				node.attrs["name"] as string | null,
				pos
			)

			return
		}

		if (node.isInline) return

		listed(node.attrs["before"], pos)
		node.forEach((child, offset) => visit(child, pos + 1 + offset))
		listed(node.attrs["end"], pos + node.nodeSize - 1)
	}

	doc.forEach((child, offset) => visit(child, offset))
	listed(doc.attrs["end"], doc.content.size)

	return out
}

// The move ranges (both sides) of each move whose content `ranges` touch: a move is accepted or rejected whole.
function movePairs(session: DocxEditorSession, doc: Node, ranges: readonly (readonly [number, number])[]): [number, number][] {
	let moves = false

	for (const [from, to] of ranges)
		doc.nodesBetween(from, to, node => {
			if (node.isInline && (marks.move_from.isInSet(node.marks) !== undefined || marks.move_to.isInSet(node.marks) !== undefined))
				moves = true

			return !moves
		})

	if (!moves) return []

	const known = moveRanges(session, doc)
	const out: [number, number][] = []

	for (const range of known) {
		if (!ranges.some(([from, to]) => range.from < to && range.to > from)) continue

		out.push([range.from, range.to])

		for (const other of known)
			if (other !== range && other.name !== undefined && other.name === range.name) out.push([other.from, other.to])
	}

	return out
}

// Move range markers inside [from, to): a resolved move keeps none.
function collectMoveMarkers(doc: Node, from: number, to: number, plan: Plan): void {
	const size = doc.content.size
	const lists = (node: Node, pos: number): void => {
		for (const name of ["before", "end"]) {
			const list = node.attrs[name] as MarkerList | undefined

			if (list === undefined || !list.some(marker => MOVE_MARKERS.has(marker.kind))) continue

			const kept = list.filter(marker => !MOVE_MARKERS.has(marker.kind))

			plan.lists.push({ pos, name, list: kept.length === 0 ? NO_MARKERS : kept })
		}
	}

	doc.nodesBetween(Math.max(0, from), Math.min(size, to), (node, pos) => {
		if (node.type === nodes.marker) {
			if (MOVE_MARKERS.has(node.attrs["kind"] as string)) plan.markers.push(pos)

			return false
		}

		if (node.isBlock) lists(node, pos)

		return !node.isInline
	})

	if (to >= size) lists(doc, -1)
}

// ---- applying -----------------------------------------------------------------------------------------------

// Tables losing rows or cells, with the rows and the cells (outside those rows) that go.
function tableRemovals(plan: Plan): Map<number, { rows: number[]; cells: Plan["cells"] }> {
	const out = new Map<number, { rows: number[]; cells: Plan["cells"] }>()
	const of = (table: number): { rows: number[]; cells: Plan["cells"] } => {
		let entry = out.get(table)

		if (entry === undefined) {
			entry = { rows: [], cells: [] }
			out.set(table, entry)
		}

		return entry
	}

	for (const row of plan.rows) {
		if (!row.remove) continue

		const entry = of(row.table)

		if (!entry.rows.includes(row.row)) entry.rows.push(row.row)
	}

	for (const cell of plan.cells) if (cell.remove) of(cell.table).cells.push(cell)

	for (const entry of out.values()) entry.cells = entry.cells.filter(cell => !entry.rows.includes(cell.row))

	return out
}

// Takes rows, then cells, out of the table at `pos` (the whole table when every row goes). Cells that cannot go
// (a vertical merge runs through what would move) stay, counted skipped.
function removeFromTable(
	tr: Transaction,
	pos: number,
	removal: { rows: readonly number[]; cells: Plan["cells"] },
	part: StoryPart,
	store: PropsStore,
	plan: Plan
): void {
	const table = tr.doc.nodeAt(pos)

	if (table?.type !== nodes.table) return

	if (removal.rows.length >= table.childCount) {
		tr.delete(pos, pos + table.nodeSize)

		return
	}

	if (removal.rows.length > 0) removeTableRows(tr, pos, table, removal.rows, part, store)
	if (removal.cells.length === 0) return

	// Row indices after the rows above them went.
	const cells = removal.cells.map(cell => ({ row: cell.row - removal.rows.filter(row => row < cell.row).length, index: cell.index }))
	const now = tr.doc.nodeAt(pos) ?? table

	if (!removeCells(tr, tableMap(now, part, store), pos, cells)) {
		for (const cell of removal.cells) {
			plan.resolved.delete(cell.key)
			plan.skipped.add(cell.key)
		}
	}
}

// Takes rows out of the table at `pos`, from the bottom (some row stays).
function removeTableRows(tr: Transaction, pos: number, table: Node, rows: readonly number[], part: StoryPart, store: PropsStore): void {
	const sorted = [...rows].sort((a, b) => b - a)
	let bottom = (sorted[0] ?? 0) + 1
	let top = bottom - 1

	for (const row of sorted.slice(1)) {
		if (row === top - 1) {
			top = row

			continue
		}

		removeRows(tr, tableMap(tr.doc.nodeAt(pos) ?? table, part, store), pos, top, bottom)
		bottom = row + 1
		top = row
	}

	removeRows(tr, tableMap(tr.doc.nodeAt(pos) ?? table, part, store), pos, top, bottom)
}

// Plans with more operations than this are applied as one rebuild of the blocks they touch: a step per
// operation copies the document's block list for each (accepting a deletion of a long document that way runs
// out of memory).
const BULK = 64

function apply(state: EditorState, plan: Plan, part: StoryPart, store: PropsStore): Transaction | null {
	const count =
		plan.removeMarks.length +
		plan.addMarks.length +
		plan.attrs.length +
		plan.changes.length +
		plan.deletes.length +
		plan.markers.length +
		plan.joins.length +
		plan.lists.length +
		plan.rows.length +
		plan.cells.length

	if (count > BULK) return keepRanges(state.doc, applyBulk(state, plan, part, store), plan)

	const tr = state.tr

	// What moves nothing first: marks and attributes.
	for (const { from, to, mark } of plan.removeMarks) tr.removeMark(from, to, mark)
	for (const { from, to, mark } of plan.addMarks) tr.addMark(from, to, mark)
	for (const { pos, name, value } of plan.attrs) tr.setNodeAttribute(pos, name, value)

	for (const { pos, change } of plan.changes) {
		const node = tr.doc.nodeAt(pos)

		if (node !== null) setParagraphChange(tr, pos, node, change)
	}

	for (const { pos, name, list } of plan.lists) {
		if (pos < 0) tr.setDocAttribute(name, list)
		else tr.setNodeAttribute(pos, name, list)
	}

	for (const row of plan.rows)
		if (!row.remove && !plan.rows.some(other => other.remove && other.table === row.table && other.row === row.row))
			clearRowRevision(tr, row.table, row.row, row.flag)

	for (const cell of plan.cells) {
		const goes = (other: Plan["cells"][number]): boolean =>
			other.remove && other.table === cell.table && other.row === cell.row && other.index === cell.index

		if (
			!cell.remove &&
			!plan.cells.some(goes) &&
			!plan.rows.some(row => row.remove && row.table === cell.table && row.row === cell.row)
		)
			clearCellRevision(tr, cell.table, cell.row, cell.index, cell.flag)
	}

	// Then what removes content, from the end.
	const removals: { pos: number; run: () => void }[] = [
		...[...tableRemovals(plan)].map(([table, removal]) => ({
			pos: table,
			run: () => removeFromTable(tr, table, removal, part, store, plan)
		})),
		...plan.deletes.map(([from, to]) => ({ pos: from, run: () => void tr.delete(from, to) })),
		// (Move ranges overlap: a marker may be listed twice.)
		...[...new Set(plan.markers)]
			.filter(pos => !plan.deletes.some(([from, to]) => pos >= from && pos < to))
			.map(pos => ({ pos, run: () => void tr.delete(pos, pos + 1) })),
		...plan.joins.map(pos => ({
			pos,
			run: () => {
				if (!joinWithNext(tr, pos, plan)) plan.skipped.add(`p:${String(pos)}`)
			}
		}))
	]

	// A join at a paragraph's end comes before deletions inside it (higher position first).
	removals.sort((a, b) => b.pos - a.pos)

	for (const removal of removals) removal.run()

	return keepRanges(state.doc, tr.docChanged ? closeHistory(markInternal(tr)) : null, plan)
}

const KEPT_RANGES: ReadonlySet<string> = new Set([
	"bookmarkStart",
	"bookmarkEnd",
	"commentRangeStart",
	"commentRangeEnd",
	"commentReference",
	"permissionStart",
	"permissionEnd"
])

// Range markers (and comment references) of `doc` by key, with where they are.
function rangeAnchors(doc: Node): Map<string, { marker: MarkerAttrs; pos: number }> {
	const out = new Map<string, { marker: MarkerAttrs; pos: number }>()
	const add = (marker: MarkerAttrs, pos: number): void => {
		if (marker.id !== null && KEPT_RANGES.has(marker.kind)) out.set(`${marker.kind}:${marker.id}`, { marker, pos })
	}

	doc.descendants((node, pos) => {
		for (const name of ["before", "end"]) {
			const list: unknown = node.attrs[name]

			if (Array.isArray(list)) for (const marker of list as MarkerList) add(marker, name === "before" ? pos : pos + node.nodeSize)
		}

		if (node.isInline) {
			for (const marker of survivorsOf(node)) add(marker, pos)

			return false
		}

		return true
	})

	return out
}

// What resolving took out with removed rows, tables or paragraphs of a range whose other markers stay (a comment
// started in a rejected inserted row and ended after it) goes back where the content was: the range shrinks.
// Ranges that went whole (a comment its resolution drops) stay out.
function keepRanges(before: Node, tr: Transaction | null, plan: Plan): Transaction | null {
	if (tr === null) return null

	const was = rangeAnchors(before)

	if (was.size === 0) return tr

	const is = rangeAnchors(tr.doc)
	// A range by its family (comment, bookmark, permission: their ids are apart) and id.
	const range = (marker: MarkerAttrs): string => `${marker.kind.slice(0, 4)}:${marker.id ?? ""}`
	const staying = new Set([...is.values()].map(({ marker }) => range(marker)))
	const lost = new Map<number, MarkerAttrs[]>()

	for (const [key, { marker, pos }] of was) {
		const comment = marker.kind.startsWith("comment")

		if (is.has(key) || (comment && plan.dropped.has(marker.id ?? "")) || !staying.has(range(marker))) continue

		const at = tr.mapping.map(pos, -1)
		const list = lost.get(at)

		if (list === undefined) lost.set(at, [marker])
		else list.push(marker)
	}

	if (lost.size === 0) return tr

	for (const [pos, markers] of [...lost].sort((a, b) => b[0] - a[0])) reinsert(tr, pos, markers)

	return tr
}

// ---- applying a large plan: one rebuild ----------------------------------------------------------------------

// What the plan does, by position of the node it acts on.
interface Rebuild {
	readonly marks: Map<number, { readonly remove: Mark[]; readonly add: Mark[] }>
	readonly deletes: readonly [number, number][]
	readonly markers: ReadonlySet<number>
	readonly attrs: Map<number, Record<string, unknown>>
	readonly changes: Map<number, ParagraphChange>
	readonly lists: Map<number, Record<string, MarkerList>>
	readonly joins: ReadonlySet<number>
	// Revisions cleared from rows ("table:row") and cells ("table:row:cell"): the RevisionFlag bits.
	readonly rows: Map<string, number>
	readonly cells: Map<string, number>
	// Tables every row of which goes, and tables with cells to clear (their rows are rebuilt).
	readonly gone: ReadonlySet<number>
	readonly cellTables: ReadonlySet<number>
	// Every position an operation acts on, sorted (what a subtree holds is found by binary search).
	readonly positions: readonly number[]
	readonly plan: Plan
}

function rebuildOf(state: EditorState, plan: Plan, removals: Map<number, { rows: number[]; cells: Plan["cells"] }>): Rebuild {
	const positions: number[] = []
	const marks = new Map<number, { remove: Mark[]; add: Mark[] }>()
	const at = <T>(map: Map<number, T>, pos: number, make: () => T): T => {
		let entry = map.get(pos)

		if (entry === undefined) {
			entry = make()
			map.set(pos, entry)
			positions.push(pos)
		}

		return entry
	}

	for (const { from, mark } of plan.removeMarks) at(marks, from, () => ({ remove: [], add: [] })).remove.push(mark)
	for (const { from, mark } of plan.addMarks) at(marks, from, () => ({ remove: [], add: [] })).add.push(mark)

	const attrs = new Map<number, Record<string, unknown>>()

	for (const { pos, name, value } of plan.attrs) at<Record<string, unknown>>(attrs, pos, () => ({}))[name] = value

	const changes = new Map<number, ParagraphChange>()

	for (const { pos, change } of plan.changes) {
		changes.set(pos, change)
		positions.push(pos)
	}

	const lists = new Map<number, Record<string, MarkerList>>()

	for (const { pos, name, list } of plan.lists) if (pos >= 0) at<Record<string, MarkerList>>(lists, pos, () => ({}))[name] = list

	const rows = new Map<string, number>()
	const cells = new Map<string, number>()
	const gone = new Set<number>()

	for (const [table, removal] of removals) {
		const node = state.doc.nodeAt(table)

		if (node !== null && removal.rows.length >= node.childCount) gone.add(table)
	}

	const removedRow = (table: number, row: number): boolean => removals.get(table)?.rows.includes(row) === true

	for (const row of plan.rows) {
		if (row.remove || removedRow(row.table, row.row)) continue

		const key = `${String(row.table)}:${String(row.row)}`

		rows.set(key, (rows.get(key) ?? 0) | row.flag)
		positions.push(row.table)
	}

	for (const cell of plan.cells) {
		if (cell.remove || removedRow(cell.table, cell.row)) continue
		if (removals.get(cell.table)?.cells.some(other => other.row === cell.row && other.index === cell.index) === true) continue

		const key = `${String(cell.table)}:${String(cell.row)}:${String(cell.index)}`

		cells.set(key, (cells.get(key) ?? 0) | cell.flag)
		positions.push(cell.table)
	}

	for (const table of removals.keys()) positions.push(table)
	for (const pos of plan.joins) positions.push(pos)
	for (const pos of plan.markers) positions.push(pos)
	for (const [from] of plan.deletes) positions.push(from)

	positions.sort((a, b) => a - b)

	return {
		marks,
		deletes: [...plan.deletes].sort((a, b) => a[0] - b[0]),
		markers: new Set(plan.markers),
		attrs,
		changes,
		lists,
		joins: new Set(plan.joins),
		rows,
		cells,
		gone,
		cellTables: new Set([...cells.keys()].map(key => Number(key.split(":")[0]))),
		positions,
		plan
	}
}

// Whether an operation acts inside [from, to).
function touches(rebuild: Rebuild, from: number, to: number): boolean {
	const positions = rebuild.positions
	let low = 0
	let high = positions.length

	while (low < high) {
		const middle = (low + high) >> 1

		if ((positions[middle] ?? 0) < from) low = middle + 1
		else high = middle
	}

	return low < positions.length && (positions[low] ?? 0) < to
}

function deletedAt(rebuild: Rebuild, pos: number): boolean {
	const deletes = rebuild.deletes
	let low = 0
	let high = deletes.length

	while (low < high) {
		const middle = (low + high) >> 1

		if ((deletes[middle]?.[0] ?? 0) <= pos) low = middle + 1
		else high = middle
	}

	const range = deletes[low - 1]

	return range !== undefined && pos < range[1]
}

// A row or cell without its tracked insertion or deletion (`flags`), as clearRowRevision and clearCellRevision
// leave it.
function withoutRevision(node: Node, flags: number): Node {
	const type = node.type === nodes.table_row ? marks.row_props : marks.cell_props
	const value = (type.isInSet(node.marks)?.attrs["value"] ?? {}) as Record<string, unknown>
	const rest: Record<string, unknown> = {}

	for (const [key, entry] of Object.entries(value))
		if (
			entry !== undefined &&
			!(key === "ins" && (flags & RevisionFlag.Inserted) !== 0) &&
			!(key === "del" && (flags & RevisionFlag.Deleted) !== 0)
		)
			rest[key] = entry

	return node.type.create(
		{ ...node.attrs, revision: (node.attrs["revision"] as number) & ~flags },
		node.content,
		Object.keys(rest).length === 0 ? type.removeFromSet(node.marks) : type.create({ value: rest }).addToSet(node.marks)
	)
}

// The children of `parent` (content at `start`, written at `out`) rebuilt, with joined paragraphs joined.
// `row`: the table and row index when `parent` is a row; `table`: the table's position when it is a table.
// `placed`: where the tables left to the steps after the rebuild end up.
function rebuildChildren(
	rebuild: Rebuild,
	parent: Node,
	start: number,
	out: number,
	context: { readonly table?: number; readonly row?: { readonly table: number; readonly index: number } },
	placed: Map<number, number>,
	fill = true
): Node[] {
	const children: Node[] = []
	let size = 0
	// A paragraph joined with what follows it, and where it was.
	let pending: { node: Node; pos: number } | undefined

	parent.forEach((child, offset, index) => {
		const pos = start + offset

		// A table that goes: what was joined with it joins what follows.
		if (child.type === nodes.table && rebuild.gone.has(pos)) return

		let carry: MarkerList | undefined

		if (pending !== undefined && child.type !== nodes.paragraph) {
			const joined = pending

			pending = undefined

			// Left empty with no paragraph after it, a joined paragraph goes (its markers to the block after it).
			if (joined.node.content.size === 0 && parent.childCount > 1 && "before" in child.attrs) {
				carry = joined.node.attrs["before"] as MarkerList
			} else {
				rebuild.plan.skipped.add(`p:${String(joined.pos)}`)
				children.push(joined.node)
				size += joined.node.nodeSize
			}
		}

		let node =
			touches(rebuild, pos, pos + child.nodeSize) || (context.table !== undefined && rebuild.cellTables.has(context.table))
				? rebuildNode(
						rebuild,
						child,
						pos,
						out + size,
						context.table === undefined ? {} : { row: { table: context.table, index } },
						placed
					)
				: child
		const key =
			context.table !== undefined
				? `${String(context.table)}:${String(index)}`
				: context.row === undefined
					? undefined
					: `${String(context.row.table)}:${String(context.row.index)}:${String(index)}`
		const flags = key === undefined ? 0 : context.table !== undefined ? (rebuild.rows.get(key) ?? 0) : (rebuild.cells.get(key) ?? 0)

		if (flags !== 0) node = withoutRevision(node, flags)

		if (carry !== undefined && carry.length > 0)
			node = node.type.create(
				{ ...node.attrs, before: [...carry, ...(node.attrs["before"] as MarkerList)] },
				node.content,
				node.marks
			)

		if (pending !== undefined) {
			// The next paragraph keeps its properties; markers between the two stand at the join, inline.
			const joined = pending
			const between = betweenNodes(rebuild.plan, node.attrs["before"] as MarkerList)

			pending = undefined
			node = node.type.create(
				{ ...node.attrs, before: joined.node.attrs["before"] },
				joined.node.content.append(Fragment.fromArray(between)).append(node.content),
				node.marks
			)
		}

		if (node.type === nodes.paragraph && rebuild.joins.has(pos)) {
			pending = { node, pos }

			return
		}

		children.push(node)
		size += node.nodeSize
	})

	if (pending !== undefined) {
		rebuild.plan.skipped.add(`p:${String(pending.pos)}`)
		children.push(pending.node)
	}

	// A container left empty gets what its content needs (a table that went was its only block).
	if (children.length === 0 && fill) parent.type.contentMatch.fillBefore(Fragment.empty, true)?.forEach(node => children.push(node))

	return children
}

function rebuildNode(
	rebuild: Rebuild,
	node: Node,
	pos: number,
	out: number,
	context: { readonly row?: { readonly table: number; readonly index: number } },
	placed: Map<number, number>
): Node {
	let content = node.content

	if (node.isTextblock) {
		const inline: Node[] = []
		let changed = false

		node.forEach((child, offset) => {
			const at = pos + 1 + offset

			if (deletedAt(rebuild, at) || rebuild.markers.has(at)) {
				changed = true

				return
			}

			const ops = rebuild.marks.get(at)
			let set = child.marks

			if (ops !== undefined) {
				for (const mark of ops.remove) set = mark.removeFromSet(set)
				for (const mark of ops.add) if (node.type.allowsMarkType(mark.type)) set = mark.addToSet(set)
			}

			if (set !== child.marks) changed = true

			inline.push(set === child.marks ? child : child.mark(set))
		})

		if (changed) content = Fragment.fromArray(inline)
	} else if (!node.isLeaf && node.content.size > 0) {
		if (node.type === nodes.table) placed.set(pos, out)

		const children = rebuildChildren(
			rebuild,
			node,
			pos + 1,
			out + 1,
			node.type === nodes.table
				? { table: pos }
				: node.type === nodes.table_row && context.row !== undefined
					? { row: context.row }
					: {},
			placed
		)

		content = Fragment.fromArray(children)
	}

	const attrs = rebuild.attrs.get(pos)
	const lists = rebuild.lists.get(pos)
	const change = rebuild.changes.get(pos)

	if (attrs === undefined && lists === undefined && change === undefined) return content === node.content ? node : node.copy(content)

	return node.type.create(
		{ ...node.attrs, ...attrs, ...lists },
		content,
		change === undefined ? node.marks : withParagraphChange(node.marks, change)
	)
}

function applyBulk(state: EditorState, plan: Plan, part: StoryPart, store: PropsStore): Transaction | null {
	const tr = state.tr
	const doc = state.doc
	const removals = tableRemovals(plan)
	const rebuild = rebuildOf(state, plan, removals)
	const low = rebuild.positions[0]
	const high = rebuild.positions[rebuild.positions.length - 1]
	const placed = new Map<number, number>()

	if (low !== undefined && high !== undefined) {
		// The top-level blocks the operations touch, and those a join at the end of the last one takes in (past
		// tables that go).
		let first = -1
		let last = -1
		let from = 0
		let to = 0

		doc.forEach((child, offset, index) => {
			if (offset + child.nodeSize <= low || offset > high) return

			if (first < 0) {
				first = index
				from = offset
			}

			last = index
			to = offset + child.nodeSize
		})

		while (first >= 0 && last + 1 < doc.childCount) {
			const child = doc.child(last)
			const at = to - child.nodeSize

			if (!(child.type === nodes.paragraph && rebuild.joins.has(at)) && !(child.type === nodes.table && rebuild.gone.has(at))) break

			last++
			to += doc.child(last).nodeSize
		}

		if (first >= 0) {
			const range = doc.copy(doc.content.cut(from, to))
			const children = rebuildChildren(rebuild, range, from, from, {}, placed, from === 0 && to === doc.content.size)

			tr.replaceWith(from, to, children)
		}
	}

	// Rows and cells that go from tables that stay (by their positions now), from the end.
	const tables = [...removals].filter(([table]) => !rebuild.gone.has(table)).sort((a, b) => b[0] - a[0])

	for (const [table, removal] of tables) removeFromTable(tr, placed.get(table) ?? table, removal, part, store, plan)

	for (const { pos, name, list } of plan.lists) if (pos < 0) tr.setDocAttribute(name, list)

	return tr.docChanged ? closeHistory(markInternal(tr)) : null
}

// Joins the paragraph at `pos` with the next one, which keeps its properties (the joined paragraph ends with
// the next one's mark); markers before the first stay before the joined one. Without a paragraph after it,
// a paragraph left empty goes (its markers to the block after it); otherwise it stays: false.
function joinWithNext(tr: Transaction, pos: number, plan: Plan): boolean {
	const node = tr.doc.nodeAt(pos)

	if (node?.type !== nodes.paragraph) return false

	const end = pos + node.nodeSize
	const next = tr.doc.nodeAt(end)

	if (next?.type !== nodes.paragraph) {
		const $pos = tr.doc.resolve(pos)
		const before = node.attrs["before"] as MarkerList

		if (node.content.size > 0 || $pos.parent.childCount < 2 || next === null || !("before" in next.attrs)) return false

		tr.delete(pos, end)

		if (before.length > 0) tr.setNodeAttribute(pos, "before", [...before, ...(next.attrs["before"] as MarkerList)])

		return true
	}

	// Markers between the two (the next one's `before`) stand where they were: at the join, inline.
	const between = betweenNodes(plan, next.attrs["before"] as MarkerList)
	const joined = next.type.create(
		{ ...next.attrs, before: node.attrs["before"] },
		node.content.append(Fragment.fromArray(between)).append(next.content),
		next.marks
	)

	tr.replaceWith(pos, end + next.nodeSize, joined)

	return true
}

export type { ReviewEntry }
