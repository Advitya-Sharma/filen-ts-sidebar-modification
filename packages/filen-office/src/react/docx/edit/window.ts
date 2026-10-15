import { Fragment, type Node } from "prosemirror-model"
import { EditorState, TextSelection, type Transaction } from "prosemirror-state"
import { Mapping } from "prosemirror-transform"
import type { EditorView } from "prosemirror-view"
import { docxSchema } from "../../../core/docx/edit/index"
import { deletesBreakAtCaret, endsSection } from "../../../core/docx/stories/sections"
import type { PagedEditor } from "../../../core/docx/paged/index"

// A windowed editing view of a story: the view holds only some of the story's top-level blocks, each run of
// the others one gap paragraph of the same size (never rendered as text), so positions are the story's, a
// transaction of the view applies to the session's state step for step, and what a keystroke costs does not
// grow with the document. The input sink holds the blocks around the selection; the reflowed view also those
// on screen.

export const IDENTITY = new Mapping()

// The window is moved once the selection comes this close to its edge (so moving through the document moves
// it every few blocks, not on every block).
const WINDOW_MARGIN = 5

const GAPS = new WeakSet<Node>()
// Gaps' text only gives them their size (a slice of one string, so a gap costs no copy of it).
let gapText = "-"

function gapParagraph(size: number): Node {
	const length = Math.max(0, size - 2)

	while (gapText.length < length) gapText += gapText

	const gap = docxSchema.nodes.paragraph.create(null, length === 0 ? null : docxSchema.text(gapText.slice(0, length)))

	GAPS.add(gap)

	return gap
}

export function isGap(node: Node): boolean {
	return GAPS.has(node)
}

// The gaps' text grows to a document's size: a closed document does not keep it.
export function releaseGapText(): void {
	gapText = "-"
}

// Where the story's top-level blocks start, when the editor keeps that for the document (StoryEdits).
export interface StoryStarts {
	readonly document: Node
	readonly starts: ArrayLike<number>
	entryAt(pos: number): number
}

// The top-level index of `pos` in `doc`.
export function blockIndex(doc: Node, story: StoryStarts | undefined, pos: number): number {
	const count = doc.childCount
	const index =
		story !== undefined && story.document === doc
			? story.entryAt(pos)
			: doc.resolve(Math.max(0, Math.min(pos, doc.content.size))).index(0)

	return Math.max(0, Math.min(count - 1, index))
}

function positionOf(doc: Node, index: number): number {
	let pos = 0

	for (let k = 0; k < index; k++) pos += doc.child(k).nodeSize

	return pos
}

// The windows ([from, to) index pairs, ascending) holding `size` blocks on each side of blocks `a` through
// `b` (the selection's ends) of `count`: `previous` itself while the selection is well inside it (or it
// reaches the document's edge); a selection spanning many blocks gets the blocks around each end.
export function windowAround(previous: readonly number[], a: number, b: number, count: number, size: number): readonly number[] {
	const margin = Math.min(WINDOW_MARGIN, size >> 1)
	const [lo = 0, hi = 0] = previous
	const kept =
		previous.length === 2 &&
		hi <= count &&
		a >= lo &&
		b < hi &&
		(lo === 0 || a - lo >= margin) &&
		(hi === count || hi - 1 - b >= margin)

	if (kept) return previous

	return b - a > 2 * size
		? [Math.max(0, a - size), Math.min(count, a + size + 1), Math.max(0, b - size), Math.min(count, b + size + 1)]
		: [Math.max(0, a - size), Math.min(count, b + size + 1)]
}

// The union of window lists, ascending and merged.
export function mergeWindows(...lists: (readonly number[])[]): number[] {
	const pairs: [number, number][] = []

	for (const list of lists)
		for (let k = 0; k + 1 < list.length; k += 2) if ((list[k + 1] ?? 0) > (list[k] ?? 0)) pairs.push([list[k] ?? 0, list[k + 1] ?? 0])

	pairs.sort((x, y) => x[0] - y[0])

	const out: number[] = []

	for (const [from, to] of pairs) {
		const last = out.length - 1

		if (last > 0 && from <= (out[last] ?? 0)) out[last] = Math.max(out[last] ?? 0, to)
		else out.push(from, to)
	}

	return out
}

export function sameWindows(a: readonly number[], b: readonly number[]): boolean {
	if (a.length !== b.length) return false

	for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return false

	return true
}

export interface WindowedState {
	readonly state: EditorState
	// The gap paragraphs and gap rows in it, in order (each run of blocks or rows outside the windows is one).
	readonly gaps: readonly Node[]
	// The tables held with some of their rows only, by top-level index, as the view holds them.
	readonly tables: ReadonlyMap<number, Node>
}

// Rows of long tables: a table of more rows than this holds only those asked for (around the selection, on
// screen), each run of the others one gap row of the run's size (positions stay the table's).
export const ROW_WINDOW_MIN = 40

const GAP_ROWS = new WeakMap<Node, number>()

function gapRow(size: number, rows: number): Node {
	const nodes = docxSchema.nodes
	const paragraph = gapParagraph(Math.max(2, size - 4))
	const row = nodes.table_row.create(null, nodes.table_cell.create(null, paragraph))

	GAP_ROWS.set(row, rows)

	return row
}

// The rows a gap row stands for, or undefined for a row of the document.
export function gapRowCount(node: Node): number | undefined {
	return GAP_ROWS.get(node)
}

// The rows holding positions `positions` of `table` (at `tablePos`), `margin` rows on each side, as [from, to)
// pairs (unmerged).
export function rowsAround(table: Node, tablePos: number, positions: readonly number[], margin: number): number[] {
	const out: number[] = []
	const count = table.childCount

	for (const pos of positions) {
		if (pos <= tablePos || pos >= tablePos + table.nodeSize) continue

		let at = tablePos + 1
		let row = 0

		for (; row < count - 1; row++) {
			const end = at + table.child(row).nodeSize

			if (pos < end) break

			at = end
		}

		out.push(Math.max(0, row - margin), Math.min(count, row + margin + 1))
	}

	return out
}

// `table` holding only rows `rows` ([from, to) pairs, ascending and merged), each run of the others one gap row
// of the same size (one of `previous` when there is one of that size and count).
function windowedTable(table: Node, rows: readonly number[], previous: readonly Node[], gaps: Node[]): Node {
	const count = table.childCount
	const children: Node[] = []
	const gap = (start: number, end: number): void => {
		let size = 0

		for (let k = start; k < end; k++) size += table.child(k).nodeSize

		const found = previous.find(other => other.nodeSize === size && GAP_ROWS.get(other) === end - start && !gaps.includes(other))
		const node = found ?? gapRow(size, end - start)

		gaps.push(node)
		children.push(node)
	}
	let at = 0

	for (let w = 0; w + 1 < rows.length; w += 2) {
		// Rows asked for that the table no longer has (rows deleted since) are none.
		const start = Math.min(count, Math.max(at, rows[w] ?? 0))
		const end = Math.min(count, rows[w + 1] ?? 0)

		if (start > at) gap(at, start)

		for (let k = start; k < end; k++) children.push(table.child(k))

		at = Math.max(at, end)
	}

	if (at < count) gap(at, count)

	return table.copy(Fragment.from(children))
}

// The view's state for the session's `real`: the blocks of `windows`, each run of the others one gap paragraph
// of the run's size (one of `previous` of that size when there is one), and the selection there. A long table
// held holds the rows `rows` gives for it (all when it gives none).
export function windowedState(
	real: EditorState,
	windows: readonly number[],
	story: StoryStarts | undefined,
	previous: readonly Node[],
	rows?: (index: number, table: Node, pos: number) => readonly number[] | undefined
): WindowedState {
	const doc = real.doc
	const count = doc.childCount
	const tables = new Map<number, Node>()

	if (rows === undefined && windows.length === 2 && windows[0] === 0 && windows[1] === count)
		return { state: EditorState.create({ doc, selection: real.selection, storedMarks: real.storedMarks }), gaps: [], tables }

	const starts = story !== undefined && story.document === doc ? story.starts : undefined
	const children: Node[] = []
	const gaps: Node[] = []
	const positionOfChild = (index: number): number =>
		index >= count ? doc.content.size : starts === undefined ? positionOf(doc, index) : (starts[index] ?? 0)
	const gap = (start: number, end: number): void => {
		const size = positionOfChild(end) - positionOfChild(start)
		const node = previous.find(other => other.nodeSize === size && isGap(other) && !gaps.includes(other)) ?? gapParagraph(size)

		gaps.push(node)
		children.push(node)
	}
	let at = 0

	for (let w = 0; w + 1 < windows.length; w += 2) {
		const start = Math.max(at, windows[w] ?? 0)
		const end = Math.min(count, windows[w + 1] ?? 0)

		if (start > at) gap(at, start)

		for (let k = start; k < end; k++) {
			const child = doc.child(k)
			const held =
				rows !== undefined && child.type === docxSchema.nodes.table && child.childCount > ROW_WINDOW_MIN
					? rows(k, child, positionOfChild(k))
					: undefined

			if (held === undefined) {
				children.push(child)
			} else {
				const shown = windowedTable(child, held, previous, gaps)

				tables.set(k, shown)
				children.push(shown)
			}
		}

		at = Math.max(at, end)
	}

	if (at < count) gap(at, count)

	if (tables.size === 0 && gaps.length === 0)
		return { state: EditorState.create({ doc, selection: real.selection, storedMarks: real.storedMarks }), gaps, tables }

	const viewDoc = doc.copy(Fragment.from(children))

	return {
		state: EditorState.create({ doc: viewDoc, selection: real.selection.map(viewDoc, IDENTITY), storedMarks: real.storedMarks }),
		gaps,
		tables
	}
}

// `tr` (on a windowed view's state) as a transaction on the session's `real`: its steps, selection, stored
// marks, time and metadata; undefined when a step does not apply there.
export function sessionTransaction(tr: Transaction, viewDoc: Node, real: EditorState): Transaction | undefined {
	// The view holds the whole document: the transaction is the session's as it is.
	if (viewDoc === real.doc) return tr

	const out = real.tr

	for (const step of tr.steps) if (out.maybeStep(step).failed !== null) return undefined

	if (tr.selectionSet) out.setSelection(tr.selection.map(out.doc, IDENTITY))
	if (tr.storedMarksSet) out.setStoredMarks(tr.storedMarks)
	if (tr.scrolledIntoView) out.scrollIntoView()

	const meta: unknown = Reflect.get(tr, "meta")

	if (typeof meta === "object" && meta !== null) for (const key in meta) out.setMeta(key, Reflect.get(meta, key))

	return out.setTime(tr.time)
}

// The selection the browser holds in `view` that ProseMirror has not read yet (a caret move whose
// selectionchange is still queued), as a selection-only transaction of the view; undefined when there is none.
// Read where a key handler moved the DOM selection itself (Selection.modify), and before a state made elsewhere
// replaces the view's (marks drawn anew, a command's change): ProseMirror would otherwise put its old selection
// back in the DOM (on that update, or from its check shortly after the view takes focus) and lose the move.
export function unreadSelection(view: EditorView): Transaction | undefined {
	if (view.composing || !view.hasFocus()) return undefined

	const current = view.state.selection

	// Node and cell selections are drawn from the state, not read from the browser.
	if (!(current instanceof TextSelection)) return undefined

	const dom = view.dom.ownerDocument.getSelection()
	const anchorNode = dom?.anchorNode ?? null
	const focusNode = dom?.focusNode ?? null

	if (dom === null || anchorNode === null || focusNode === null || !view.dom.contains(anchorNode) || !view.dom.contains(focusNode))
		return undefined

	let anchor: number
	let head: number

	try {
		anchor = view.posAtDOM(anchorNode, dom.anchorOffset)
		head = view.posAtDOM(focusNode, dom.focusOffset)
	} catch {
		return undefined
	}

	if (anchor === current.anchor && head === current.head) return undefined

	const doc = view.state.doc

	if (anchor < 0 || head < 0 || anchor > doc.content.size || head > doc.content.size) return undefined

	const next = TextSelection.between(doc.resolve(anchor), doc.resolve(head))

	return next.eq(current) ? undefined : view.state.tr.setSelection(next).setMeta("addToHistory", false)
}

// An editing view of `editor` that had focus was taken down (the pages swapped for Mobile view or back): the
// view shown in its place takes focus (takeFocus), so typing goes on.
const HANDED = new WeakSet<object>()

export function handFocus(editor: object): void {
	HANDED.add(editor)
}

export function takeFocus(editor: object): boolean {
	return HANDED.delete(editor)
}

// ---- input both views handle alike --------------------------------------------------------------------------

// An Escape the editor does not take: no format painter to end, the body being edited (Escape leaves a header,
// footer, note or text box), nothing being composed. It is the host's (closing a dialog or the editor), left
// unprevented.
export function escapesToHost(editor: PagedEditor, event: KeyboardEvent): boolean {
	return (
		event.key === "Escape" &&
		!event.isComposing &&
		editor.painter === undefined &&
		editor.activeStory === editor.body.session.story.story.index
	)
}

// Backspace at the start of the paragraph after a section break: the caret goes to the end of the paragraph
// ending the section and section.deleteBreak deletes that break (the caret stays where it was otherwise).
export function deletesBreakBefore(editor: PagedEditor): boolean {
	const session = editor.session
	const { $from, empty } = session.state.selection
	const depth = $from.depth

	if (session !== editor.body.session || !empty || depth < 1 || $from.parentOffset > 0 || $from.index(depth - 1) === 0) return false

	const before = $from.node(depth - 1).child($from.index(depth - 1) - 1)

	if (before.type !== docxSchema.nodes.paragraph || !endsSection(session.story.part, before)) return false

	const selection = session.state.selection

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, $from.before() - 1)))

	if (deletesBreakAtCaret(editor)) {
		void editor.execute({ type: "section.deleteBreak" })

		return true
	}

	session.dispatch(session.state.tr.setSelection(selection.map(session.state.doc, IDENTITY)))

	return false
}

// Where the text being composed starts: the occurrence of `text` around the caret nearest to where the
// composition started (`start`, mapped through the edits; a browser may rewrite more of the paragraph than
// the composition), or `start` itself.
export function compositionFrom(doc: Node, head: number, start: number, text: string): number {
	const $head = doc.resolve(head)
	const parent = $head.parent

	if (!parent.isTextblock || text.length === 0) return Math.min(start, doc.content.size)

	const content = $head.start()
	const offset = $head.parentOffset
	const lo = Math.max(0, offset - text.length)
	const around = parent.textBetween(lo, Math.min(parent.content.size, offset + text.length), undefined, "￼")
	let best = -1

	for (let at = around.indexOf(text); at >= 0; at = around.indexOf(text, at + 1)) {
		const pos = content + lo + at

		// The caret is inside or at either end of the composed text.
		if (pos <= head && pos + text.length >= head && (best < 0 || Math.abs(pos - start) < Math.abs(best - start))) best = pos
	}

	return best >= 0 ? best : Math.min(Math.max(start, content), content + parent.content.size)
}
