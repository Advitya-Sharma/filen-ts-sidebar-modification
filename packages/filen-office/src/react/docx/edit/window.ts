import { Fragment, Slice, type Node } from "prosemirror-model"
import { EditorState, TextSelection, type Transaction } from "prosemirror-state"
import { Mapping, ReplaceStep, type Step } from "prosemirror-transform"
import type { EditorView } from "prosemirror-view"
import { docxSchema, type DocxEditorSession } from "../../../core/docx/edit/index"
import { cellStep, insertTab, type KeyLike } from "../../../core/docx/paged/keys"
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

// A story's top-level block starts kept up with its transactions, reading only the blocks they changed (as
// StoryEdits keeps its own, which the paged editor brings up to date only while the pages are laid out).
// What a transaction's steps replaced, in the document before them; undefined when they replaced nothing (mark
// steps). One pass: the range so far is carried through each later step and joined with what that step
// replaced, then mapped back once (a mapping inverted per step would be quadratic).
export function changedRange(tr: Transaction): { readonly from: number; readonly to: number } | undefined {
	let from = Infinity
	let to = -Infinity

	for (const map of tr.mapping.maps) {
		if (from <= to) {
			from = map.map(from, -1)
			to = map.map(to, 1)
		}

		map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
			from = Math.min(from, newStart)
			to = Math.max(to, newEnd)
		})
	}

	if (from > to) return undefined

	const back = tr.mapping.invert()

	return { from: back.map(from, -1), to: back.map(to, 1) }
}

export class LiveStarts implements StoryStarts {
	document: Node
	starts = new Int32Array(0)

	constructor(doc: Node) {
		this.document = doc
		this.rebuild(doc)
	}

	entryAt(pos: number): number {
		const starts = this.starts
		let low = 0
		let high = starts.length - 1

		while (low < high) {
			const middle = (low + high + 1) >> 1

			if ((starts[middle] ?? 0) <= pos) low = middle
			else high = middle - 1
		}

		return low
	}

	// The transactions a session applied (in order, each following the one before).
	applyAll(transactions: readonly Transaction[]): void {
		for (const tr of transactions) {
			if (!tr.docChanged) continue

			if (tr.before === this.document) this.change(tr)
			else this.rebuild(tr.doc)
		}
	}

	private rebuild(doc: Node): void {
		const starts = new Int32Array(doc.childCount)
		let pos = 0

		for (let k = 0; k < starts.length; k++) {
			starts[k] = pos
			pos += doc.child(k).nodeSize
		}

		this.document = doc
		this.starts = starts
	}

	private change(tr: Transaction): void {
		const before = tr.before
		const after = tr.doc
		const changed = changedRange(tr)

		// Mark and attribute steps replace nothing: every block keeps its size.
		if (changed === undefined) {
			this.document = after

			return
		}

		const { from, to } = changed
		const oldCount = before.childCount
		const newCount = after.childCount

		if (oldCount === 0 || newCount === 0) {
			this.rebuild(after)

			return
		}

		const old = this.starts
		const delta = newCount - oldCount
		let first = this.entryAt(from)
		let oldLast = Math.max(first, this.entryAt(Math.max(from, to - 1)))

		// Unchanged neighbours are the same nodes: widen until they are.
		while (first > 0 && after.child(first - 1) !== before.child(first - 1)) first--

		let newLast = oldLast + delta

		while (newLast + 1 < newCount && oldLast + 1 < oldCount && after.child(newLast + 1) !== before.child(oldLast + 1)) {
			newLast++
			oldLast++
		}

		if (newLast < first - 1 || oldLast >= oldCount || newLast >= newCount) {
			this.rebuild(after)

			return
		}

		const starts = delta === 0 ? old : new Int32Array(newCount)

		if (starts !== old) starts.set(old.subarray(0, first))

		let pos = first === 0 ? 0 : (old[first - 1] ?? 0) + before.child(first - 1).nodeSize

		for (let k = first; k <= newLast; k++) {
			starts[k] = pos
			pos += after.child(k).nodeSize
		}

		const shift = pos - (old[oldLast + 1] ?? pos)

		for (let k = newLast + 1; k < newCount; k++) starts[k] = (old[k - delta] ?? 0) + shift

		this.document = after
		this.starts = starts
	}
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
	const composing = tr.getMeta("composition") !== undefined

	// The view holds the whole document: the transaction is the session's as it is.
	if (viewDoc === real.doc && !composing) return tr

	const out = real.tr

	for (const step of tr.steps) {
		const own = composing ? typedPart(step, out.doc) : step

		if (own !== null && out.maybeStep(own).failed !== null) return undefined
	}

	if (tr.selectionSet) out.setSelection(tr.selection.map(out.doc, IDENTITY))
	if (tr.storedMarksSet) out.setStoredMarks(tr.storedMarks)
	if (tr.scrolledIntoView) out.scrollIntoView()

	const meta: unknown = Reflect.get(tr, "meta")

	if (typeof meta === "object" && meta !== null) for (const key in meta) out.setMeta(key, Reflect.get(meta, key))

	return out.setTime(tr.time)
}

function surrogate(text: string, low: boolean): boolean {
	const code = text.charCodeAt(0)

	return low ? code >= 0xdc00 && code <= 0xdfff : code >= 0xd800 && code <= 0xdbff
}

// ProseMirror reads what a composition typed as the replacement of its whole text node, the text around the
// caret too: narrowed to what differs (as its own reading of typed text is), so a tracked change records the
// typed text alone and not the rest as deleted and typed again. Null when the step changes nothing.
function typedPart(step: Step, doc: Node): Step | null {
	if (!(step instanceof ReplaceStep) || step.slice.openStart !== 0 || step.slice.openEnd !== 0) return step

	const { from, to } = step
	const $from = doc.resolve(from)
	const after = step.slice.content

	if (!$from.parent.isTextblock || to > $from.end()) return step

	for (let k = 0; k < after.childCount; k++) if (!after.child(k).isInline) return step

	const before = $from.parent.slice($from.parentOffset, $from.parentOffset + to - from).content
	let start = before.findDiffStart(after)

	if (start === null) return null

	let { a: endA, b: endB } = before.findDiffEnd(after) ?? { a: before.size, b: after.size }

	if (endA < start && before.size < after.size) {
		endB = start + (endB - endA)
		endA = start
	} else if (endB < start) {
		endA = start + (endA - endB)
		endB = start
	}

	// Never between the halves of a surrogate pair.
	if (start > 0 && surrogate(after.textBetween(start - 1, start, "", "\ufffc"), false)) start--
	if (endB < after.size && surrogate(after.textBetween(endB, endB + 1, "", "\ufffc"), true)) {
		endA++
		endB++
	}

	if (start === 0 && endA === before.size && endB === after.size) return step

	return new ReplaceStep(from + start, from + endA, new Slice(after.cut(start, endB), 0, 0))
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
export function escapesToHost(editor: PagedEditor, event: Pick<KeyLike, "key"> & { readonly isComposing: boolean }): boolean {
	return (
		event.key === "Escape" &&
		!event.isComposing &&
		editor.painter === undefined &&
		editor.activeStory === editor.body.session.story.story.index
	)
}

// Tab as Word has it (Microsoft support, "Keyboard shortcuts in Word"; "Use a screen reader to create bulleted
// or numbered lists with Word"; "Add or delete table rows and columns"): in a table, the next cell with its
// content selected (Shift+Tab the previous one), a new row below the last from the last cell (tracked while
// tracking is on); at a list paragraph's start, or over several paragraphs, one list level down (Shift+Tab
// up); otherwise a tab character. Ctrl+Tab (Option+Tab on a Mac) is a tab character anywhere, a table cell
// too. True when taken: a Tab does not move focus out of the editor (Escape does, escapesToHost) unless what it
// would do is refused (document protection). `run` applies a transaction on `session` (the story being
// edited) as the view's own keys do.
export function tabKey(
	editor: PagedEditor,
	session: DocxEditorSession,
	event: KeyLike,
	apple: boolean,
	run: (tr: Transaction | null | false) => boolean
): boolean {
	// Cmd+Tab, Alt+Tab: the system's.
	if (event.key !== "Tab" || event.metaKey || (!apple && event.altKey)) return false

	const state = session.state
	// A tab character, unless the session refuses it (its document stays as it was).
	const typeTab = (): boolean => {
		run(insertTab(state) ?? false)

		return session.state.doc !== state.doc
	}
	// Whether the story takes edits at the selection (document protection may allow none).
	const editable = (): boolean => editor.canExecute("list.indent")

	if (apple ? event.altKey || event.ctrlKey : event.ctrlKey) return !event.shiftKey && typeTab()

	const dir = event.shiftKey ? -1 : 1
	const step = cellStep(state, dir, editor.view)

	if (step !== null && step !== false && "lastRow" in step) {
		const selection = state.selection

		// The caret in the table's last row, the row command adds one below it.
		run(step.lastRow)

		if (editor.execute({ type: "table.row.insertBelow" }) !== false) return true

		if (session.state.doc === state.doc) run(session.state.tr.setSelection(selection))

		return editable()
	}

	if (step !== null) return run(step)

	const { $from, $to, empty } = state.selection
	const several = $from.depth === 0 || $to.depth === 0 || !$from.sameParent($to)
	const listStart =
		empty &&
		$from.parent.type === docxSchema.nodes.paragraph &&
		startOfText($from.parent, $from.parentOffset) &&
		session.formats.listOf($from.parent) !== undefined

	if (several || listStart) {
		if (!editable()) return false

		void editor.execute({ type: dir > 0 ? "list.indent" : "list.outdent" })

		return true
	}

	return dir > 0 ? typeTab() : editable()
}

// Inline nodes that take no room: range markers (bookmarks, comment ranges), Word's proofing and rendered page
// break marks.
const ZERO_WIDTH = new Set([docxSchema.nodes.marker, docxSchema.nodes.proof_err, docxSchema.nodes.rendered_break])

// Whether only zero-width nodes come before `offset` in `paragraph`: the caret is where its text starts.
function startOfText(paragraph: Node, offset: number): boolean {
	let at = 0

	for (let k = 0; k < paragraph.childCount && at < offset; k++) {
		const child = paragraph.child(k)

		if (!ZERO_WIDTH.has(child.type)) return false

		at += child.nodeSize
	}

	return true
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

// At compositionend: what the composition committed, which ProseMirror reads only after the event (a microtask
// later), read now. Ending the composition marks what it typed as a tracked insertion, which draws that text
// anew: before the read, the text the browser committed would be lost and the composition's last state left.
export function readComposition(view: EditorView): void {
	const observer: unknown = Reflect.get(view, "domObserver")
	const flush: unknown = typeof observer === "object" && observer !== null ? Reflect.get(observer, "flush") : undefined

	if (typeof flush === "function") flush.call(observer)
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
