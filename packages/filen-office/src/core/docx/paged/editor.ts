import type { Node } from "prosemirror-model"
import { NodeSelection, TextSelection, type EditorState, type Selection, type Transaction } from "prosemirror-state"
import { LINE_FLOATS, LINE_INTS, LineFloat, LineInt, type HitTester } from "../../text/index"
import { DocxEditorSession, type DocxEditorOptions } from "../edit/index"
import { Region, SOURCE_INTS, SourceInt, type DocumentLayout, type LayoutEvent, type PageDisplayList } from "../layout/index"
import { DocumentEdits, StoryEdits, type StoryChange } from "./edits"
import { PagedGeometry, type Affinity, type CaretBox, type HitResult, type PageRect } from "./geometry"
import { PageIndex } from "./pages"

// Paged editing of a document, headless: an editing session whose document the layout follows edit by
// edit (only the pages that change are laid out again: the ones in view within the edit's task, the rest in
// slices after it), and the caret, selection and pointer geometry over the pages as shown. A view draws
// pages, caret and selection from here and feeds input to the session; nothing here touches the DOM.

export interface PagedEditorOptions {
	readonly now?: () => number
	// Synchronous layout per edit beyond the pages in view (ms).
	readonly budgetMs?: number
}

export type PagedEditorEvent =
	// Pages laid out again (indices), the page count, and whether layout has caught up with the document.
	| { readonly type: "pages"; readonly changed: readonly number[]; readonly count: number; readonly final: boolean }
	// The selection, composition or caret geometry changed.
	| { readonly type: "selection" }
	// Editing moved to another story (the body, a header, a footer, a note): `state` is that story's.
	| { readonly type: "story"; readonly story: number }

// One story being edited: its session, the layout's view of it and its geometry over the pages.
export interface EditedStory {
	readonly session: DocxEditorSession
	readonly edits: StoryEdits
	readonly index: PageIndex
	readonly geometry: PagedGeometry
	readonly region: Region
}

const STORY_REGIONS: Readonly<Record<string, Region | undefined>> = {
	body: Region.BODY,
	header: Region.HEADER,
	footer: Region.FOOTER,
	footnote: Region.FOOTNOTE,
	endnote: Region.ENDNOTE
}

export interface PagedSelection {
	readonly from: number
	readonly to: number
	readonly anchor: number
	readonly head: number
	readonly empty: boolean
	// The selection is one atom (a locked block, an image).
	readonly node: boolean
}

const DEFAULT_BUDGET_MS = 4
const BANDS = new WeakMap<PageDisplayList, { top: number; bottom: number } | null>()

export class PagedEditor {
	readonly edits: DocumentEdits
	readonly layout: DocumentLayout
	readonly body: EditedStory
	private active: EditedStory
	private readonly stories = new Map<number, EditedStory>()
	private readonly tester: HitTester
	private readonly sessionOptions: Omit<DocxEditorOptions, "story">
	private readonly now: () => number
	private readonly budget: number
	private readonly listeners = new Set<(event: PagedEditorEvent) => void>()
	// Pages as shown: the layout's, and while it lays out again in the background, earlier ones it has not
	// reached yet.
	private readonly shown: PageDisplayList[] = []
	private readonly changed = new Set<number>()
	private flushQueued = false
	private finalSeen = false
	private disposed = false
	private readonly unsubscribe: (() => void)[] = []
	// Which side of a line wrap the caret is drawn on, and the x vertical motion keeps.
	private affinityValue: Affinity = 1
	private goal: number | undefined = undefined
	// The range of text being composed (an IME preedit), shown underlined.
	private compositionRange: { from: number; to: number } | undefined = undefined
	// The pages in view (the view tells), laid out within an edit's task.
	private visible: { first: number; last: number } = { first: 0, last: 0 }
	// Set while a transaction from geometry (a vertical move, a click) is applied: its affinity and goal stay.
	private keepMotion = false

	constructor(init: {
		readonly session: DocxEditorSession
		readonly edits: DocumentEdits
		readonly layout: DocumentLayout
		readonly tester: HitTester
		readonly options?: PagedEditorOptions
		// Options of the sessions editing other stories (headers, footers, notes), made when entered.
		readonly sessionOptions?: Omit<DocxEditorOptions, "story">
	}) {
		this.edits = init.edits
		this.layout = init.layout
		this.tester = init.tester
		this.sessionOptions = init.sessionOptions ?? {}
		this.now = init.options?.now ?? defaultNow()
		this.budget = init.options?.budgetMs ?? DEFAULT_BUDGET_MS

		for (const page of init.layout.pages) this.shown[page.index] = page

		this.body = this.edited(init.session)
		this.active = this.body
		this.finalSeen = init.layout.final
		this.unsubscribe.push(init.layout.subscribe(event => this.onLayout(event)))
	}

	private edited(session: DocxEditorSession): EditedStory {
		const edits = new StoryEdits(this.edits, session.story, session.state.doc, session.context)
		const region = STORY_REGIONS[session.story.story.kind] ?? Region.TEXTBOX
		const index = new PageIndex(edits, region)
		const story: EditedStory = {
			session,
			edits,
			index,
			geometry: new PagedGeometry(edits, index, { page: at => this.shown[at] }, this.tester, region),
			region
		}

		for (const page of this.shown) if (page !== undefined) index.index(page)

		this.stories.set(session.story.story.index, story)
		this.unsubscribe.push(session.onApplied(transactions => this.onApplied(story, transactions)))

		return story
	}

	// The story being edited.
	get session(): DocxEditorSession {
		return this.active.session
	}

	get story(): StoryEdits {
		return this.active.edits
	}

	get index(): PageIndex {
		return this.active.index
	}

	get geometry(): PagedGeometry {
		return this.active.geometry
	}

	get activeStory(): number {
		return this.active.session.story.story.index
	}

	get activeRegion(): Region {
		return this.active.region
	}

	// Every session (for saving): the body's and those of stories entered.
	get sessions(): readonly DocxEditorSession[] {
		return [...this.stories.values()].map(story => story.session)
	}

	// Edits story `story` (a header, footer or note shown on the pages) from now on, its caret drawn on page
	// `page` (the story shows on many); `page` -1 and story the body's: back to the body.
	enterStory(story: number, page: number): boolean {
		const target = story < this.edits.storyCount ? this.edits.model.stories[story] : undefined

		if (target === undefined || STORY_REGIONS[target.kind] === undefined) return false

		let edited = this.stories.get(story)

		if (edited === undefined) {
			const session = new DocxEditorSession(this.body.session.context, { ...this.sessionOptions, story: target })

			edited = this.edited(session)
		}

		edited.geometry.preferPage(page)

		if (edited === this.active) return true

		this.active = edited
		this.affinityValue = 1
		this.goal = undefined
		this.compositionRange = undefined
		this.emit({ type: "story", story })
		this.emit({ type: "selection" })

		return true
	}

	exitStory(): void {
		this.enterStory(this.body.session.story.story.index, -1)
	}

	// The band of page `page` the body's lines take (top, bottom), or undefined without body lines.
	bodyBand(page: number): { top: number; bottom: number } | undefined {
		const list = this.shown[page]

		if (list === undefined) return undefined

		const cached = BANDS.get(list)

		if (cached !== undefined) return cached === null ? undefined : cached

		const text = list.text
		let top = Infinity
		let bottom = -Infinity

		for (let line = 0; line < text.lineCount; line++) {
			const source = text.lineInts[line * LINE_INTS + LineInt.SOURCE] ?? -1

			if (list.sourceInts[source * SOURCE_INTS + SourceInt.REGION] !== Region.BODY) continue

			const lineTop = text.lineFloats[line * LINE_FLOATS + LineFloat.TOP] ?? 0

			top = Math.min(top, lineTop)
			bottom = Math.max(bottom, lineTop + (text.lineFloats[line * LINE_FLOATS + LineFloat.HEIGHT] ?? 0))
		}

		const band = top === Infinity ? null : { top, bottom }

		BANDS.set(list, band)

		return band ?? undefined
	}

	// The header or footer whose area holds page point (x, y) (above or below the body's lines), or -1.
	storyAt(page: number, _x: number, y: number): number {
		const list = this.shown[page]

		if (list === undefined) return -1

		const band = this.bodyBand(page)
		let header = -1
		let footer = -1

		for (let source = 0; source < list.sourceCount; source++) {
			const region = list.sourceInts[source * SOURCE_INTS + SourceInt.REGION]
			const story = list.sourceInts[source * SOURCE_INTS + SourceInt.STORY] ?? -1

			if (region === Region.HEADER) header = story
			else if (region === Region.FOOTER) footer = story
		}

		if (band === undefined) return -1
		if (y < band.top) return header
		if (y > band.bottom) return footer

		return -1
	}

	subscribe(listener: (event: PagedEditorEvent) => void): () => void {
		this.listeners.add(listener)

		return () => {
			this.listeners.delete(listener)
		}
	}

	dispose(): void {
		this.disposed = true

		for (const stop of this.unsubscribe) stop()

		this.unsubscribe.length = 0
		this.listeners.clear()
	}

	get state(): EditorState {
		return this.session.state
	}

	get pageCount(): number {
		return this.shown.length
	}

	// Whether layout has caught up with every edit.
	get settled(): boolean {
		return this.layout.final && this.changed.size === 0
	}

	page(index: number): PageDisplayList | undefined {
		return this.shown[index]
	}

	get affinity(): Affinity {
		return this.affinityValue
	}

	get composition(): { from: number; to: number } | undefined {
		return this.compositionRange
	}

	selection(): PagedSelection {
		const selection = this.session.state.selection

		return {
			from: selection.from,
			to: selection.to,
			anchor: selection.anchor,
			head: selection.head,
			empty: selection.empty,
			node: selection instanceof NodeSelection
		}
	}

	// ---- input --------------------------------------------------------------------------------------------

	// Applies a transaction (from the input view or a command); false when the session refused it.
	dispatch(tr: Transaction): boolean {
		return this.session.dispatch(tr)
	}

	// The view reports the pages in view; edits lay them out before returning.
	setVisible(first: number, last: number): void {
		this.visible = { first: Math.max(0, first), last: Math.max(first, last) }
		this.refreshStale(first, last)
	}

	// IME composition: the range being composed (or none), drawn underlined.
	setComposition(range: { from: number; to: number } | undefined): void {
		const current = this.compositionRange

		if (current?.from === range?.from && current?.to === range?.to) return

		this.compositionRange = range
		this.emit({ type: "selection" })
	}

	private onApplied(story: EditedStory, transactions: readonly Transaction[]): void {
		if (this.disposed) return

		let selectionSet = false

		for (const tr of transactions) if (tr.selectionSet || tr.docChanged) selectionSet = true

		if (selectionSet && !this.keepMotion && story === this.active) {
			this.affinityValue = 1
			this.goal = undefined
		}

		this.syncDefinitions()

		const change = story.edits.applyAll(transactions)

		if (change !== undefined) {
			if (story === this.body) this.relayout(change)
			else this.relayoutStory(story, change)
		}

		this.emit({ type: "selection" })
	}

	// Lists and styles commands added are laid out as the save will define them.
	private syncDefinitions(): void {
		const context = this.body.session.context
		const planned = this.edits.plan(context.newLists, context.newStyles)

		if (planned !== undefined) this.layout.redefine(planned.styles, planned.numbering)
	}

	// A header or footer changed: every page showing it gets it laid out again (those in view now, the others
	// when shown); a note changed: the body from the page it is on (its area takes room from the body).
	private relayoutStory(story: EditedStory, change: StoryChange): void {
		const layout = this.layout
		const index = story.session.story.story.index

		if (story.region === Region.HEADER || story.region === Region.FOOTER) {
			const first = layout.restory(index)

			if (first >= 0) this.refreshStale(this.visible.first, this.visible.last)
			if (!layout.final) void layout.continue()

			return
		}

		const page = story.index.restartPage([...change.replaced, ...change.around])
		const restart = layout.beginEdit(page < 0 ? 0 : page, undefined)
		const result = layout.advance(this.now() + this.budget, Math.max(restart, this.visible.last))

		if (result !== "done") void layout.continue()

		this.flush()
	}

	// ---- layout -------------------------------------------------------------------------------------------

	private relayout(change: StoryChange): void {
		const layout = this.layout
		const found = change.whole ? 0 : this.body.index.restartPage([...change.replaced, ...change.around])
		// A block the layout has not reached yet is laid out when it gets there.
		const first = found < 0 && !layout.final ? layout.pages.length : found
		const restart = layout.beginEdit(first < 0 ? 0 : first, change.whole ? undefined : change.dirty)
		// The pages in view after the restart page are laid out now, with what the budget allows beyond.
		const through = Math.max(restart, this.visible.last)
		const result = layout.advance(this.now() + this.budget, through)

		if (result === "fonts") {
			void layout.prepareFonts().then(() => {
				if (!this.disposed) void layout.continue()
			})
		} else if (result === "more") {
			void layout.continue()
		}

		this.flush()
	}

	private onLayout(event: LayoutEvent): void {
		if (this.disposed) return

		if (event.final) {
			this.finalSeen = true

			for (const index of this.layout.dirty) this.changed.add(index)
		} else {
			this.finalSeen = false

			for (let index = event.from; index < event.to; index++) this.changed.add(index)
		}

		if (!this.flushQueued) {
			this.flushQueued = true
			void Promise.resolve().then(() => this.flush())
		}
	}

	// Takes pages the layout laid out into the pages shown and tells the view.
	private flush(): void {
		this.flushQueued = false

		if (this.disposed) return

		const layout = this.layout
		const changed: number[] = []

		for (const index of this.changed) {
			const page = layout.pages[index]

			if (page === undefined || this.shown[index] === page) continue

			this.shown[index] = page
			this.indexPage(page)
			changed.push(index)
		}

		this.changed.clear()

		if (this.finalSeen && layout.final && this.shown.length > layout.pages.length) {
			this.shown.length = layout.pages.length

			for (const story of this.stories.values()) story.index.truncate(layout.pages.length)
		}

		if (layout.final) this.refreshStale(this.visible.first, this.visible.last, changed)

		if (changed.length > 0 || layout.final) {
			changed.sort((a, b) => a - b)
			this.emit({ type: "pages", changed, count: this.shown.length, final: layout.final })
		}
	}

	// Pages in [first, last] showing outdated page numbers or counts are brought up to date.
	private refreshStale(first: number, last: number, changed?: number[]): void {
		const layout = this.layout

		if (!layout.final || layout.stale.size === 0) return

		const updated: number[] = changed ?? []

		for (let index = Math.max(0, first); index <= last; index++) {
			if (!layout.stale.has(index)) continue

			const result = layout.refresh(index)

			if (result === "done") {
				const page = layout.pages[index]

				if (page !== undefined) {
					this.shown[index] = page
					this.indexPage(page)
					updated.push(index)
				}
			} else {
				// The body shows page numbers too: laid out again from there; or a header or footer changed height:
				// to the end (what its height does to the body is no state layout compares).
				layout.beginEdit(index, result === "numbers" ? [] : undefined)
				void layout.continue()

				return
			}
		}

		if (changed === undefined && updated.length > 0)
			this.emit({ type: "pages", changed: updated, count: this.shown.length, final: true })
	}

	private indexPage(page: PageDisplayList): void {
		for (const story of this.stories.values()) story.index.index(page)
	}

	private emit(event: PagedEditorEvent): void {
		for (const listener of this.listeners) listener(event)
	}

	// ---- geometry -----------------------------------------------------------------------------------------

	// Where the caret (the selection's head) is drawn.
	caret(): CaretBox | undefined {
		return this.geometry.locate(this.session.state.selection.head, this.affinityValue)
	}

	// The selection's boxes on page `page` (nothing for an empty selection).
	selectionRects(page: number): PageRect[] {
		const selection = this.session.state.selection

		if (selection.empty) return []

		return this.geometry.rects(selection.from, selection.to, page)
	}

	// Boxes of any range on a page (composition underline, find results, spelling marks).
	rangeRects(from: number, to: number, page: number): PageRect[] {
		return this.geometry.rects(from, to, page)
	}

	// The selection's two ends as drawn (for touch selection handles and a menu next to the selection).
	ends(): { readonly start: CaretBox | undefined; readonly end: CaretBox | undefined } {
		const selection = this.session.state.selection

		return {
			start: this.geometry.locate(selection.from, 1),
			end: this.geometry.locate(selection.to, selection.empty ? this.affinityValue : -1)
		}
	}

	hitTest(page: number, x: number, y: number): HitResult | undefined {
		return this.geometry.hitTest(page, x, y)
	}

	// ---- selection commands -------------------------------------------------------------------------------

	// Puts the caret at a hit (extending the selection from its anchor with `extend`).
	select(hit: HitResult, extend = false): boolean {
		const state = this.session.state
		const doc = state.doc

		if (hit.atom && !extend) {
			const node = doc.nodeAt(hit.pos)

			if (node !== null) return this.apply(state.tr.setSelection(NodeSelection.create(doc, hit.pos)), hit.affinity, undefined)
		}

		const anchor = extend ? state.selection.anchor : hit.pos

		return this.apply(state.tr.setSelection(textSelection(doc, anchor, hit.pos)), hit.affinity, undefined)
	}

	// Selects from `anchor` to the hit (a drag).
	selectRange(anchor: number, hit: HitResult): boolean {
		const state = this.session.state

		return this.apply(state.tr.setSelection(textSelection(state.doc, anchor, hit.pos)), hit.affinity, undefined)
	}

	selectWord(pos: number): boolean {
		const range = this.geometry.wordRange(pos)
		const state = this.session.state

		return this.apply(state.tr.setSelection(textSelection(state.doc, range.from, range.to)), -1, undefined)
	}

	selectParagraph(pos: number): boolean {
		const range = this.geometry.paragraphRange(pos)
		const state = this.session.state

		return this.apply(state.tr.setSelection(textSelection(state.doc, range.from, range.to)), -1, undefined)
	}

	// The caret one line down or up, keeping the x it started from; false at the document's ends.
	moveVertical(dir: 1 | -1, extend: boolean): boolean {
		const state = this.session.state
		const head = state.selection.head
		const caret = this.geometry.locate(head, this.affinityValue)
		const goal = this.goal ?? caret?.x ?? 0
		const hit = this.geometry.verticalMove(head, this.affinityValue, goal, dir)

		if (hit === undefined) {
			// Past the first or last line: the document's start or end.
			const pos = dir > 0 ? endOf(state.doc) : startOf(state.doc)
			const anchor = extend ? state.selection.anchor : pos

			return this.apply(state.tr.setSelection(textSelection(state.doc, anchor, pos)), 1, undefined)
		}

		const anchor = extend ? state.selection.anchor : hit.pos

		return this.apply(state.tr.setSelection(textSelection(state.doc, anchor, hit.pos)), hit.affinity, goal)
	}

	// The caret to the start or end of its line.
	moveLineBoundary(dir: 1 | -1, extend: boolean): boolean {
		const state = this.session.state
		const hit = this.geometry.lineBoundary(state.selection.head, this.affinityValue, dir)

		if (hit === undefined) return false

		const anchor = extend ? state.selection.anchor : hit.pos

		return this.apply(state.tr.setSelection(textSelection(state.doc, anchor, hit.pos)), hit.affinity, undefined)
	}

	// The caret to the document's start or end.
	moveDocumentBoundary(dir: 1 | -1, extend: boolean): boolean {
		const state = this.session.state
		const pos = dir > 0 ? endOf(state.doc) : startOf(state.doc)
		const anchor = extend ? state.selection.anchor : pos

		return this.apply(state.tr.setSelection(textSelection(state.doc, anchor, pos)), 1, undefined)
	}

	// The caret moved by a hit the view found (a page up or down: the point a view height away).
	moveTo(hit: HitResult, extend: boolean, keepGoal: boolean): boolean {
		const state = this.session.state
		const anchor = extend ? state.selection.anchor : hit.pos

		return this.apply(state.tr.setSelection(textSelection(state.doc, anchor, hit.pos)), hit.affinity, keepGoal ? this.goal : undefined)
	}

	// The x vertical motion keeps (the caret's own when none was kept).
	goalX(): number | undefined {
		return this.goal ?? this.caret()?.x
	}

	private apply(tr: Transaction, affinity: Affinity, goal: number | undefined): boolean {
		this.keepMotion = true

		try {
			const applied = this.session.dispatch(tr.scrollIntoView())

			if (applied) {
				this.affinityValue = affinity
				this.goal = goal
				this.emit({ type: "selection" })
			}

			return applied
		} finally {
			this.keepMotion = false
		}
	}
}

function defaultNow(): () => number {
	const performance: unknown = Reflect.get(globalThis, "performance")
	const now: unknown = typeof performance === "object" && performance !== null ? Reflect.get(performance, "now") : undefined

	return typeof now === "function" ? () => (now as () => number).call(performance) : Date.now
}

function textSelection(doc: Node, anchor: number, head: number): Selection {
	const size = doc.content.size
	const $anchor = doc.resolve(Math.max(0, Math.min(anchor, size)))
	const $head = doc.resolve(Math.max(0, Math.min(head, size)))

	return TextSelection.between($anchor, $head)
}

function startOf(doc: Node): number {
	return TextSelection.atStart(doc).from
}

function endOf(doc: Node): number {
	return TextSelection.atEnd(doc).to
}
