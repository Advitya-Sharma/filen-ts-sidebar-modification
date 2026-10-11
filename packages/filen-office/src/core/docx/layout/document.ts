import type { LineSpacingRule } from "../../fonts/index"
import { Exclusions, LINE_FLOATS, LINE_INTS, LineInt, MeasureCache, StyleTable, Wrap, type TextMeasurer } from "../../text/index"
import type { DocxModel, Section } from "../model/document"
import { Item, RevisionView } from "../model/kinds"
import type { StoryPart } from "../model/part"
import { NumberingCounters, type CounterSnapshot, type ListMarker } from "../numbering/counter"
import { PropsStore } from "../props/store"
import type { ResolvedSectionProps } from "../props/types"
import type { Numbering } from "../numbering/numbering"
import type { DocxStyles } from "../styles/styles"
import { resolveSection, type PropertyResolver } from "../styles/resolve"
import type { PageValues, StoryContext } from "./content"
import { REF_SPAN, type LayoutEdits } from "./edits"
import { numberText } from "./fields"
import {
	Area,
	AreaStart,
	BlockFlow,
	DEFAULT_TUNING,
	Engine,
	FlowStatus,
	nextToken,
	type FlowSnapshot,
	type LayoutTuning,
	type PageFrame
} from "./flow"
import { layoutCompat, pageGeometry, type LayoutCompat, type PageGeometry } from "./geometry"
import { FootnoteArea, NoteNumbering, type FootnoteState, type NoteNumberingSnapshot, type PlacedNote } from "./notes"
import { ObjectTable, type ObjectSizer } from "./objects"
import { PageGraphic, Layer, OutputMark, PageDisplayList, PageOutput, Region, SOURCE_INTS, SourceFlag, SourceInt } from "./output"
import { DecorationTable, RunStyles, borderWidth, type DecorationValue, type FontSource, type RunStyleInfo } from "./styles"
import { EPSILON, twips } from "./units"

export interface LayoutOptions {
	// Which tracked changes show (default Final, as Word's Simple Markup lays out).
	readonly view?: RevisionView
	// Lay out hidden text (w:vanish) as Word does with "Show hidden text" on.
	readonly showHidden?: boolean
	// The store property XML is parsed into (a new one by default).
	readonly store?: PropsStore
	readonly sizer?: ObjectSizer
	readonly lineSpacingRule?: LineSpacingRule
	readonly tuning?: Partial<LayoutTuning>
	// Time slice in ms between yields (default 8; Infinity runs to the end without yielding).
	readonly sliceMs?: number
	// Gives the host a turn between slices (web: scheduler.yield or a MessageChannel; default a microtask).
	readonly yield?: () => Promise<void>
	readonly now?: () => number
	readonly signal?: { readonly aborted: boolean }
	// Stops after this many pages (default 100000).
	readonly maxPages?: number
	// NUMPAGES before the layout knows it (default: Word's recorded page breaks + 1).
	readonly estimatedPages?: number
	// Record where each item's content lands in its paragraph's text (PageDisplayList.sourceMapPool).
	readonly recordMaps?: boolean
	// Unsaved edits to lay out in place of their source (read whenever layout restarts).
	readonly edits?: LayoutEdits
}

// A top-level body block an edit changed, as laid out now: relaying out cannot stop at a page that matches
// the previous layout before it has laid each of them out.
export interface DirtyBlock {
	readonly section: number
	readonly ref: number
}

export type AdvanceResult = "done" | "more" | "fonts"

export type RefreshResult = "done" | "numbers" | "moved"

// Entries [from, to) of a table added since the previous event.
export interface TableRange {
	readonly from: number
	readonly to: number
}

export interface LayoutEvent {
	readonly type: "pages"
	readonly from: number
	readonly to: number
	readonly final: boolean
	// Style (styles, styleInfo), decoration and object entries added since the previous event: everything
	// pages [from, to) use that earlier events did not deliver, so a worker can send just these.
	readonly styles: TableRange
	readonly decorations: TableRange
	readonly objects: TableRange
}

// Where a page's flow starts: for editing, which relays out from the first page that changed.
export interface PageAnchor {
	readonly section: number
	// The story the page's flow continues (the body; an endnote's story when endnotes run on).
	readonly story: number
	// The top-level block (paragraph or table item of the story) the page starts in, -1 after the last;
	// `part` is the part it is laid out from and `index` its place among the section's blocks.
	readonly block: number
	readonly part: number
	readonly index: number
	// Inside a table: the row the page starts at; -1 otherwise.
	readonly row: number
	// Inside a paragraph: the offset into its laid-out text where the page's first line starts; 0 otherwise.
	readonly offset: number
}

// A Word page break (w:lastRenderedPageBreak) located in this layout: the page its line landed on and the
// line's index among the body lines of the whole document (table cells in their layout order).
export interface RenderedBreak {
	readonly page: number
	readonly line: number
	// The paragraph holding it, and which of the paragraph's breaks it is.
	readonly item: number
	readonly ordinal: number
}

// The incremental result of layoutDocument: pages arrive as they are laid out.
export class DocumentLayout {
	readonly model: DocxModel
	readonly pages: PageDisplayList[] = []
	readonly anchors: PageAnchor[] = []
	readonly styles: StyleTable
	readonly objects: ObjectTable
	readonly renderedBreaks: RenderedBreak[] = []
	// Index of each page's first body line (as RenderedBreak.line counts).
	readonly pageFirstLines: number[] = []
	final = false
	cancelled = false
	error: unknown = undefined
	// Pages laid out again after the layout first delivered them (page count fields).
	readonly dirty = new Set<number>()
	// Pages kept from an earlier layout whose page number or count fields show old values (see refresh).
	readonly stale = new Set<number>()
	done: Promise<void> = Promise.resolve()
	private readonly runStyles: RunStyles
	private readonly listeners = new Set<(event: LayoutEvent) => void>()

	constructor(model: DocxModel, styles: StyleTable, runStyles: RunStyles, objects: ObjectTable) {
		this.model = model
		this.styles = styles
		this.runStyles = runStyles
		this.objects = objects
	}

	get styleInfo(): readonly RunStyleInfo[] {
		return this.runStyles.info
	}

	get decorations(): readonly DecorationValue[] {
		return this.runStyles.decorations.values
	}

	pageCount(): { count: number; final: boolean } {
		return { count: this.pages.length, final: this.final }
	}

	subscribe(listener: (event: LayoutEvent) => void): () => void {
		this.listeners.add(listener)

		return () => {
			this.listeners.delete(listener)
		}
	}

	cancel(): void {
		this.cancelled = true
	}

	private relayout: ((page: number) => Promise<number>) | undefined = undefined
	private running: () => boolean = () => false
	private driver: LayoutDriver | undefined = undefined

	attach(relayout: (page: number) => Promise<number>, running?: () => boolean, driver?: LayoutDriver): void {
		this.relayout = relayout

		if (running !== undefined) this.running = running
		if (driver !== undefined) this.driver = driver
	}

	// Editing: starts laying the document out again for edits that change what pages from `page` on show
	// (the first page showing a changed block), keeping earlier pages. `dirty` lists the changed top-level
	// blocks: once each is laid out, layout stops at the first page that starts like a page of the previous
	// layout and keeps the rest (renumbered when the page count changed). Call advance() until "done".
	// Returns the first page laid out again.
	// Without `dirty` layout runs to the end. A run under way (the first layout, or the background part of an
	// earlier edit's) is taken over: layout carries on from the earlier restart page.
	beginEdit(page: number, dirty: readonly DirtyBlock[] | undefined): number {
		if (this.driver === undefined) throw new Error("No layout to edit")

		this.cancelled = false

		return this.driver.begin(page, dirty)
	}

	// Carries on laying out in slices (as the first layout does) until done; resolves then.
	continue(): Promise<void> {
		return this.driver?.continue() ?? Promise.resolve()
	}

	// Lays out synchronously until `deadline` (the clock passed in LayoutOptions.now) or until pages up to
	// `through` are laid out: "done" when the layout is complete, "fonts" when fonts must load first
	// (prepareFonts()).
	advance(deadline: number, through = Infinity): AdvanceResult {
		return this.driver?.advance(deadline, through) ?? "done"
	}

	prepareFonts(): Promise<void> {
		return this.driver?.prepareFonts() ?? Promise.resolve()
	}

	// Brings a stale page up to date (its header and footer laid out again): "done", or what keeps it from it:
	// "numbers" when its body shows page values that changed (lay out again from it, beginEdit(page, [])),
	// "moved" when the header or footer changed height (beginEdit(page, undefined): to the end).
	refresh(page: number): RefreshResult {
		return this.driver?.refresh(page) ?? "numbers"
	}

	// Styles and list definitions an editor added (what a save will add): later layout resolves with them.
	redefine(styles: DocxStyles, numbering: Numbering | undefined): void {
		this.driver?.redefine(styles, numbering)
	}

	// A header, footer or other page story's text changed: the pages showing it become stale (refresh() lays
	// them out again). Returns the first such page, or -1.
	restory(story: number): number {
		return this.driver?.restory(story) ?? -1
	}

	get editing(): boolean {
		return this.driver !== undefined
	}

	// Where the last relayout for edits met the previous layout again and kept its pages (-1: ran to the end).
	get resyncedAt(): number {
		return this.driver?.resyncedAt ?? -1
	}

	// Lays the document out again from page `page` on: pages before the nearest page layout can restart at
	// (at or before `page`) are kept, the rest arrive again as events. Resolves to the first page laid out
	// again once the layout is final. The model is the one laid out (an edit will come with a new model or a
	// patch to it); call it once `done` has settled, never while a layout runs (that rejects).
	async relayoutFrom(page: number): Promise<number> {
		if (this.relayout === undefined) return 0

		if (this.running()) throw new Error("relayoutFrom while the layout runs")

		this.cancelled = false

		const run = this.relayout(Math.max(0, page))

		this.done = run.then(() => undefined)

		return run
	}

	// Bytes of all pages' display lists.
	get byteLength(): number {
		return this.pages.reduce((sum, page) => sum + page.byteLength, 0)
	}

	private sentStyles = 0
	private sentDecorations = 0
	private sentObjects = 0

	// Tells subscribers about pages [from, to), with the table entries added since the last event.
	notify(from: number, to: number, final: boolean): void {
		const styles = this.runStyles.info.length
		const decorations = this.runStyles.decorations.values.length
		const objects = this.objects.count
		const event: LayoutEvent = {
			type: "pages",
			from,
			to,
			final,
			styles: { from: this.sentStyles, to: styles },
			decorations: { from: this.sentDecorations, to: decorations },
			objects: { from: this.sentObjects, to: objects }
		}

		this.sentStyles = styles
		this.sentDecorations = decorations
		this.sentObjects = objects

		for (const listener of this.listeners) listener(event)
	}
}

// The paginator's state where a page opens: layout can restart there.
interface PageSnapshot {
	readonly sectionIndex: number
	readonly section: ResolvedSectionProps
	readonly flow: FlowSnapshot
	readonly pageNumber: number
	readonly restart: number | undefined
	readonly sectionFirstPage: boolean
	readonly nextStart: AreaStart
	readonly bodyLines: number
	readonly lineSection: number
	readonly endnotesDone: number
	readonly globalLine: number
	readonly sectionPageCounts: readonly number[]
	readonly counters: CounterSnapshot | undefined
	readonly notes: NoteNumberingSnapshot
	readonly carried: readonly PlacedNote[]
	// Endnotes referenced and not placed yet.
	readonly endnotes: readonly number[]
	readonly renderedBreaks: number
	// List markers counted before the page opened for paragraphs laid out from it on (built ahead).
	readonly markers: readonly (readonly [number, ListMarker | null])[]
	// Note numbers assigned the same way.
	readonly assigned: readonly (readonly [number, string])[]
}

const Step = { PROGRESS: 0, PAGE: 1, SUSPEND: 2, FINISHED: 3 } as const

type Step = (typeof Step)[keyof typeof Step]

// What to lay out again when a page's header or footer changes (their text, or the page numbers or counts
// they show), the body kept.
interface PageRecord {
	readonly index: number
	// Its header or footer shows the page count (NUMPAGES, SECTIONPAGES), or the page number.
	readonly totals: boolean
	readonly numbers: boolean
	// Where the header ends and the footer begins (they set the body's area).
	readonly headerBottom: number
	readonly footerTop: number
	// The page's content before its header and footer were appended.
	readonly body: OutputMark
	readonly header: number
	readonly footer: number
	readonly geometry: PageGeometry
	readonly values: PageValues
	readonly frame: PageFrame
	readonly info: ConstructorParameters<typeof PageDisplayList>[1]
}

function countRenderedBreaks(part: StoryPart, start: number, end: number): number {
	const kinds = part.kinds
	let count = 0

	for (let item = start; item < end; item++) if (kinds[item] === Item.LastRenderedPageBreak) count++

	return count
}

class Paginator {
	private readonly engine: Engine
	private readonly layout: DocumentLayout
	private readonly model: DocxModel
	private readonly body: StoryPart | undefined
	private bodyStory: StoryContext | undefined
	private readonly compat: LayoutCompat
	private readonly maxPages: number
	private readonly page = new PageOutput()
	private readonly header = new PageOutput()
	private readonly footer = new PageOutput()
	private readonly exclusions = new Exclusions()
	private readonly footnotes: FootnoteArea
	private readonly area = new Area()
	private readonly storyCounters = new Map<number, NumberingCounters | undefined>()
	// By page index.
	private records: (PageRecord | undefined)[] = []
	private readonly sectionPageCounts: number[] = []
	private readonly sectionEstimates: number[] = []
	// Per page, the values page count fields in its body, notes or text boxes showed (undefined: none).
	private readonly bodyTotals: (PageValues | undefined)[] = []
	// Per page, whether its body, notes or text boxes show the page number.
	private readonly bodyNumbers: boolean[] = []
	private estimate: number
	private readonly initialNotes: NoteNumberingSnapshot
	private sectionIndex = -1
	private section: ResolvedSectionProps = resolveSection({})
	private flows: BlockFlow[] = []
	private flowIndex = 0
	private pageOpen = false
	private pageNumber = 0
	private restart: number | undefined = undefined
	private sectionFirstPage = true
	// The open page is its section's first.
	private firstOfSection = true
	private geometry: PageGeometry | undefined = undefined
	private column = 0
	// Where the section's part of the open page began (undefined until its first step there; null when it
	// cannot be laid out again: it began inside a table): to balance its columns.
	private balanceStart: BalanceStart | null | undefined = undefined
	// Each column of the page has notes of its own (the page began in a section with columns); the lowest
	// content of the column being filled.
	private notesPerColumn = false
	private columnBottom = 0
	private sectionTop = 0
	private bodyTop = 0
	private bodyBottom = 0
	private contentBottom = 0
	private values: PageValues = { page: 1, pageFormat: undefined, numPages: 1, sectionPages: 1 }
	private frame: PageFrame = { width: 816, height: 1056, marginLeft: 96, marginRight: 96, marginTop: 96, marginBottom: 96, odd: true }
	private headerStory = -1
	private footerStory = -1
	private headerBottom = 0
	private footerTop = 0
	private pageSection = 0
	// The properties of the section the open page belongs to.
	private pageProps: ResolvedSectionProps = resolveSection({})
	private anchor: PageAnchor = { section: 0, story: -1, block: -1, part: -1, index: -1, row: -1, offset: 0 }
	// The section line numbers were last counted in.
	private lineSection = -1
	// The section whose endnotes were placed last.
	private endnotesDone = -1
	private bodyLines = 0
	// Why the next page opens.
	private nextStart: AreaStart = AreaStart.DOCUMENT
	finished = false

	constructor(engine: Engine, layout: DocumentLayout, estimate: number | undefined, maxPages: number) {
		this.engine = engine
		this.layout = layout
		this.model = engine.model
		this.compat = engine.compat
		this.maxPages = maxPages
		this.footnotes = new FootnoteArea(engine)
		this.initialNotes = engine.notes.snapshot()
		engine.notes.pageNotes = () => this.footnotes.count

		const story = this.model.body
		const body = story === undefined ? undefined : this.model.parts[story.part]

		this.body = body
		this.bodyStory =
			body === undefined || story === undefined
				? undefined
				: { part: body, story: story.index, counters: this.counters(), noteMark: "", breaksAsLines: false }

		let total = 0

		for (const section of this.model.sections) {
			const count = body === undefined ? 0 : countRenderedBreaks(body, section.start, section.end)

			this.sectionEstimates.push(count + 1)
			total += count
		}

		this.estimate = estimate ?? Math.max(1, total + 1)
	}

	// The resolver's definitions changed (styles and lists an editor added): counting goes on with them.
	redefine(styles: DocxStyles, numbering: Numbering | undefined): void {
		this.engine.resolver.redefine(styles, numbering)

		if (numbering === undefined) return

		const body = this.bodyStory

		if (body?.counters !== undefined) body.counters.use(numbering)
		else if (body !== undefined) this.bodyStory = { ...body, counters: new NumberingCounters(numbering) }

		for (const [story, counters] of this.storyCounters) {
			if (counters !== undefined) counters.use(numbering)
			else this.storyCounters.delete(story)
		}
	}

	private counters(): NumberingCounters | undefined {
		const numbering = this.engine.resolver.numbering

		return numbering === undefined ? undefined : new NumberingCounters(numbering)
	}

	private storyContext(story: number): StoryContext | undefined {
		const root = this.model.stories[story]
		const part = root === undefined ? undefined : this.model.parts[root.part]

		if (root === undefined || part === undefined) return undefined

		let counters = this.storyCounters.get(story)

		if (!this.storyCounters.has(story)) {
			counters = this.counters()
			this.storyCounters.set(story, counters)
		}

		return { part, story, counters, noteMark: "", breaksAsLines: true }
	}

	// Lays out every header, footer and note once off-page, so the fonts they need (tables and text boxes
	// included) are loaded before layout starts: they are laid out where layout cannot wait for fonts.
	scanStories(): void {
		const scratch = new PageOutput()

		// By index: an editing layout's model drops stories of overlay parts it no longer uses.
		for (let index = 0; index < this.model.stories.length; index++) {
			const story = this.model.stories[index]

			if (story === undefined) continue
			if (story.kind !== "header" && story.kind !== "footer" && story.kind !== "footnote" && story.kind !== "endnote") continue

			const context = this.storyContext(story.index)

			if (context === undefined) continue

			const region = story.kind === "header" ? Region.HEADER : story.kind === "footer" ? Region.FOOTER : Region.FOOTNOTE

			this.engine.layoutBox({ ...context, counters: undefined }, region, scratch, this.values, this.frame, 0, 0, 600)
			scratch.reset()
		}
	}

	private resolved(index: number): ResolvedSectionProps {
		const section = this.model.sections[index]
		const body = this.body

		if (section === undefined || body === undefined || section.sectPr < 0) return resolveSection({})

		return this.engine.resolver.section(this.engine.store.section(body.document, section.sectPr))
	}

	step(): Step {
		if (this.finished) return Step.FINISHED

		if (this.sectionIndex < 0) {
			if (!this.startSection(0)) {
				this.finished = true

				return Step.FINISHED
			}
		}

		if (!this.pageOpen) {
			this.takeSnapshot()

			if (this.resyncs !== undefined && this.resynchronise(this.resyncs)) return Step.FINISHED

			this.openPage()
		}

		const flow = this.flows[this.flowIndex]

		if (flow === undefined) return this.nextSection()

		if (
			this.balanceStart === undefined &&
			this.column === 0 &&
			this.flowIndex === 0 &&
			flow.region === Region.BODY &&
			(this.geometry?.columnX.length ?? 1) > 1
		) {
			const state = flow.snapshot()

			this.balanceStart =
				state === undefined
					? null
					: {
							body: flow,
							flow: state,
							mark: this.page.mark(),
							notes: this.footnotes.save(),
							exclusions: this.exclusions.count,
							endnotes: this.engine.endnotes.length,
							empty: this.area.empty,
							start: this.area.start
						}
		}

		const status = flow.run(this.area)

		this.contentBottom = Math.max(this.contentBottom, flow.y)
		this.columnBottom = Math.max(this.columnBottom, flow.y)

		switch (status) {
			case FlowStatus.DONE:
				this.flowIndex++

				return this.flowIndex >= this.flows.length ? this.nextSection() : Step.PROGRESS
			case FlowStatus.SUSPEND:
				return Step.SUSPEND
			case FlowStatus.PAGE_BREAK:
				this.closePage()
				this.nextStart = AreaStart.BREAK

				return Step.PAGE
			case FlowStatus.COLUMN_BREAK:
				return this.nextColumn(AreaStart.BREAK)
			default:
				return this.nextColumn(AreaStart.OVERFLOW)
		}
	}

	// Snapshots by page index; a page opened inside a table or among endnotes has none.
	private readonly snapshots: (PageSnapshot | undefined)[] = []

	private takeSnapshot(): void {
		const index = this.layout.pages.length
		const flow = this.flows[this.flowIndex]
		const state = this.flowIndex === 0 && flow?.region === Region.BODY ? flow.snapshot() : undefined

		this.snapshots[index] =
			state === undefined
				? undefined
				: {
						sectionIndex: this.sectionIndex,
						section: this.section,
						flow: state,
						pageNumber: this.pageNumber,
						restart: this.restart,
						sectionFirstPage: this.sectionFirstPage,
						nextStart: this.nextStart,
						bodyLines: this.bodyLines,
						lineSection: this.lineSection,
						endnotesDone: this.endnotesDone,
						globalLine: this.globalLine,
						sectionPageCounts: this.sectionPageCounts.slice(),
						counters: this.bodyStory?.counters?.snapshot(),
						notes: this.engine.notes.snapshot(),
						carried: this.footnotes.carriedNotes,
						endnotes: this.engine.endnotes.slice(),
						renderedBreaks: this.layout.renderedBreaks.length,
						markers: flow === undefined ? [] : this.markersAhead(flow),
						assigned: flow === undefined ? [] : this.engine.notes.assignedWhere((part, item) => this.ahead(flow, part, item))
					}
	}

	// The nearest page at or before `page` layout can restart at.
	restartPage(page: number): number {
		for (let index = Math.min(page, this.snapshots.length - 1); index > 0; index--)
			if (this.snapshots[index] !== undefined) return index

		return 0
	}

	// Drops pages from `page` (one with a snapshot, or 0) and continues from there. With `dirty`, the dropped
	// pages are kept aside: layout stops where it meets them again once every dirty block is laid out.
	// `own`: the pages this run laid out from `page` on may be met again too (not when they are laid out again
	// for being wrong).
	restartAt(page: number, dirty?: readonly DirtyBlock[], own = true): void {
		const layout = this.layout
		const snapshot = page > 0 ? this.snapshots[page] : undefined

		this.resyncs = dirty === undefined ? undefined : this.previousLayout(own ? page : layout.pages.length, dirty)
		this.missedFonts = this.missedFonts >= page ? -1 : this.missedFonts
		this.resyncedAt = -1
		this.pageOpen = false
		this.finished = false
		this.records.length = Math.min(this.records.length, page)
		this.bodyNumbers.length = Math.min(this.bodyNumbers.length, page)
		this.snapshots.length = page
		this.bodyTotals.length = Math.min(this.bodyTotals.length, page)
		layout.pages.length = page
		layout.anchors.length = page
		layout.pageFirstLines.length = page
		layout.final = false

		for (const index of [...layout.dirty]) if (index >= page) layout.dirty.delete(index)
		for (const index of [...layout.stale]) if (index >= page) layout.stale.delete(index)

		if (snapshot === undefined) {
			this.restartFromStart()

			return
		}

		layout.renderedBreaks.length = snapshot.renderedBreaks
		this.sectionIndex = snapshot.sectionIndex
		this.section = snapshot.section

		const section = this.model.sections[snapshot.sectionIndex]
		const body = this.body
		const story = this.bodyStory

		const flow = this.bodyFlow(snapshot.sectionIndex)

		if (section === undefined || body === undefined || story === undefined || flow === undefined) return

		this.flows = [flow]
		this.flowIndex = 0
		this.pageNumber = snapshot.pageNumber
		this.restart = snapshot.restart
		this.sectionFirstPage = snapshot.sectionFirstPage
		this.nextStart = snapshot.nextStart
		this.bodyLines = snapshot.bodyLines
		this.lineSection = snapshot.lineSection
		this.endnotesDone = snapshot.endnotesDone
		this.globalLine = snapshot.globalLine
		this.sectionPageCounts.splice(0, this.sectionPageCounts.length, ...snapshot.sectionPageCounts)

		// Taken before the document had lists (planned ones came since): nothing counted yet.
		if (snapshot.counters !== undefined) story.counters?.restore(snapshot.counters)
		else story.counters?.reset()

		this.engine.notes.restore(snapshot.notes)
		this.engine.notes.forgetAssigned(this.isBodyPart, snapshot.assigned)
		this.footnotes.carriedNotes = snapshot.carried
		this.engine.setEndnotes(snapshot.endnotes)
		this.forgetMarkers()

		for (const [key, marker] of snapshot.markers) this.engine.markers.set(key, marker)

		// The page opens again with the restored state; the flow resumes in its first area.
		this.snapshots[page] = snapshot
		this.openPage()
		flow.restore(snapshot.flow, this.area)
		this.anchor = this.anchorHere()
	}

	private restartFromStart(): void {
		this.layout.renderedBreaks.length = 0
		this.sectionIndex = -1
		this.section = resolveSection({})
		this.flows = []
		this.flowIndex = 0
		this.pageNumber = 0
		this.restart = undefined
		this.sectionFirstPage = true
		this.nextStart = AreaStart.DOCUMENT
		this.bodyLines = 0
		this.lineSection = -1
		this.endnotesDone = -1
		this.globalLine = 0
		this.sectionPageCounts.length = 0
		this.bodyStory?.counters?.reset()
		this.engine.notes.restore(this.initialNotes)
		this.engine.notes.forgetAssigned(this.isBodyPart, [])
		this.footnotes.carriedNotes = []
		this.engine.truncateEndnotes(0)
		this.forgetMarkers()
	}

	// The parts whose paragraphs the body flow lays out.
	private readonly isBodyPart = (part: number): boolean => {
		const edits = this.engine.edits

		return edits === undefined ? part === (this.body?.index ?? -1) : edits.bodyPart(part)
	}

	// Whether item `item` of `part` lies in a block of `flow` from the one being laid out to the last one built
	// ahead of it (keep chains, frames, joined paragraphs).
	private ahead(flow: BlockFlow, part: number, item: number): boolean {
		const last = Math.max(flow.index, flow.furthest)

		for (let k = flow.index; k <= last && k < flow.count; k++) {
			const block = flow.blocks[k] ?? -1

			if (flow.parts[k] === part && item >= block && item < (flow.partAt(k)?.end(block) ?? block + 1)) return true

			// An edited paragraph counts under its source's key.
			const key = this.engine.memoKey(flow.parts[k] ?? -1, block)
			const keyPart = Math.floor(key / REF_SPAN)

			if (keyPart === part && key - keyPart * REF_SPAN === item) return true
		}

		return false
	}

	// List markers counted for blocks the page that opens lays out again (built ahead before it opened).
	private markersAhead(flow: BlockFlow): [number, ListMarker | null][] {
		const out: [number, ListMarker | null][] = []

		for (const [key, marker] of this.engine.markers) {
			const part = Math.floor(key / REF_SPAN)

			if (this.isBodyPart(part) && this.ahead(flow, part, key - part * REF_SPAN)) out.push([key, marker])
		}

		return out
	}

	// List markers the body flow counted are counted again (layout restarts at a page).
	private forgetMarkers(): void {
		const markers = this.engine.markers

		for (const key of markers.keys()) if (this.isBodyPart(Math.floor(key / REF_SPAN))) markers.delete(key)
	}

	// The body flow of section `index`: the model's blocks, or the edited body's.
	private bodyFlow(index: number): BlockFlow | undefined {
		const section = this.model.sections[index]
		const body = this.body
		const story = this.bodyStory

		if (section === undefined || body === undefined || story === undefined) return undefined

		const flow = new BlockFlow(this.engine, story, Region.BODY, true)
		const blocks = this.engine.edits?.body

		if (blocks === undefined) return flow.setRange(body, section.start, section.end)

		return flow.setRefs(body, blocks.refs, index === 0 ? 0 : (blocks.sectionEnds[index - 1] ?? 0), blocks.sectionEnds[index] ?? 0)
	}

	private startSection(index: number): boolean {
		const flow = this.bodyFlow(index)

		if (flow === undefined) return false

		this.sectionIndex = index
		this.section = this.resolved(index)
		this.flows = [flow]
		this.flowIndex = 0
		this.engine.notes.startSection(this.section.footnotes, this.section.endnotes, index === 0)

		const start = this.section.pageNumbering.start

		if (start !== undefined) this.restart = start

		return true
	}

	private nextSection(): Step {
		const index = this.sectionIndex + 1
		const section = this.model.sections[index]

		if (this.endnotesDone !== this.sectionIndex) {
			this.endnotesDone = this.sectionIndex

			if (this.queueEndnotes(section === undefined)) return Step.PROGRESS
		}

		if (section === undefined) {
			this.closePage()
			this.finished = true

			return Step.FINISHED
		}

		const previous = this.section
		const next = this.resolved(index)
		const sameSize = Math.abs(next.pageWidth - previous.pageWidth) < 1 && Math.abs(next.pageHeight - previous.pageHeight) < 1

		if (this.pageOpen && next.type === "continuous" && sameSize) {
			// The section's columns end level before the next section starts below them.
			this.balanceColumns()

			if (this.geometry !== undefined) this.columnSeparators(this.geometry)

			if (this.notesPerColumn) this.closeColumnNotes()

			this.startSection(index)
			this.balanceStart = undefined
			this.geometry = pageGeometry(next, this.compat, this.frame.odd)
			this.engine.gridPitch = this.geometry.gridPitch
			this.sectionTop = this.contentBottom
			this.column = 0
			this.setArea(0, this.sectionTop, this.sectionTop <= this.bodyTop + EPSILON)
			this.area.start = AreaStart.BREAK

			if (this.notesPerColumn) this.footnotes.startColumn(this.area.left, this.area.width)

			return Step.PROGRESS
		}

		if (
			this.pageOpen &&
			next.type === "nextColumn" &&
			sameSize &&
			this.geometry !== undefined &&
			this.column + 1 < this.geometry.columnX.length
		) {
			this.startSection(index)

			return this.nextColumn(AreaStart.BREAK)
		}

		// The page closes with the section it belongs to before the next one starts.
		if (this.pageOpen) this.closePage()

		this.startSection(index)

		this.sectionFirstPage = true
		this.nextStart = AreaStart.BREAK

		const upcoming = this.restart ?? this.pageNumber + 1
		const wrongParity = (next.type === "evenPage" && upcoming % 2 === 1) || (next.type === "oddPage" && upcoming % 2 === 0)

		if (wrongParity && this.layout.pages.length > 0) {
			// Word leaves a blank page to start the section on the right side.
			this.openPage()
			this.closePage()
			this.sectionFirstPage = true
		}

		return Step.PAGE
	}

	// Endnotes referenced so far, after the section (w:pos sectEnd) or the document's last one.
	private queueEndnotes(atEnd: boolean): boolean {
		const engine = this.engine

		if (engine.endnotes.length === 0 || (engine.notes.endnotes.position !== "sectEnd" && !atEnd)) return false

		const stories = engine.takeEndnotes()

		const flows: BlockFlow[] = []
		const model = this.model
		let separator = -1

		for (const id of model.settings.endnoteSeparators) {
			const story = model.endnotes.get(id)

			if (story?.noteType === "separator") separator = story.index
		}

		for (const index of separator >= 0 ? [separator, ...stories] : stories) {
			const story = model.stories[index]
			const part = story === undefined ? undefined : model.parts[story.part]

			if (story === undefined || part === undefined) continue

			const context: StoryContext = {
				part,
				story: index,
				counters: undefined,
				noteMark: engine.notes.marks.get(index) ?? "",
				breaksAsLines: true
			}

			flows.push(new BlockFlow(engine, context, Region.ENDNOTE, true).setContainer(part, story.item))
		}

		if (flows.length === 0) return false

		this.flows = flows
		this.flowIndex = 0

		return true
	}

	// Lays the section's part of the page out again with the lowest column bottom that still holds it, so the
	// columns end level (the first ones longer when they cannot be equal).
	private balanceColumns(): void {
		const geometry = this.geometry
		const start = this.balanceStart
		const flow = this.flows[0]

		// Not when the section's endnotes ran on after its text (they are no part to lay out again).
		if (geometry === undefined || start === undefined || start === null || flow !== start.body || geometry.columnX.length < 2) return

		let low = this.sectionTop
		let high = this.bodyBottom

		for (let k = 0; k < BALANCE_STEPS && high - low > 0.5; k++) {
			const middle = (low + high) / 2

			if (this.fillColumns(start, flow, middle, geometry.columnX.length)) high = middle
			else low = middle
		}

		if (!this.fillColumns(start, flow, high, geometry.columnX.length))
			this.fillColumns(start, flow, this.bodyBottom, geometry.columnX.length)
	}

	// Lays the section's part of the page out from `start` into its columns ending at `bottom`; whether it
	// all fitted.
	private fillColumns(start: BalanceStart, flow: BlockFlow, bottom: number, columns: number): boolean {
		this.page.rollback(start.mark)
		this.footnotes.load(start.notes)
		this.exclusions.count = start.exclusions
		this.engine.truncateEndnotes(start.endnotes)
		this.contentBottom = this.sectionTop
		this.column = 0
		this.setArea(0, this.sectionTop, start.empty)
		this.area.start = start.start
		this.area.bottom = bottom
		flow.restore(start.flow, this.area)

		for (;;) {
			const status = flow.run(this.area)

			this.contentBottom = Math.max(this.contentBottom, flow.y)

			if (status === FlowStatus.DONE) return true

			if ((status !== FlowStatus.OVERFLOW && status !== FlowStatus.COLUMN_BREAK) || this.column + 1 >= columns) return false

			if (this.notesPerColumn) this.closeColumnNotes()

			this.column++
			this.setArea(this.column, this.sectionTop, this.sectionTop <= this.bodyTop + EPSILON)
			this.area.start = status === FlowStatus.COLUMN_BREAK ? AreaStart.BREAK : AreaStart.OVERFLOW
			this.area.bottom = bottom

			if (this.notesPerColumn) this.footnotes.startColumn(this.area.left, this.area.width)
		}
	}

	// Draws the notes of the column being left at its bottom; text that comes later on the page keeps clear
	// of them.
	private closeColumnNotes(): void {
		const top = this.footnotes.close(this.page, this.bodyBottom, this.columnBottom)

		if (top < this.bodyBottom - EPSILON)
			this.exclusions.add(this.area.left, top, this.area.left + this.area.width, this.bodyBottom, Wrap.BOTH)
	}

	private columnSeparators(geometry: PageGeometry): void {
		if (!geometry.separator) return

		for (let c = 1; c < geometry.columnX.length; c++) {
			const previousEnd = (geometry.columnX[c - 1] ?? 0) + (geometry.columnWidths[c - 1] ?? 0)
			const x = geometry.left + (previousEnd + (geometry.columnX[c] ?? 0)) / 2

			this.page.addGraphic(
				PageGraphic.COLUMN_SEPARATOR,
				0,
				Layer.TEXT,
				x - 0.5,
				this.sectionTop,
				1,
				Math.max(1, this.contentBottom - this.sectionTop)
			)
		}
	}

	private nextColumn(start: AreaStart): Step {
		const geometry = this.geometry

		if (geometry !== undefined && this.column + 1 < geometry.columnX.length) {
			if (this.notesPerColumn) this.closeColumnNotes()

			this.column++
			this.setArea(this.column, this.sectionTop, this.sectionTop <= this.bodyTop + EPSILON)
			this.area.start = start

			if (this.notesPerColumn) this.footnotes.startColumn(this.area.left, this.area.width)

			return Step.PROGRESS
		}

		this.closePage()
		this.nextStart = start

		return Step.PAGE
	}

	private setArea(column: number, top: number, empty: boolean): void {
		const geometry = this.geometry
		const area = this.area

		if (geometry === undefined) return

		this.columnBottom = top
		area.page = this.page
		area.region = Region.BODY
		area.left = geometry.left + (geometry.columnX[column] ?? 0)
		area.width = geometry.columnWidths[column] ?? geometry.width - geometry.left - geometry.right
		area.top = top
		area.bottom = this.bodyBottom
		area.empty = empty
		area.breaks = true
		area.inTable = false
		area.exclusions = this.exclusions
		area.footnotes = this.footnotes
		area.values = this.values
		area.frame = this.frame
		area.token = nextToken()
	}

	// Lays out a header or footer story at `top`; its height.
	private layoutStory(story: number, output: PageOutput, region: Region, top: number): number {
		const context = this.storyContext(story)
		const geometry = this.geometry

		if (context === undefined || geometry === undefined) return 0

		const root = this.model.stories[story]

		if (root === undefined) return 0

		return this.engine.layoutBox(
			context,
			region,
			output,
			this.values,
			this.frame,
			geometry.left,
			top,
			geometry.width - geometry.left - geometry.right
		)
	}

	private anchorHere(): PageAnchor {
		const flow = this.flows[this.flowIndex]
		const position = flow?.position()

		return {
			section: this.sectionIndex,
			story: flow?.story.story ?? -1,
			block: position?.block ?? -1,
			part: position?.part ?? -1,
			index: position?.index ?? -1,
			row: position?.row ?? -1,
			offset: position?.offset ?? 0
		}
	}

	private headerFooter(set: Section["effectiveHeaders"], odd: boolean): number {
		if (this.sectionFirstPage && this.section.titlePage) return set.first
		if (this.compat.evenAndOddHeaders && !odd) return set.even

		return set.default
	}

	private openPage(): void {
		const section = this.section
		const sectionInfo = this.model.sections[this.sectionIndex]
		const number = this.restart ?? this.pageNumber + 1

		this.restart = undefined
		this.pageNumber = number

		const odd = number % 2 !== 0
		const geometry = pageGeometry(section, this.compat, odd)

		this.geometry = geometry
		this.engine.gridPitch = geometry.gridPitch
		this.page.reset()
		this.header.reset()
		this.footer.reset()
		this.exclusions.clear()
		this.engine.notes.startPage()
		this.values = {
			page: number,
			pageFormat: section.pageNumbering.format,
			numPages: this.estimate,
			sectionPages: this.sectionEstimates[this.sectionIndex] ?? 1
		}
		this.frame = {
			width: geometry.width,
			height: geometry.height,
			marginLeft: geometry.left,
			marginRight: geometry.right,
			marginTop: geometry.top,
			marginBottom: geometry.height - geometry.bottom,
			odd
		}
		this.pageSection = this.sectionIndex
		this.pageProps = section

		// Their font is asked for before the page's text: it is loaded before the page is done.
		if (section.lineNumbering !== undefined) this.lineNumberStyle()

		this.anchor = this.anchorHere()

		const textWidth = geometry.width - geometry.left - geometry.right

		this.headerStory = sectionInfo === undefined ? -1 : this.headerFooter(sectionInfo.effectiveHeaders, odd)
		this.footerStory = sectionInfo === undefined ? -1 : this.headerFooter(sectionInfo.effectiveFooters, odd)

		let headerBottom = 0
		let footerTop = geometry.height

		if (this.headerStory >= 0)
			headerBottom = geometry.header + this.layoutStory(this.headerStory, this.header, Region.HEADER, geometry.header)

		if (this.footerStory >= 0) {
			const height = this.layoutStory(this.footerStory, this.footer, Region.FOOTER, 0)

			footerTop = geometry.height - geometry.footer - height
			this.footer.move(new OutputMark(), null, 0, footerTop)
		}

		this.headerBottom = headerBottom
		this.footerTop = footerTop

		this.bodyTop = geometry.topFixed ? geometry.top : Math.max(geometry.top, headerBottom)
		this.bodyBottom = geometry.bottomFixed ? geometry.bottom : Math.min(geometry.bottom, footerTop)

		if (this.bodyBottom - this.bodyTop < 24) this.bodyBottom = this.bodyTop + 24

		this.notesPerColumn = geometry.columnX.length > 1
		this.footnotes.startPage(
			geometry.left + (this.notesPerColumn ? (geometry.columnX[0] ?? 0) : 0),
			this.notesPerColumn ? (geometry.columnWidths[0] ?? textWidth) : textWidth,
			this.values,
			this.frame
		)
		this.column = 0
		this.balanceStart = undefined
		this.sectionTop = this.bodyTop
		this.contentBottom = this.bodyTop
		this.setArea(0, this.bodyTop, true)
		this.area.start = this.nextStart
		this.nextStart = AreaStart.OVERFLOW
		this.pageOpen = true
		this.firstOfSection = this.sectionFirstPage
		this.sectionFirstPage = false
	}

	private closePage(): void {
		if (!this.pageOpen) return

		const page = this.page
		const geometry = this.geometry
		const layout = this.layout

		this.pageOpen = false

		if (geometry === undefined) return

		this.verticalAlign(page)
		this.footnotes.close(page, this.bodyBottom, this.notesPerColumn ? this.columnBottom : this.contentBottom)

		this.columnSeparators(geometry)

		this.lineNumbers(page)
		this.pageBorders(page, geometry)

		const index = layout.pages.length
		const bodyFirst = this.firstBodyItem(page)
		const info = {
			index,
			number: this.pageNumber,
			numberText: numberText(this.pageNumber, this.values.pageFormat),
			section: this.pageSection,
			width: geometry.width,
			height: geometry.height,
			bodyFirst,
			bodyLast: this.lastBodyItem(page)
		}

		this.renderedBreaks(page, index)

		// Drawn with stand-in metrics where a header, footer, note or number needed a font still loading.
		if (this.missedFonts < 0 && this.engine.styles.hasMissing) this.missedFonts = index

		this.bodyTotals[index] = this.usesTotals(page) ? this.values : undefined
		this.bodyNumbers[index] = usesFlag(page, SourceFlag.PAGE_FIELDS)

		const totals = this.usesTotals(this.header) || this.usesTotals(this.footer)
		const numbers = usesFlag(this.header, SourceFlag.PAGE_FIELDS) || usesFlag(this.footer, SourceFlag.PAGE_FIELDS)

		this.records[index] = {
			index,
			totals,
			numbers,
			headerBottom: this.headerBottom,
			footerTop: this.footerTop,
			body: page.mark(),
			header: this.headerStory,
			footer: this.footerStory,
			geometry,
			values: this.values,
			frame: this.frame,
			info
		}

		page.append(this.header, 0, 0)
		page.append(this.footer, 0, 0)
		layout.pages.push(new PageDisplayList(page, info))
		layout.anchors.push(this.anchor)
		this.sectionPageCounts[this.pageSection] = (this.sectionPageCounts[this.pageSection] ?? 0) + 1
		layout.notify(index, index + 1, false)
		this.trimMemos()

		if (layout.pages.length >= this.maxPages) this.finished = true
	}

	// w:pgBorders: lines at their space from the page edge or around the text area.
	private pageBorders(page: PageOutput, geometry: PageGeometry): void {
		const borders = this.pageProps.pageBorders

		if (borders === undefined) return
		if (borders.display === "firstPage" && !this.firstOfSection) return
		if (borders.display === "notFirstPage" && this.firstOfSection) return

		const resolver = this.engine.resolver
		const decorations = this.engine.decorations
		const fromPage = borders.offsetFrom === "page"
		const layer = borders.zOrder === "back" ? Layer.BEHIND : Layer.TEXT
		const space = (border: { readonly space?: number } | undefined): number => ((border?.space ?? 0) * 4) / 3
		const top = resolver.border(borders.top)
		const bottom = resolver.border(borders.bottom)
		const left = resolver.border(borders.left)
		const right = resolver.border(borders.right)
		const x0 = fromPage ? space(borders.left) : geometry.left - space(borders.left)
		const x1 = fromPage ? geometry.width - space(borders.right) : geometry.width - geometry.right + space(borders.right)
		const y0 = fromPage ? space(borders.top) : this.bodyTop - space(borders.top)
		const y1 = fromPage ? geometry.height - space(borders.bottom) : this.bodyBottom + space(borders.bottom)
		const line = (border: typeof top, x: number, y: number, width: number, height: number): void => {
			if (border !== undefined) page.addGraphic(PageGraphic.BORDER, decorations.border(border), layer, x, y, width, height)
		}

		line(top, x0, y0, x1 - x0, top === undefined ? 0 : borderWidth(top))
		line(bottom, x0, y1 - (bottom === undefined ? 0 : borderWidth(bottom)), x1 - x0, bottom === undefined ? 0 : borderWidth(bottom))
		line(left, x0, y0, left === undefined ? 0 : borderWidth(left), y1 - y0)
		line(right, x1 - (right === undefined ? 0 : borderWidth(right)), y0, right === undefined ? 0 : borderWidth(right), y1 - y0)
	}

	private usesTotals(output: PageOutput): boolean {
		return usesFlag(output, SourceFlag.TOTAL_FIELDS)
	}

	// The body flow's list markers of blocks already behind it are no longer needed.
	private trimMemos(): void {
		const markers = this.engine.markers
		const flow = this.flows[this.flowIndex]

		if (markers.size < 8192 || flow === undefined || flow.region !== Region.BODY) return

		for (const key of markers.keys()) {
			const part = Math.floor(key / REF_SPAN)
			const item = key - part * REF_SPAN

			if (!this.isBodyPart(part) || this.ahead(flow, part, item)) continue

			// The block before the one being laid out stays: a keep chain may still roll back into it.
			const previous = flow.index - 1
			const block = flow.blocks[previous] ?? -1

			if (previous >= 0 && flow.parts[previous] === part && item >= block && item < (flow.partAt(previous)?.end(block) ?? block))
				continue

			markers.delete(key)
		}
	}

	private firstBodyItem(page: PageOutput): number {
		for (let s = 0; s < page.sourceCount; s++)
			if (page.sourceInts[s * SOURCE_INTS] === Region.BODY) return page.sourceInts[s * SOURCE_INTS + SourceInt.ITEM] ?? -1

		return -1
	}

	private lastBodyItem(page: PageOutput): number {
		for (let s = page.sourceCount - 1; s >= 0; s--)
			if (page.sourceInts[s * SOURCE_INTS] === Region.BODY) return page.sourceInts[s * SOURCE_INTS + SourceInt.ITEM] ?? -1

		return -1
	}

	// Section vertical alignment (w:vAlign) on pages of one section.
	private verticalAlign(page: PageOutput): void {
		const align = this.pageProps.vAlign

		if ((align !== "center" && align !== "bottom") || this.pageSection !== this.sectionIndex) return

		const free = this.bodyBottom - this.footnotes.reserved - this.contentBottom

		if (free <= EPSILON) return

		page.move(new OutputMark(), null, 0, align === "center" ? free / 2 : free)
	}

	// Line numbers in the margin (w:lnNumType), counting body lines outside tables.
	// Line numbers' style (w:lnNumType draws them in the LineNumber character style).
	private lineNumberStyle(): number {
		const engine = this.engine

		return engine.styles.style(engine.resolver.run(engine.resolver.paragraph({}), LINE_NUMBER_RUN), "ascii")
	}

	private lineNumbers(page: PageOutput): void {
		const numbering = this.pageProps.lineNumbering
		const geometry = this.geometry

		if (numbering === undefined || geometry === undefined || this.body === undefined) return

		const countBy = Math.max(1, numbering.countBy ?? 1)
		const distance = numbering.distance === undefined ? 18 : twips(numbering.distance)
		const text = page.text
		const engine = this.engine
		const style = this.lineNumberStyle()
		const start = numbering.start ?? 1
		let n =
			numbering.restart === "continuous" || (numbering.restart === "newSection" && this.lineSection === this.pageSection)
				? this.bodyLines
				: 0

		this.lineSection = this.pageSection
		const lines = text.lineCount

		for (let line = 0; line < lines; line++) {
			const source = text.lineInts[line * LINE_INTS + LineInt.SOURCE] ?? 0

			if (page.sourceInts[source * SOURCE_INTS] !== Region.BODY || (page.sourceFlags(source) & SourceFlag.IN_TABLE) !== 0) continue

			n++

			if ((start - 1 + n) % countBy !== 0) continue

			const label = String(start - 1 + n)
			const width = engine.styles.table.cache.width(engine.styles.table.fontKey[style] ?? 0, label, 0, label.length)
			const top = text.lineTop(line)
			const height = text.lineHeight(line)
			const baseline = text.lineBaseline(line)
			const x = (text.lineFloats[line * LINE_FLOATS + 3] ?? geometry.left) - distance - width
			const labelSource = page.addSource(Region.LINE_NUMBER, -1, -1, 0, label)

			text.addLine(top, height, baseline, x, width, 0, label.length, 0, labelSource)
			text.addFragment(x, width, 0, 0, 0, label.length, style, 0, 0)
		}

		this.bodyLines = n
	}

	private renderedBreaks(page: PageOutput, index: number): void {
		const layout = this.layout
		const text = page.text
		let first = -1

		for (let line = 0; line < text.lineCount; line++) {
			const source = text.lineInts[line * LINE_INTS + LineInt.SOURCE] ?? 0

			if (page.sourceInts[source * SOURCE_INTS] !== Region.BODY) continue

			const global = this.globalLine++

			if (first < 0) first = global

			const breaks = page.sourceBreaks[source]

			if (breaks === undefined) continue

			const start = text.lineInts[line * LINE_INTS + LineInt.TEXT_START] ?? 0
			const end = text.lineInts[line * LINE_INTS + LineInt.TEXT_END] ?? 0

			breaks.forEach((offset, ordinal) => {
				if (offset >= start && (offset < end || (offset === end && (text.lineFlags(line) & 1) !== 0))) {
					layout.renderedBreaks.push({
						page: index,
						line: global,
						item: page.sourceInts[source * SOURCE_INTS + SourceInt.ITEM] ?? -1,
						ordinal
					})
				}
			})
		}

		layout.pageFirstLines.push(first < 0 ? this.globalLine : first)
	}

	private globalLine = 0

	// The first page whose body showed page counts other than the layout's, or -1.
	totalsMismatch(): number {
		const total = this.layout.pages.length

		for (let index = 0; index < this.bodyTotals.length; index++) {
			const values = this.bodyTotals[index]

			if (values === undefined) continue

			const section = this.layout.pages[index]?.section ?? 0

			if (values.numPages !== total || values.sectionPages !== (this.sectionPageCounts[section] ?? 1)) return index
		}

		return -1
	}

	// Takes the page counts of the layout so far as the estimates fields show from now on.
	adoptTotals(): void {
		this.estimate = Math.max(1, this.layout.pages.length)

		for (let section = 0; section < this.sectionEstimates.length; section++) {
			const count = this.sectionPageCounts[section]

			if (count !== undefined) this.sectionEstimates[section] = count
		}
	}

	// Page count fields: pages whose headers or footers used the estimate are laid out again, or with `lazy`
	// marked stale for refresh().
	finish(lazy = false): void {
		const layout = this.layout
		const total = layout.pages.length

		this.resyncs = undefined

		for (const entry of this.records) {
			if (entry === undefined || !entry.totals) continue

			const sectionPages = this.sectionPageCounts[entry.info.section] ?? 1

			if (entry.values.numPages === total && entry.values.sectionPages === sectionPages) continue

			if (lazy) layout.stale.add(entry.index)
			else void this.redoHeaderFooter(entry)
		}
	}

	// Lays page `entry.index` out again with its current page values (number and counts): its header and
	// footer, the body kept. "numbers" when the body shows page values that changed too, "moved" when the
	// header or footer changed height (the body's area moved): the page needs layout.
	redoHeaderFooter(entry: PageRecord): RefreshResult {
		const layout = this.layout
		const list = layout.pages[entry.index]

		if (list === undefined) return "numbers"

		const values: PageValues = {
			...entry.values,
			page: list.number,
			numPages: layout.pages.length,
			sectionPages: this.sectionPageCounts[entry.info.section] ?? 1
		}
		const body = this.bodyTotals[entry.index]

		if (
			(body !== undefined && (body.numPages !== values.numPages || body.sectionPages !== values.sectionPages)) ||
			(this.bodyNumbers[entry.index] === true && entry.values.page !== values.page)
		)
			return "numbers"

		const page = this.page
		const geometry = entry.geometry
		const info = { ...entry.info, index: entry.index, number: list.number, numberText: list.numberText }

		this.values = values
		this.frame = entry.frame
		this.geometry = geometry
		page.reset()
		this.header.reset()
		this.footer.reset()

		const headerBottom =
			entry.header >= 0 ? geometry.header + this.layoutStory(entry.header, this.header, Region.HEADER, geometry.header) : 0
		let footerTop = geometry.height

		if (entry.footer >= 0) {
			const height = this.layoutStory(entry.footer, this.footer, Region.FOOTER, 0)

			footerTop = geometry.height - geometry.footer - height
			this.footer.move(new OutputMark(), null, 0, footerTop)
		}

		// A header or footer of another height moves the body: it is laid out again instead.
		if (
			(!geometry.topFixed && Math.abs(Math.max(geometry.top, headerBottom) - Math.max(geometry.top, entry.headerBottom)) > EPSILON) ||
			(!geometry.bottomFixed && Math.abs(Math.min(geometry.bottom, footerTop) - Math.min(geometry.bottom, entry.footerTop)) > EPSILON)
		)
			return "moved"

		page.append(list, 0, 0, -Infinity, Infinity, null, entry.body)
		page.append(this.header, 0, 0)
		page.append(this.footer, 0, 0)
		layout.pages[entry.index] = new PageDisplayList(page, info)
		layout.dirty.add(entry.index)
		this.records[entry.index] = { ...entry, values, info, headerBottom, footerTop }

		return "done"
	}

	// A stale page brought up to date (see DocumentLayout.refresh).
	refresh(index: number): RefreshResult {
		const layout = this.layout

		if (!layout.stale.has(index)) return "done"

		const entry = this.records[index]
		const result = entry === undefined ? "numbers" : this.redoHeaderFooter(entry)

		if (result === "done") layout.stale.delete(index)

		return result
	}

	// Pages showing story `story` as their header or footer are marked stale (its text changed); the first, or
	// -1.
	restory(story: number): number {
		let first = -1

		for (const entry of this.records) {
			if (entry === undefined || (entry.header !== story && entry.footer !== story)) continue

			this.layout.stale.add(entry.index)

			if (first < 0) first = entry.index
		}

		// Pages of the earlier layout a relayout under way may keep: kept stale.
		const resync = this.resyncs

		resync?.records.forEach((entry, k) => {
			if (entry !== undefined && (entry.header === story || entry.footer === story)) resync.stale.add(resync.from + k)
		})

		return first
	}

	// Whether the page being laid out shows story `story` as its header or footer.
	shows(story: number): boolean {
		return this.pageOpen && (this.headerStory === story || this.footerStory === story)
	}

	private resyncs: Resync | undefined = undefined
	// The first page closed while fonts it may have used were loading (-1: none).
	missedFonts = -1

	// The layout as it stands, from `page` on, for a relayout to meet again: the pages laid out so far and,
	// while an earlier relayout is still under way, the earlier layout's pages from the one starting where
	// that run stands (none to meet again when no such page is found: the run then lays out to the end).
	private previousLayout(page: number, dirty: readonly DirtyBlock[]): Resync | undefined {
		const layout = this.layout
		const progress = layout.pages.length
		const previous = this.resyncs

		if (previous === undefined) {
			if (page >= progress) return undefined

			return {
				from: page,
				pages: layout.pages.slice(page),
				anchors: layout.anchors.slice(page),
				firstLines: layout.pageFirstLines.slice(page),
				snapshots: this.snapshots.slice(page, progress),
				bodyTotals: this.bodyTotals.slice(page, progress),
				records: this.records.slice(page, progress),
				bodyNumbers: this.bodyNumbers.slice(page, progress),
				renderedBreaks: layout.renderedBreaks.slice(),
				stale: within(layout.stale, page, progress),
				redone: within(layout.dirty, page, progress),
				sectionPageCounts: this.sectionPageCounts.slice(),
				globalLine: this.globalLine,
				dirty,
				dirtyIndices: new Map(),
				byAnchor: undefined
			}
		}

		// The earlier layout's pages continue the run's pages only from the page beginning where the run stands.
		const anchor = this.pageOpen ? this.anchor : this.anchorHere()
		const cut = anchor.block < 0 ? -1 : previous.anchors.findIndex(old => sameAnchor(old, anchor))

		if (cut < 0) return undefined

		const first = previous.from + cut
		const rest = <T>(own: readonly T[], earlier: readonly T[]): T[] => own.slice(page, progress).concat(earlier.slice(cut))
		const breaks = layout.renderedBreaks.slice()

		for (const entry of previous.renderedBreaks) if (entry.page >= first) breaks.push({ ...entry, page: entry.page - first + progress })

		const marks = (own: ReadonlySet<number>, earlier: ReadonlySet<number>): Set<number> => {
			const set = within(own, page, progress)

			for (const index of earlier) if (index >= first) set.add(index - first + progress)

			return set
		}

		return {
			from: page,
			pages: rest(layout.pages, previous.pages),
			anchors: rest(layout.anchors, previous.anchors),
			firstLines: rest(layout.pageFirstLines, previous.firstLines),
			snapshots: rest(this.snapshots, previous.snapshots),
			bodyTotals: rest(this.bodyTotals, previous.bodyTotals),
			records: rest(this.records, previous.records),
			bodyNumbers: rest(this.bodyNumbers, previous.bodyNumbers),
			renderedBreaks: breaks,
			stale: marks(layout.stale, previous.stale),
			redone: marks(layout.dirty, previous.redone),
			sectionPageCounts: previous.sectionPageCounts.slice(),
			globalLine: previous.globalLine,
			dirty: previous.dirty.concat(dirty),
			dirtyIndices: new Map(),
			byAnchor: undefined
		}
	}

	// Pages laid out so far in the current run (a relayout under way reaches no further).
	get progress(): number {
		return this.layout.pages.length
	}

	// Whether a relayout for edits is still under way and may meet the previous layout again.
	get resyncing(): boolean {
		return this.resyncs !== undefined
	}

	// At a page about to open while laying out again for edits: when the flow has laid out every dirty block
	// and stands where a page of the previous layout started, in the same state, the rest of the previous
	// layout is kept from there (moved and renumbered when pages were added or removed before it).
	private resynchronise(resync: Resync): boolean {
		const layout = this.layout
		const index = layout.pages.length
		const snapshot = this.snapshots[index]
		const flow = this.flows[this.flowIndex]

		if (index <= resync.from || snapshot === undefined || flow === undefined || !this.dirtyDone(resync, flow)) return false

		const anchor = this.anchorHere()
		const old = this.matchingPage(resync, anchor, index)

		if (old < 0) return false

		const oldSnapshot = resync.snapshots[old - resync.from]

		if (oldSnapshot === undefined || !sameState(snapshot, oldSnapshot)) return false

		this.keepTail(resync, old, index, snapshot, oldSnapshot, anchor)

		return true
	}

	// Whether every dirty block lies behind the body flow.
	private dirtyDone(resync: Resync, flow: BlockFlow): boolean {
		const section = this.sectionIndex

		for (const block of resync.dirty) {
			if (block.section > section) return false
			if (block.section < section) continue

			let at = resync.dirtyIndices.get(block.ref)

			if (at === undefined) {
				at = -1

				for (let k = 0; k < flow.count; k++) {
					if (flow.refAt(k) === block.ref) {
						at = k

						break
					}
				}

				resync.dirtyIndices.set(block.ref, at)
			}

			if (at >= flow.index) return false
		}

		return true
	}

	// The page of the previous layout that starts where `anchor` stands, or -1.
	private matchingPage(resync: Resync, anchor: PageAnchor, index: number): number {
		// Most edits leave the page count as it was: the same page index first.
		const candidate = resync.anchors[index - resync.from]

		if (candidate !== undefined && sameAnchor(candidate, anchor)) return index

		if (resync.byAnchor === undefined) {
			resync.byAnchor = new Map()

			resync.anchors.forEach((old, k) => {
				const key = anchorKey(old)

				if (!resync.byAnchor?.has(key)) resync.byAnchor?.set(key, resync.from + k)
			})
		}

		return resync.byAnchor.get(anchorKey(anchor)) ?? -1
	}

	// Keeps the previous layout's pages from `old` on as pages from `index` on.
	private keepTail(
		resync: Resync,
		old: number,
		index: number,
		snapshot: PageSnapshot,
		oldSnapshot: PageSnapshot,
		anchor: PageAnchor
	): void {
		const layout = this.layout
		const shift = index - old
		// A section numbering its pages anew from this page on: its numbers are as they were.
		const numberShift = snapshot.restart === undefined ? snapshot.pageNumber - oldSnapshot.pageNumber : 0
		const lineShift = this.globalLine - oldSnapshot.globalLine
		const blockShift = (anchor.index ?? 0) - (resync.anchors[old - resync.from]?.index ?? 0)
		const breakShift = layout.renderedBreaks.length - oldSnapshot.renderedBreaks
		const section = this.sectionIndex
		const from = old - resync.from
		const stale = shift !== 0 || numberShift !== 0
		const count = resync.pages.length - from
		// Page numbers move in this section and the later ones up to the first that numbers its pages anew.
		const renumbered =
			numberShift === 0 ? section + 1 : this.numberedFrom(section, resync.pages[resync.pages.length - 1]?.section ?? section)

		for (let k = 0; k < count; k++) {
			const page = resync.pages[from + k]
			const at = index + k

			if (page === undefined) continue

			const number = page.section < renumbered ? page.number + numberShift : page.number
			const moved =
				page.index === at && number === page.number
					? page
					: page.moved(at, number, numberText(number, resync.records[from + k]?.values.pageFormat ?? this.formatOf(page.section)))
			const oldAnchor = resync.anchors[from + k]
			const record = resync.records[from + k]
			const kept = resync.snapshots[from + k]

			layout.pages.push(moved)
			layout.anchors.push(
				oldAnchor === undefined || oldAnchor.section !== section || blockShift === 0
					? (oldAnchor ?? anchor)
					: { ...oldAnchor, index: oldAnchor.index + blockShift }
			)
			layout.pageFirstLines.push((resync.firstLines[from + k] ?? 0) + lineShift)
			this.snapshots[at] =
				kept === undefined
					? undefined
					: rebased(kept, section, renumbered, numberShift, lineShift, blockShift, breakShift, snapshot, oldSnapshot)
			this.bodyTotals[at] = resync.bodyTotals[from + k]
			this.records[at] =
				record === undefined
					? undefined
					: { ...record, index: at, info: { ...record.info, index: at, number, numberText: moved.numberText } }
			this.bodyNumbers[at] = resync.bodyNumbers[from + k] ?? false

			if (resync.redone.has(old + k)) layout.dirty.add(at)

			if (
				resync.stale.has(old + k) ||
				(stale && (record?.numbers === true || record?.totals === true || this.bodyNumbers[at] === true))
			)
				layout.stale.add(at)
		}

		for (const entry of resync.renderedBreaks) {
			if (entry.page >= old) layout.renderedBreaks.push({ ...entry, page: entry.page + shift, line: entry.line + lineShift })
		}

		// Section page counts: this section's grew by the pages laid out, the later ones are as they were.
		for (let s = section; s < resync.sectionPageCounts.length; s++) {
			const before = oldSnapshot.sectionPageCounts[s] ?? 0
			const now = snapshot.sectionPageCounts[s] ?? 0

			this.sectionPageCounts[s] = (resync.sectionPageCounts[s] ?? 0) + (s === section ? now - before : 0)
		}

		this.globalLine = resync.globalLine + lineShift
		this.resyncs = undefined
		this.finished = true
		this.resyncedAt = index

		if (shift !== 0 || numberShift !== 0) layout.notify(index, layout.pages.length, false)
	}

	// The first section after `section`, up to `last`, whose page numbering starts anew (last + 1 when none).
	private numberedFrom(section: number, last: number): number {
		for (let index = section + 1; index <= last; index++) if (this.resolved(index).pageNumbering.start !== undefined) return index

		return last + 1
	}

	// Where the last relayout met the previous layout again (-1 when it ran to the end).
	resyncedAt = -1

	private formatOf(section: number): string | undefined {
		const sectPr = this.model.sections[section]?.sectPr ?? -1
		const body = this.body

		return body === undefined || sectPr < 0
			? undefined
			: this.engine.resolver.section(this.engine.store.section(body.document, sectPr)).pageNumbering.format
	}
}

// Previous-layout state a relayout for edits compares new pages with (see resynchronise).
interface Resync {
	readonly from: number
	readonly pages: readonly PageDisplayList[]
	readonly anchors: readonly PageAnchor[]
	readonly firstLines: readonly number[]
	readonly snapshots: readonly (PageSnapshot | undefined)[]
	readonly bodyTotals: readonly (PageValues | undefined)[]
	readonly records: readonly (PageRecord | undefined)[]
	readonly bodyNumbers: readonly boolean[]
	readonly renderedBreaks: readonly RenderedBreak[]
	// Pages (by index, as numbered here) still to refresh (DocumentLayout.stale) or to deliver again
	// (DocumentLayout.dirty).
	readonly stale: Set<number>
	readonly redone: ReadonlySet<number>
	readonly sectionPageCounts: readonly number[]
	readonly globalLine: number
	readonly dirty: readonly DirtyBlock[]
	// Index of each dirty block in its section's flow (-1 when absent), found once.
	readonly dirtyIndices: Map<number, number>
	byAnchor: Map<string, number> | undefined
}

// The members of `set` in [start, end).
function within(set: ReadonlySet<number>, start: number, end: number): Set<number> {
	const result = new Set<number>()

	for (const index of set) if (index >= start && index < end) result.add(index)

	return result
}

function usesFlag(output: PageOutput, flag: number): boolean {
	for (let s = 0; s < output.sourceCount; s++) if ((output.sourceFlags(s) & flag) !== 0) return true

	return false
}

function anchorKey(anchor: PageAnchor): string {
	return `${anchor.section}:${anchor.story}:${anchor.part}:${anchor.block}:${anchor.row}:${anchor.offset}`
}

function sameAnchor(a: PageAnchor, b: PageAnchor): boolean {
	return (
		a.section === b.section &&
		a.story === b.story &&
		a.part === b.part &&
		a.block === b.block &&
		a.row === b.row &&
		a.offset === b.offset
	)
}

function sameValues(a: PageValues | undefined, b: PageValues | undefined, pageShift: number): boolean {
	if (a === undefined || b === undefined) return a === b

	return a.page === b.page + pageShift && a.pageFormat === b.pageFormat && a.numPages === b.numPages && a.sectionPages === b.sectionPages
}

function sameMarker(a: ListMarker | null | undefined, b: ListMarker | null | undefined): boolean {
	if (a === undefined || a === null || b === undefined || b === null) return a === b

	return a.numId === b.numId && a.ilvl === b.ilvl && a.value === b.value && a.text === b.text && a.level === b.level
}

function sameCounters(a: CounterSnapshot | undefined, b: CounterSnapshot | undefined): boolean {
	if (a === undefined || b === undefined) return a === b
	if (a.sets.size !== b.sets.size || a.overridesApplied.size !== b.overridesApplied.size) return false

	for (const [id, set] of a.sets) {
		const other = b.sets.get(id)

		if (other === undefined || set.values.length !== other.values.length) return false

		for (let k = 0; k < set.values.length; k++) if (set.values[k] !== other.values[k] || set.active[k] !== other.active[k]) return false
	}

	for (const [id, applied] of a.overridesApplied) if (b.overridesApplied.get(id) !== applied) return false

	return true
}

function sameNotes(a: NoteNumberingSnapshot, b: NoteNumberingSnapshot): boolean {
	const same = (x: NoteNumberingSnapshot["footnotes"], y: NoteNumberingSnapshot["footnotes"]): boolean =>
		x.format === y.format && x.start === y.start && x.restart === y.restart && x.position === y.position

	return (
		a.footnoteNext === b.footnoteNext &&
		a.endnoteNext === b.endnoteNext &&
		same(a.footnotes, b.footnotes) &&
		same(a.endnotes, b.endnotes)
	)
}

function sameEntries<T>(
	a: readonly (readonly [number, T])[],
	b: readonly (readonly [number, T])[],
	same: (x: T, y: T) => boolean
): boolean {
	if (a.length !== b.length) return false

	const map = new Map(b)

	for (const [key, value] of a) {
		if (!map.has(key)) return false

		const other = map.get(key) as T

		if (!same(value, other)) return false
	}

	return true
}

// Whether a page opening now continues exactly as a page of the previous layout did (its anchor matched):
// everything later pages depend on is the same, the page number apart.
function sameState(a: PageSnapshot, b: PageSnapshot): boolean {
	const fa = a.flow
	const fb = b.flow
	const pageShift = a.pageNumber - b.pageNumber

	return (
		a.sectionIndex === b.sectionIndex &&
		a.restart === b.restart &&
		a.sectionFirstPage === b.sectionFirstPage &&
		a.nextStart === b.nextStart &&
		a.bodyLines === b.bodyLines &&
		a.lineSection === b.lineSection &&
		a.endnotesDone === b.endnotesDone &&
		fa.inParagraph === fb.inParagraph &&
		fa.item === fb.item &&
		fa.offset === fb.offset &&
		fa.rows === fb.rows &&
		fa.pendingMark === fb.pendingMark &&
		fa.after === fb.after &&
		fa.contextual === fb.contextual &&
		fa.styleId === fb.styleId &&
		fa.hasPrevious === fb.hasPrevious &&
		fa.breakBeforeDone === fb.breakBeforeDone &&
		fa.noRoom === fb.noRoom &&
		sameValues(fa.values, fb.values, pageShift) &&
		sameMarker(fa.marker, fb.marker) &&
		sameCounters(a.counters, b.counters) &&
		sameNotes(a.notes, b.notes) &&
		a.carried.length === b.carried.length &&
		a.carried.every(
			(note, k) => note.story === b.carried[k]?.story && note.from === b.carried[k]?.from && note.height === b.carried[k]?.height
		) &&
		a.endnotes.length === b.endnotes.length &&
		a.endnotes.every((story, k) => story === b.endnotes[k]) &&
		sameEntries(a.markers, b.markers, sameMarker) &&
		sameEntries(a.assigned, b.assigned, (x, y) => x === y)
	)
}

// A snapshot of a kept page, moved to where the page now is.
function rebased(
	kept: PageSnapshot,
	section: number,
	renumbered: number,
	numberShift: number,
	lineShift: number,
	blockShift: number,
	breakShift: number,
	now: PageSnapshot,
	then: PageSnapshot
): PageSnapshot {
	const counts = kept.sectionPageCounts.slice()

	for (let s = 0; s < counts.length; s++) {
		if (s < section) counts[s] = now.sectionPageCounts[s] ?? 0
		else if (s === section) counts[s] = (counts[s] ?? 0) + (now.sectionPageCounts[s] ?? 0) - (then.sectionPageCounts[s] ?? 0)
	}

	const same = kept.sectionIndex === section
	const shifted = numberShift !== 0 && kept.sectionIndex < renumbered
	const values = kept.flow.values
	const flow =
		(same && blockShift !== 0) || (shifted && values !== undefined)
			? {
					...kept.flow,
					index: same ? kept.flow.index + blockShift : kept.flow.index,
					values: shifted && values !== undefined ? { ...values, page: values.page + numberShift } : values
				}
			: kept.flow

	return {
		...kept,
		pageNumber: shifted ? kept.pageNumber + numberShift : kept.pageNumber,
		globalLine: kept.globalLine + lineShift,
		renderedBreaks: kept.renderedBreaks + breakShift,
		sectionPageCounts: counts,
		flow
	}
}

const LINE_NUMBER_RUN = { styleId: "LineNumber" }

// Halvings of the column bottom when balancing columns (to under a pixel on any page).
const BALANCE_STEPS = 12

// Where a section's part of a page began, to lay it out again.
interface BalanceStart {
	readonly body: BlockFlow
	readonly flow: FlowSnapshot
	readonly mark: OutputMark
	readonly notes: FootnoteState
	readonly exclusions: number
	readonly endnotes: number
	readonly empty: boolean
	readonly start: AreaStart
}

// Layouts again for page count fields in the text, at most.
const MAX_TOTALS_ROUNDS = 3

// Steps a document's paginator: synchronously within a time budget (editing lays out the pages in view in
// the keystroke's task), or in slices from layoutDocument's loop.
export class LayoutDriver {
	private readonly paginator: Paginator
	private readonly layout: DocumentLayout
	private readonly styles: RunStyles
	private readonly now: () => number
	private rounds = 0
	// The run lays out again for edits: page count fields of earlier pages are refreshed when shown.
	private lazy = false
	// Starts the slice loop when it is not running (set by layoutDocument).
	continue: () => Promise<void> = () => Promise.resolve()

	constructor(paginator: Paginator, layout: DocumentLayout, styles: RunStyles, now: () => number) {
		this.paginator = paginator
		this.layout = layout
		this.styles = styles
		this.now = now
	}

	// Restarts at the nearest page layout can restart at, at or before `page` (no later than where a run
	// under way stands); with `dirty`, the run may stop where it meets the previous layout again.
	begin(page: number, dirty: readonly DirtyBlock[] | undefined): number {
		const paginator = this.paginator
		const layout = this.layout
		// A layout still under way without an earlier edit run cannot stop early: it has pages left to lay out.
		const resync = dirty !== undefined && (layout.final || paginator.resyncing)
		const target = layout.final ? page : Math.min(page, paginator.progress)
		const from = paginator.restartPage(target)

		this.rounds = 0
		this.lazy = dirty !== undefined
		paginator.restartAt(from, resync ? dirty : undefined)
		layout.notify(from, from, false)

		return from
	}

	advance(deadline: number, through: number): AdvanceResult {
		const paginator = this.paginator
		const layout = this.layout

		if (layout.final) return "done"

		this.redoMissed()

		for (;;) {
			const step = paginator.step()

			if (step === Step.FINISHED) {
				if (this.styles.hasMissing) return "fonts"
				if (this.redoMissed()) continue

				// Page count fields in the text showed estimates: laid out again from there with the counts found,
				// a bounded number of times (the counts can change again).
				const from = this.rounds < MAX_TOTALS_ROUNDS ? paginator.totalsMismatch() : -1

				if (from >= 0) {
					this.rounds++
					paginator.adoptTotals()

					const start = paginator.restartPage(from)

					paginator.restartAt(start)
					layout.notify(start, start, false)

					continue
				}

				paginator.finish(this.lazy)
				layout.final = true
				layout.notify(0, layout.pages.length, true)

				return "done"
			}

			// Fonts are loaded as soon as a step needed them, whatever the slicing: a layout that missed fonts
			// outside a suspendable flow then continues the same way however it was sliced.
			if (this.styles.hasMissing) return "fonts"

			if (layout.pages.length > through && this.now() >= deadline) return "more"
		}
	}

	prepareFonts(): Promise<void> {
		return this.styles.prepareMissing()
	}

	// Pages closed while fonts were loading are laid out again once they are loaded (a run then comes out the
	// same whenever fonts arrive). Whether it restarted.
	private redoMissed(): boolean {
		const paginator = this.paginator
		const missed = paginator.missedFonts

		if (missed < 0 || this.styles.hasMissing) return false

		const start = paginator.restartPage(Math.min(missed, paginator.progress))

		paginator.restartAt(start, paginator.resyncing ? [] : undefined, false)
		paginator.missedFonts = -1
		this.layout.notify(start, start, false)

		return true
	}

	refresh(page: number): RefreshResult {
		return this.paginator.refresh(page)
	}

	restory(story: number): number {
		const paginator = this.paginator
		const first = paginator.restory(story)

		// The page being laid out started with the story as it was: laid out again.
		if (!this.layout.final && paginator.shows(story)) this.begin(paginator.progress, [])

		return first
	}

	redefine(styles: DocxStyles, numbering: Numbering | undefined): void {
		this.paginator.redefine(styles, numbering)
	}

	// Where the last run met the previous layout again, or -1.
	get resyncedAt(): number {
		return this.paginator.resyncedAt
	}
}

function defaultYield(): Promise<void> {
	return Promise.resolve()
}

// Lays out a document into pages, incrementally: the returned layout fills as pages are laid out (subscribe
// to hear about them; `done` settles at the end). Work is cut into slices of `sliceMs` with the host given a
// turn in between, so the same code runs in a worker or on a UI thread. `fonts` must have the document's
// font table declared; faces load as text needs them.
export function layoutDocument(
	model: DocxModel,
	resolver: PropertyResolver,
	fonts: FontSource,
	measurer: TextMeasurer,
	options: LayoutOptions = {}
): DocumentLayout {
	const cache = new MeasureCache(measurer)
	const table = new StyleTable(cache, options.lineSpacingRule ?? "gdi")
	const decorations = new DecorationTable()
	const runStyles = new RunStyles(table, fonts, decorations)
	const objects = new ObjectTable()
	const layout = new DocumentLayout(model, table, runStyles, objects)
	const engine = new Engine({
		model,
		resolver,
		store: options.store ?? new PropsStore(),
		styles: runStyles,
		decorations,
		objects,
		sizer: options.sizer,
		notes: new NoteNumbering(model),
		view: options.view ?? RevisionView.Final,
		showHidden: options.showHidden ?? false,
		compat: layoutCompat(model.settings),
		tuning: { ...DEFAULT_TUNING, ...options.tuning },
		recordMaps: options.recordMaps ?? false,
		...(options.edits === undefined ? {} : { edits: options.edits })
	})
	const paginator = new Paginator(engine, layout, options.estimatedPages, options.maxPages ?? 100000)
	const slice = options.sliceMs ?? 8
	const now = options.now ?? Date.now
	const pause = options.yield ?? defaultYield
	const driver = new LayoutDriver(paginator, layout, runStyles, now)

	// Runs the paginator to the end (or cancellation), from wherever it stands; edits may move it back
	// between slices (it then carries on with their run).
	let running = false
	const drive = async (scan: boolean, later = false): Promise<void> => {
		running = true

		try {
			if (scan) {
				paginator.scanStories()
				await runStyles.prepareMissing()
			}

			// Continuing an edit's layout: the edit's own task has done its share already.
			if (later) await pause()

			for (;;) {
				if (layout.cancelled || options.signal?.aborted === true) {
					layout.cancelled = true

					return
				}

				const result = driver.advance(now() + slice, -1)

				if (result === "done") break
				if (result === "fonts") await runStyles.prepareMissing()

				await pause()
			}
		} catch (error) {
			layout.error = error

			throw error
		} finally {
			running = false
		}
	}

	driver.continue = (): Promise<void> => {
		if (!running) layout.done = drive(false, true)

		return layout.done
	}

	layout.attach(
		async (page: number) => {
			const from = driver.begin(page, undefined)

			await drive(false)

			return from
		},
		() => running,
		driver
	)
	layout.done = drive(true)

	return layout
}
