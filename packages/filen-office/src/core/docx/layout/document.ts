import type { LineSpacingRule } from "../../fonts/index"
import { Exclusions, LINE_FLOATS, LINE_INTS, LineInt, MeasureCache, StyleTable, type TextMeasurer } from "../../text/index"
import type { DocxModel, Section } from "../model/document"
import { Item, RevisionView } from "../model/kinds"
import type { StoryPart } from "../model/part"
import { NumberingCounters } from "../numbering/counter"
import { PropsStore } from "../props/store"
import type { ResolvedSectionProps } from "../props/types"
import { resolveSection, type PropertyResolver } from "../styles/resolve"
import type { PageValues, StoryContext } from "./content"
import { numberText } from "./fields"
import { Area, AreaStart, BlockFlow, DEFAULT_TUNING, Engine, FlowStatus, nextToken, type LayoutTuning, type PageFrame } from "./flow"
import { layoutCompat, pageGeometry, type LayoutCompat, type PageGeometry } from "./geometry"
import { FootnoteArea, NoteNumbering } from "./notes"
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
}

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
	// The top-level block (paragraph or table item of the story) the page starts in, -1 after the last.
	readonly block: number
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

const Step = { PROGRESS: 0, PAGE: 1, SUSPEND: 2, FINISHED: 3 } as const

type Step = (typeof Step)[keyof typeof Step]

// A page whose header or footer showed estimated page counts: what to lay out again once they are known.
interface PendingTotals {
	readonly index: number
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
	private readonly bodyStory: StoryContext | undefined
	private readonly compat: LayoutCompat
	private readonly maxPages: number
	private readonly page = new PageOutput()
	private readonly header = new PageOutput()
	private readonly footer = new PageOutput()
	private readonly exclusions = new Exclusions()
	private readonly footnotes: FootnoteArea
	private readonly area = new Area()
	private readonly storyCounters = new Map<number, NumberingCounters | undefined>()
	private readonly pending: PendingTotals[] = []
	private readonly sectionPageCounts: number[] = []
	private readonly sectionEstimates: number[] = []
	private readonly estimate: number
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
	private sectionTop = 0
	private bodyTop = 0
	private bodyBottom = 0
	private contentBottom = 0
	private values: PageValues = { page: 1, pageFormat: undefined, numPages: 1, sectionPages: 1 }
	private frame: PageFrame = { width: 816, height: 1056, marginLeft: 96, marginRight: 96, marginTop: 96, marginBottom: 96, odd: true }
	private headerStory = -1
	private footerStory = -1
	private pageSection = 0
	// The properties of the section the open page belongs to.
	private pageProps: ResolvedSectionProps = resolveSection({})
	private anchor: PageAnchor = { section: 0, story: -1, block: -1, row: -1, offset: 0 }
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

	// Builds every header, footer and note once so the fonts they need are known before layout starts.
	scanStories(): void {
		const values = this.values

		for (const story of this.model.stories) {
			if (story.kind !== "header" && story.kind !== "footer" && story.kind !== "footnote" && story.kind !== "endnote") continue

			const context = this.storyContext(story.index)

			if (context !== undefined) this.engine.scan({ ...context, counters: undefined }, story.item, values)
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

		if (!this.pageOpen) this.openPage()

		const flow = this.flows[this.flowIndex]

		if (flow === undefined) return this.nextSection()

		const status = flow.run(this.area)

		this.contentBottom = Math.max(this.contentBottom, flow.y)

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

	private startSection(index: number): boolean {
		const section = this.model.sections[index]
		const body = this.body
		const story = this.bodyStory

		if (section === undefined || body === undefined || story === undefined) return false

		this.sectionIndex = index
		this.section = this.resolved(index)
		this.flows = [new BlockFlow(this.engine, story, Region.BODY, true).setRange(body, section.start, section.end)]
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
			this.startSection(index)
			this.geometry = pageGeometry(next, this.compat, this.frame.odd)
			this.engine.gridPitch = this.geometry.gridPitch
			this.sectionTop = this.contentBottom
			this.column = 0
			this.setArea(0, this.sectionTop, this.sectionTop <= this.bodyTop + EPSILON)
			this.area.start = AreaStart.BREAK

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

		const stories = engine.endnotes.splice(0)

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

	private nextColumn(start: AreaStart): Step {
		const geometry = this.geometry

		if (geometry !== undefined && this.column + 1 < geometry.columnX.length) {
			this.column++
			this.setArea(this.column, this.sectionTop, this.sectionTop <= this.bodyTop + EPSILON)
			this.area.start = start

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

		const flow = this.flows[this.flowIndex]
		const position = flow?.position()

		this.anchor = {
			section: this.sectionIndex,
			story: flow?.story.story ?? -1,
			block: position?.block ?? -1,
			row: position?.row ?? -1,
			offset: position?.offset ?? 0
		}

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

		this.bodyTop = geometry.topFixed ? geometry.top : Math.max(geometry.top, headerBottom)
		this.bodyBottom = geometry.bottomFixed ? geometry.bottom : Math.min(geometry.bottom, footerTop)

		if (this.bodyBottom - this.bodyTop < 24) this.bodyBottom = this.bodyTop + 24

		this.footnotes.startPage(geometry.left, textWidth, this.values, this.frame)
		this.column = 0
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
		this.footnotes.close(page, this.bodyBottom, this.contentBottom)

		if (geometry.separator) {
			for (let c = 1; c < geometry.columnX.length; c++) {
				const previousEnd = (geometry.columnX[c - 1] ?? 0) + (geometry.columnWidths[c - 1] ?? 0)
				const x = geometry.left + (previousEnd + (geometry.columnX[c] ?? 0)) / 2

				page.addGraphic(
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

		const totals = this.usesTotals(this.header) || this.usesTotals(this.footer)

		if (totals) {
			this.pending.push({
				index,
				body: page.mark(),
				header: this.headerStory,
				footer: this.footerStory,
				geometry,
				values: this.values,
				frame: this.frame,
				info
			})
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
		for (let s = 0; s < output.sourceCount; s++) if ((output.sourceFlags(s) & SourceFlag.TOTAL_FIELDS) !== 0) return true

		return false
	}

	private trimMemos(): void {
		const markers = this.engine.markers

		if (markers.size < 8192) return

		const flow = this.flows[this.flowIndex]
		const keepFrom = flow === undefined ? Infinity : (flow.blocks[Math.max(0, flow.index - 1)] ?? 0)
		const bodyPart = this.body?.index ?? -1

		for (const key of markers.keys()) {
			const part = Math.floor(key / 0x4000000)
			const item = key - part * 0x4000000

			if (part === bodyPart && item < keepFrom) markers.delete(key)
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
	private lineNumbers(page: PageOutput): void {
		const numbering = this.pageProps.lineNumbering
		const geometry = this.geometry

		if (numbering === undefined || geometry === undefined || this.body === undefined) return

		const countBy = Math.max(1, numbering.countBy ?? 1)
		const distance = numbering.distance === undefined ? 18 : twips(numbering.distance)
		const text = page.text
		const engine = this.engine
		const format = engine.resolver.paragraph({})
		const style = engine.styles.style(engine.resolver.run(format, { styleId: "LineNumber" }), "ascii")
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

	// Page count fields: pages whose headers or footers used the estimate are laid out again.
	finish(): void {
		const layout = this.layout
		const total = layout.pages.length

		for (const entry of this.pending) {
			const sectionPages = this.sectionPageCounts[entry.info.section] ?? 1

			if (entry.values.numPages === total && entry.values.sectionPages === sectionPages) continue

			const values: PageValues = { ...entry.values, numPages: total, sectionPages }
			const page = this.page
			const geometry = entry.geometry

			this.values = values
			this.frame = entry.frame
			this.geometry = geometry
			page.reset()
			this.header.reset()
			this.footer.reset()

			if (entry.header >= 0) this.layoutStory(entry.header, this.header, Region.HEADER, geometry.header)

			if (entry.footer >= 0) {
				const height = this.layoutStory(entry.footer, this.footer, Region.FOOTER, 0)

				this.footer.move(new OutputMark(), null, 0, geometry.height - geometry.footer - height)
			}

			const list = layout.pages[entry.index]

			if (list === undefined) continue

			page.append(list, 0, 0, -Infinity, Infinity, null, entry.body)
			page.append(this.header, 0, 0)
			page.append(this.footer, 0, 0)
			layout.pages[entry.index] = new PageDisplayList(page, entry.info)
			layout.dirty.add(entry.index)
		}

		this.pending.length = 0
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
		tuning: { ...DEFAULT_TUNING, ...options.tuning }
	})
	const paginator = new Paginator(engine, layout, options.estimatedPages, options.maxPages ?? 100000)
	const slice = options.sliceMs ?? 8
	const now = options.now ?? Date.now
	const pause = options.yield ?? defaultYield

	layout.done = (async () => {
		try {
			paginator.scanStories()
			await runStyles.prepareMissing()

			for (;;) {
				if (layout.cancelled || options.signal?.aborted === true) {
					layout.cancelled = true

					return
				}

				const started = now()
				let finished = false

				for (;;) {
					const step = paginator.step()

					if (step === Step.FINISHED) {
						finished = true

						break
					}

					// Fonts are loaded as soon as a step needed them, whatever the slicing: a layout that missed fonts
					// outside a suspendable flow then continues the same way however it was sliced.
					if (runStyles.hasMissing) {
						await runStyles.prepareMissing()

						break
					}

					if (now() - started >= slice) break
				}

				if (finished) break

				await pause()
			}

			paginator.finish()
			layout.final = true
			layout.notify(0, layout.pages.length, true)
		} catch (error) {
			layout.error = error

			throw error
		}
	})()

	return layout
}
