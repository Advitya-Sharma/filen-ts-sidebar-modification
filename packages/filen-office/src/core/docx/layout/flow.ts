import {
	FRAGMENT_FLOATS,
	FRAGMENT_INTS,
	FRAGMENT_KIND_MASK,
	FragmentKind,
	LINE_FLOATS,
	LINE_INTS,
	LayoutCursor,
	LayoutStatus,
	LineInt,
	LineRule,
	ParagraphLayouter,
	PreparedParagraph,
	TextLayout,
	Wrap,
	kinsokuLists,
	kinsokuOf,
	type Exclusions,
	type PrepareOptions
} from "../../text/index"
import type { DocxModel } from "../model/document"
import { Item, RevisionView } from "../model/kinds"
import type { StoryPart } from "../model/part"
import type { ListMarker, NumberingCounters } from "../numbering/counter"
import type { PropsStore } from "../props/store"
import type { FrameProps, ResolvedBorder, ResolvedParagraphBorders, ResolvedParagraphProps } from "../props/types"
import type { CellFormat, ParagraphFormat, PropertyResolver } from "../styles/resolve"
import { BuiltParagraph, ContentBuilder, type PageValues, type StoryContext } from "./content"
import { FIXED_AUTO_SPACING, Geometry, paragraphGeometry, spaceAfter, spaceBefore, type AutoSpacing, type LayoutCompat } from "./geometry"
import type { NoteNumbering } from "./notes"
import { ObjectKind, type FloatSpec, type ObjectSizer, type ObjectTable } from "./objects"
import { PageGraphic, Layer, OBJECT_FLOATS, OBJECT_INTS, OutputMark, PageOutput, Region, SourceFlag } from "./output"
import { REF_SPAN, refItem, refPart, type LayoutEdits } from "./edits"
import { borderWidth, revisionAuthors, type DecorationTable, type RunStyles } from "./styles"
import { TableLayout } from "./table"
import { EPSILON, twips } from "./units"

export const FlowStatus = {
	DONE: 0,
	// The area is full; continue in the next one.
	OVERFLOW: 1,
	PAGE_BREAK: 2,
	COLUMN_BREAK: 3,
	// Fonts must be loaded first; nothing of the pending block was placed.
	SUSPEND: 4
} as const

export type FlowStatus = (typeof FlowStatus)[keyof typeof FlowStatus]

// Choices Word does not document, settled by calibration against Word's own page breaks.
export interface LayoutTuning {
	// Space before of the paragraph at the top of a page: "overflow" drops it when the page started because
	// the previous one was full, "breaks" also after page and section breaks, "never" always keeps it.
	// "compat": "breaks" in Word 2013+ documents (compatibility mode 15) and under w:suppressSpBfAfterPgBrk,
	// "overflow" otherwise.
	readonly dropSpaceAtPageTop: "overflow" | "breaks" | "compat" | "never"
	// Space between paragraphs: the sum of space after and space before, the larger of the two, or "auto": the
	// larger unless the document sets w:doNotUseHTMLParagraphAutoSpacing. Microsoft documents neither (its
	// statements cover auto spacing values only); "auto" breaks most often where Word recorded its breaks.
	readonly paragraphSpacing: "sum" | "max" | "auto"
	// What a line spacing multiple stretches: the text of a line only, or everything on it (inline objects).
	readonly multipleStretches: "text" | "all"
	// Whether a note's (or the note separator's) last space after counts in the footnote area.
	readonly noteTrailingSpace: "all" | "notes" | "none"
	// HTML auto spacing before and after a paragraph, in points (unless w:doNotUseHTMLParagraphAutoSpacing).
	readonly htmlAutoSpacing: readonly [number, number]
	// A paragraph ending above footnotes: "fit" when its space after must fit above them too, "overhang"
	// when only its lines must (as at a page bottom without notes).
	readonly noteSpaceAfter: "fit" | "overhang"
}

export const DEFAULT_TUNING: LayoutTuning = {
	dropSpaceAtPageTop: "breaks",
	paragraphSpacing: "auto",
	multipleStretches: "text",
	noteTrailingSpace: "all",
	htmlAutoSpacing: [0, 12],
	noteSpaceAfter: "fit"
}

// Why an area's page or column began.
export const AreaStart = { DOCUMENT: 0, OVERFLOW: 1, BREAK: 2 } as const

export type AreaStart = (typeof AreaStart)[keyof typeof AreaStart]

// Page geometry anchored objects position against.
export interface PageFrame {
	width: number
	height: number
	marginLeft: number
	marginRight: number
	marginTop: number
	marginBottom: number
	odd: boolean
}

// Notes referenced on the page: reserves room at the page bottom.
export interface FootnoteSink {
	readonly reserved: number
	// A note on the page continues on the next one: the page is full.
	readonly split: boolean
	mark(): number
	rollback(mark: number): void
	// Reserves room for a note referenced on a line ending at `lineBottom`; false when the note does not fit
	// above `limit` (unless `force`, which splits it to the next page).
	add(story: number, lineBottom: number, limit: number, force: boolean): boolean
}

// Where a flow places content: a column of a page, a table cell, a header, a note or a text box.
export class Area {
	page: PageOutput = new PageOutput()
	region: Region = Region.BODY
	left = 0
	width = 0
	top = 0
	bottom = Infinity
	// Nothing is above the flow's position in this column: the first line is placed whatever its height.
	empty = true
	// Body flow: page and column breaks, keep and widow rules.
	breaks = false
	inTable = false
	start: AreaStart = AreaStart.DOCUMENT
	exclusions: Exclusions | null = null
	footnotes: FootnoteSink | null = null
	values: PageValues = { page: 1, pageFormat: undefined, numPages: 1, sectionPages: 1 }
	frame: PageFrame = { width: 816, height: 1056, marginLeft: 96, marginRight: 96, marginTop: 96, marginBottom: 96, odd: true }
	// Changes whenever the area is reused for another column or page.
	token = 0

	copy(from: Area): this {
		this.page = from.page
		this.region = from.region
		this.left = from.left
		this.width = from.width
		this.top = from.top
		this.bottom = from.bottom
		this.empty = from.empty
		this.breaks = from.breaks
		this.inTable = from.inTable
		this.start = from.start
		this.exclusions = from.exclusions
		this.footnotes = from.footnotes
		this.values = from.values
		this.frame = from.frame
		this.token = from.token

		return this
	}
}

let tokens = 0

export function nextToken(): number {
	return ++tokens
}

export interface EngineInit {
	readonly model: DocxModel
	readonly resolver: PropertyResolver
	readonly store: PropsStore
	readonly styles: RunStyles
	readonly decorations: DecorationTable
	readonly objects: ObjectTable
	readonly sizer: ObjectSizer | undefined
	readonly notes: NoteNumbering
	readonly view: RevisionView
	readonly showHidden: boolean
	readonly compat: LayoutCompat
	readonly tuning: LayoutTuning
	readonly recordMaps?: boolean
	readonly edits?: LayoutEdits
}

// Services shared by every flow of one document layout.
export class Engine {
	readonly model: DocxModel
	readonly resolver: PropertyResolver
	readonly store: PropsStore
	readonly styles: RunStyles
	readonly decorations: DecorationTable
	readonly objects: ObjectTable
	readonly notes: NoteNumbering
	readonly view: RevisionView
	readonly compat: LayoutCompat
	readonly tuning: LayoutTuning
	readonly content: ContentBuilder
	readonly layouter: ParagraphLayouter
	readonly markers = new Map<number, ListMarker | null>()
	// Body content went onto the page drawn with stand-ins for fonts still loading (a table row that could not
	// wait for them): the page is laid out again once they loaded. The paginator reads and clears it.
	standIns = false
	// Placements read so far that depend on whether the page is odd or even (inside/outside alignment): a page
	// that read one is laid out again rather than kept when its number changes parity.
	parityReads = 0
	// Endnote stories in reference order.
	readonly endnotes: number[] = []
	private readonly endnoteSet = new Set<number>()
	private readonly scratch = new TextLayout()
	private readonly scratchCursor = new LayoutCursor()
	private readonly scanBuilt = new BuiltParagraph()
	gridPitch = 0
	readonly autoSpacing: AutoSpacing
	// Unsaved edits laid out in place of their source (see LayoutEdits), read whenever a flow is set up.
	readonly edits: LayoutEdits | undefined

	constructor(init: EngineInit) {
		this.model = init.model
		this.resolver = init.resolver
		this.store = init.store
		this.styles = init.styles
		this.decorations = init.decorations
		this.objects = init.objects
		this.notes = init.notes
		this.view = init.view
		this.compat = init.compat
		this.tuning = init.tuning
		this.edits = init.edits
		this.layouter = new ParagraphLayouter(init.styles.table)
		this.autoSpacing = init.compat.htmlSpacing
			? { before: (init.tuning.htmlAutoSpacing[0] * 4) / 3, after: (init.tuning.htmlAutoSpacing[1] * 4) / 3 }
			: FIXED_AUTO_SPACING
		this.content = new ContentBuilder({
			model: init.model,
			resolver: init.resolver,
			store: init.store,
			styles: init.styles,
			objects: init.objects,
			sizer: init.sizer,
			notes: init.notes,
			stretchObjects: init.tuning.multipleStretches === "all",
			view: init.view,
			showHidden: init.showHidden,
			markers: this.markers,
			recordMaps: init.recordMaps ?? false,
			memoKey: (part, item) => this.memoKey(part, item),
			authors: revisionAuthors(init.model).index
		})
	}

	memoKey(part: number, item: number): number {
		return this.edits?.memoKey?.(part, item) ?? part * REF_SPAN + item
	}

	// The part a ref of a flow's blocks points into (the model's parts plus overlays).
	part(index: number): StoryPart | undefined {
		return this.model.parts[index]
	}

	// Keeps the first `count` endnotes referenced (layout restarting at a page).
	truncateEndnotes(count: number): void {
		for (const story of this.endnotes.splice(count)) this.endnoteSet.delete(story)
	}

	// Restarts with `stories` referenced and not placed yet.
	setEndnotes(stories: readonly number[]): void {
		this.truncateEndnotes(0)

		for (const story of stories) this.addEndnote(story)
	}

	// The endnotes referenced since the last call, in order; they may be referenced again.
	takeEndnotes(): number[] {
		const stories = this.endnotes.splice(0)

		for (const story of stories) this.endnoteSet.delete(story)

		return stories
	}

	addEndnote(story: number): void {
		if (this.endnoteSet.has(story)) return

		this.endnoteSet.add(story)
		this.endnotes.push(story)
	}

	// Whether the rest of a paragraph from `cursor` takes more than one row: 1 or 2 (laid out off-page, one row
	// at most, so long paragraphs cost nothing extra).
	remainingRows(p: PreparedParagraph, g: Geometry, cursor: LayoutCursor): number {
		const scratch = this.scratch
		const copy = this.scratchCursor
		const exclusions = g.exclusions

		copy.item = cursor.item
		copy.offset = cursor.offset
		copy.rows = cursor.rows
		copy.pendingMark = cursor.pendingMark
		copy.y = 0
		scratch.reset()
		g.exclusions = null
		// A bottom at the top lets only the forced first row through.
		const status = this.layouter.layout(p, g, scratch, copy, 0, 0, true)

		g.exclusions = exclusions

		return status === LayoutStatus.DONE ? scratch.lineCount : 2
	}

	// Builds every paragraph under `container` without laying it out, so the fonts it needs are known.
	scan(story: StoryContext, container: number, values: PageValues): void {
		const part = story.part
		const end = part.end(container)
		const edited = this.edits?.stories?.get(story.story)
		const substitutes = this.edits?.substitutes

		// The story's blocks as laid out: its edited blocks from the parts they were written to.
		if (edited === undefined && (substitutes === undefined || substitutes.size === 0)) {
			this.scanItems(story, part, container + 1, end, values)

			return
		}

		const blocks: number[] = []

		if (edited !== undefined) {
			for (const value of edited) blocks.push(value)
		} else {
			for (let item = container + 1; item < end;) {
				const kind = part.kinds[item] ?? 0

				if (kind === Item.Paragraph || kind === Item.Table) {
					blocks.push(part.index * REF_SPAN + item)
					item = part.end(item)
				} else {
					item++
				}
			}
		}

		for (const block of blocks) {
			const laid = substitutes?.get(block) ?? block
			const from = this.model.parts[refPart(laid)]

			if (from !== undefined) this.scanItems(story, from, refItem(laid), from.end(refItem(laid)), values)
		}
	}

	private scanItems(story: StoryContext, part: StoryPart, start: number, end: number, values: PageValues): void {
		const content = this.content

		for (let item = start; item < end; item++) {
			if (part.kinds[item] !== Item.Paragraph) continue

			const format = content.formatOf(part, item, undefined)

			content.build(story, [part], [item], [format], 600, values, this.scanBuilt)
		}
	}

	// Lays out a story's blocks into a box (text boxes, frames): no breaks, no notes.
	layoutBox(
		story: StoryContext,
		region: Region,
		page: PageOutput,
		values: PageValues,
		frame: PageFrame,
		left: number,
		top: number,
		width: number,
		// Count the last paragraph's space after in the height.
		trailing = true
	): number {
		const part = story.part
		const root = this.model.stories[story.story]

		if (root === undefined || width <= 0) return 0

		const flow = new BlockFlow(this, story, region, false)
		const area = new Area()
		const edited = this.edits?.stories?.get(story.story)

		if (edited === undefined) flow.setContainer(part, root.item)
		else flow.setRefs(part, edited, 0, edited.length)
		area.page = page
		area.region = region
		area.left = left
		area.width = width
		area.top = top
		area.bottom = Infinity
		area.values = values
		area.frame = frame
		area.token = nextToken()
		flow.run(area)

		return (trailing ? flow.contentBottom() : flow.y) - top
	}
}

// A flow's state between two areas (see BlockFlow.snapshot).
export interface FlowSnapshot {
	readonly index: number
	readonly inParagraph: boolean
	readonly item: number
	readonly offset: number
	readonly rows: number
	readonly pendingMark: boolean
	readonly after: number
	readonly contextual: boolean
	readonly styleId: string | undefined
	readonly hasPrevious: boolean
	readonly breakBeforeDone: boolean
	readonly noRoom: number
	readonly values: PageValues | undefined
	readonly marker: ListMarker | null | undefined
	// A keep-with-next chain is under way (one that can no longer move: it began in an earlier area).
	readonly chained: boolean
	// BlockFlow.moves, each block relative to `index`.
	readonly moves: readonly number[]
}

// A run of consecutive paragraphs with identical borders, drawn as one box.
interface BorderBox {
	borders: ResolvedParagraphBorders | undefined
	shading: number
	left: number
	right: number
	top: number
	bottom: number
	// Bottom of each paragraph but the last, where between-borders go.
	readonly joins: number[]
	open: boolean
}

// Saved state to return to when a keep-with-next chain must move to the next page.
interface Chain {
	index: number
	y: number
	after: number
	contextual: boolean
	style: string | undefined
	hasPrevious: boolean
	empty: boolean
	readonly mark: OutputMark
	exclusions: number
	footnotes: number
	token: number
}

// A chain carried over from an earlier area (a page break inside it): it can no longer move, it only keeps
// the blocks after it from starting chains of their own.
const STALE_CHAIN: Chain = {
	index: 0,
	y: 0,
	after: 0,
	contextual: false,
	style: undefined,
	hasPrevious: false,
	empty: true,
	mark: new OutputMark(),
	exclusions: 0,
	footnotes: 0,
	token: -1
}

const NO_MOVES: readonly number[] = []

const SCRATCH_ITEMS: number[] = []
const SCRATCH_PARTS: StoryPart[] = []
const SCRATCH_FORMATS: ParagraphFormat[] = []

// Lays out a container's blocks (paragraphs and tables) into areas, resumably: run() returns when an
// area is full and continues where it stopped in the next area.
export class BlockFlow {
	readonly engine: Engine
	readonly story: StoryContext
	readonly region: Region
	// Fonts that are missing suspend the flow (otherwise layout goes on with what resolves).
	readonly suspendable: boolean
	cell: CellFormat | undefined = undefined
	// The container's part (blocks may be laid out from overlay parts instead: see parts).
	part: StoryPart | undefined = undefined
	// Items of the blocks, and the parts they belong to.
	blocks = new Int32Array(16)
	parts = new Int32Array(16)
	count = 0
	// The last block whose content has been built (keep chains and frames build ahead of index).
	furthest = -1
	index = 0
	y = 0
	// Space after of the last paragraph placed (pending until the next block).
	after = 0
	private contextual = false
	private styleId: string | undefined = undefined
	private hasPrevious = false

	private inParagraph = false
	private groupEnd = 0
	private readonly built = new BuiltParagraph()
	private readonly prepared = new PreparedParagraph()
	private readonly cursor = new LayoutCursor()
	private readonly saved = new LayoutCursor()
	private readonly geometry = new Geometry()
	private breakBeforeDone = false
	private floatsPlaced = false
	private source = -1
	private token = -1
	private areaPlaced = 0
	// Areas in a row where floats left no room for the pending paragraph.
	private noRoom = 0
	// Output and obstacles before the pending paragraph's anchored objects, to take them back when the
	// paragraph moves on without a line here.
	private readonly floatMark = new OutputMark()
	private floatExclusions = 0
	private readonly floatArea = new Area()
	// Lays out the paragraphs of a frame (which do not float again).
	private inFrame = false
	private table: TableLayout | undefined = undefined
	// The page values the pending paragraph was built with.
	private builtValues: PageValues | undefined = undefined
	private chain: Chain | undefined = undefined
	// Keep-with-next chains that moved off a page while the block they ran out of room at is still to come:
	// pairs of pages back to that page and the block's index. An edit to that block or one before it may let
	// the chain fit there again; blocks after it were never tried there.
	private moves: number[] = []
	private box: BorderBox = { borders: undefined, shading: 0, left: 0, right: 0, top: 0, bottom: 0, joins: [], open: false }

	constructor(engine: Engine, story: StoryContext, region: Region, suspendable: boolean) {
		this.engine = engine
		this.story = story
		this.region = region
		this.suspendable = suspendable
	}

	// Blocks of `container` (a story root or a cell): paragraphs and tables, wrappers entered.
	setContainer(part: StoryPart, container: number, cell?: CellFormat): this {
		return this.setRange(part, container + 1, part.end(container), cell)
	}

	setRange(part: StoryPart, start: number, end: number, cell?: CellFormat): this {
		const kinds = part.kinds
		const view = this.engine.view
		const substitutes = this.substitutes()

		this.part = part
		this.cell = cell
		this.count = 0

		for (let item = start; item < end;) {
			const kind = kinds[item] ?? 0

			if (kind === Item.Paragraph || kind === Item.Table) {
				this.push(part.index, item, substitutes)
				item = part.end(item)
			} else if (
				kind === Item.Sdt ||
				kind === Item.CustomXml ||
				kind === Item.AlternateContent ||
				((kind === Item.Insertion || kind === Item.MoveTo) && view !== RevisionView.Original) ||
				((kind === Item.Deletion || kind === Item.MoveFrom) && view !== RevisionView.Final)
			) {
				item++
			} else {
				item = part.end(item)
			}
		}

		this.reset()

		return this
	}

	// Blocks given as refs (the body's blocks once its structure was edited): refs[from, to).
	setRefs(part: StoryPart, refs: Float64Array, from: number, to: number): this {
		const substitutes = this.substitutes()

		this.part = part
		this.cell = undefined
		this.count = 0

		for (let k = from; k < to; k++) {
			const value = refs[k] ?? 0

			this.push(refPart(value), refItem(value), substitutes)
		}

		this.reset()

		return this
	}

	// Blocks [from, to) of another flow, as it lays them out.
	copyBlocks(flow: BlockFlow, from: number, to: number): this {
		this.part = flow.part
		this.cell = flow.cell
		this.count = 0

		for (let k = from; k < to; k++) this.push(flow.parts[k] ?? -1, flow.blocks[k] ?? -1, undefined)

		this.reset()

		return this
	}

	private substitutes(): ReadonlyMap<number, number> | undefined {
		const substitutes = this.engine.edits?.substitutes

		return substitutes === undefined || substitutes.size === 0 ? undefined : substitutes
	}

	private push(part: number, item: number, substitutes: ReadonlyMap<number, number> | undefined): void {
		const count = this.count

		if (count === this.blocks.length) {
			const blocks = new Int32Array(count * 2)
			const parts = new Int32Array(count * 2)

			blocks.set(this.blocks)
			parts.set(this.parts)
			this.blocks = blocks
			this.parts = parts
		}

		const substitute = substitutes?.get(part * REF_SPAN + item)

		this.blocks[count] = substitute === undefined ? item : refItem(substitute)
		this.parts[count] = substitute === undefined ? part : refPart(substitute)
		this.count = count + 1
	}

	// The part block k is laid out from.
	partAt(k: number): StoryPart | undefined {
		return this.engine.model.parts[this.parts[k] ?? -1]
	}

	// The ref of block k as laid out.
	refAt(k: number): number {
		return (this.parts[k] ?? 0) * REF_SPAN + (this.blocks[k] ?? 0)
	}

	reset(): void {
		this.index = 0
		this.furthest = -1
		this.y = 0
		this.after = 0
		this.contextual = false
		this.styleId = undefined
		this.hasPrevious = false
		this.inParagraph = false
		this.table = undefined
		this.chain = undefined
		this.moves.length = 0
		this.token = -1
		this.box.open = false
	}

	get done(): boolean {
		return this.index >= this.count && !this.inParagraph && this.table === undefined
	}

	// The flow's state, to continue from it later; undefined inside a table (layout restarts before it).
	snapshot(): FlowSnapshot | undefined {
		if (this.table !== undefined || this.part === undefined) return undefined

		const cursor = this.cursor
		const first = this.blocks[this.index] ?? -1

		return {
			index: this.index,
			inParagraph: this.inParagraph,
			item: cursor.item,
			offset: cursor.offset,
			rows: cursor.rows,
			pendingMark: cursor.pendingMark,
			after: this.after,
			contextual: this.contextual,
			styleId: this.styleId,
			hasPrevious: this.hasPrevious,
			breakBeforeDone: this.breakBeforeDone,
			noRoom: this.noRoom,
			values: this.builtValues,
			marker: this.engine.markers.get(this.engine.memoKey(this.parts[this.index] ?? 0, first)),
			chained: this.chain !== undefined,
			moves: this.moves.length === 0 ? NO_MOVES : this.moves.map((value, k) => (k % 2 === 0 ? value : value - this.index))
		}
	}

	// Continues from a snapshot: a paragraph that was under way is built again as it was and resumes at the
	// same line. Areas start afresh.
	restore(snapshot: FlowSnapshot, area: Area): void {
		const engine = this.engine
		const part = this.part

		this.reset()
		this.index = snapshot.index
		this.after = snapshot.after
		this.contextual = snapshot.contextual
		this.styleId = snapshot.styleId
		this.hasPrevious = snapshot.hasPrevious
		this.noRoom = snapshot.noRoom
		for (let k = 0; k < snapshot.moves.length; k += 2)
			this.moves.push(snapshot.moves[k] ?? 0, (snapshot.moves[k + 1] ?? 0) + snapshot.index)

		if (snapshot.chained) this.chain = STALE_CHAIN

		if (!snapshot.inParagraph || part === undefined) return

		const first = this.blocks[this.index] ?? -1

		if (snapshot.marker !== undefined) engine.markers.set(engine.memoKey(this.parts[this.index] ?? 0, first), snapshot.marker)

		const { items, parts, formats } = this.group()
		const values = snapshot.values ?? area.values

		engine.content.build(this.story, parts, items, formats, area.width, values, this.built)
		this.builtValues = values

		const props = this.built.format?.props

		if (props === undefined) return

		engine.layouter.prepare(engine.content.builder, this.prepareOptions(props), this.prepared)
		this.inParagraph = true
		this.groupEnd = this.index + items.length
		this.cursor.item = snapshot.item
		this.cursor.offset = snapshot.offset
		this.cursor.rows = snapshot.rows
		this.cursor.pendingMark = snapshot.pendingMark
		this.breakBeforeDone = snapshot.breakBeforeDone
		this.floatsPlaced = true
		this.source = -1
	}

	// A paragraph with PAGE fields continuing on another page is built again with that page's number; the
	// cursor keeps its place in the text, moved by the fields before it that changed length.
	private renumber(area: Area): void {
		const engine = this.engine
		const built = this.built
		const cursor = this.cursor
		const prepared = this.prepared
		const old = built.pageFields.slice()
		let offset = cursor.offset >= 0 ? cursor.offset : (prepared.start[cursor.item] ?? prepared.text.length)
		const { items, parts, formats } = this.group()

		engine.content.build(this.story, parts, items, formats, area.width, area.values, built)
		this.builtValues = area.values

		const props = built.format?.props

		if (props === undefined) return

		const fields = built.pageFields
		let shift = 0

		for (let k = 0; k + 1 < old.length && k + 1 < fields.length; k += 2) {
			if ((old[k + 1] ?? 0) > offset) break

			shift += (fields[k + 1] ?? 0) - (fields[k] ?? 0) - ((old[k + 1] ?? 0) - (old[k] ?? 0))
		}

		offset += shift
		engine.layouter.prepare(engine.content.builder, this.prepareOptions(props), prepared)

		let item = 0

		while (item + 1 < prepared.count && (prepared.start[item + 1] ?? 0) <= offset) item++

		cursor.item = offset >= prepared.text.length ? prepared.count : item
		cursor.offset = cursor.item < prepared.count && (prepared.start[item] ?? 0) !== offset ? offset : -1
	}

	// Where the flow continues: the next block, the table row it is in, and the text offset inside a
	// paragraph it is in the middle of.
	position(): { block: number; part: number; index: number; row: number; offset: number } {
		const block = this.index < this.count ? (this.blocks[this.index] ?? -1) : -1
		const part = this.index < this.count ? (this.parts[this.index] ?? -1) : -1
		const cursor = this.cursor
		const offset = !this.inParagraph
			? 0
			: cursor.offset >= 0
				? cursor.offset
				: (this.prepared.start[cursor.item] ?? this.prepared.text.length)

		return { block, part, index: this.index, row: this.table?.currentRow ?? -1, offset: cursor.rows === 0 ? 0 : offset }
	}

	// Bottom of the content including the last paragraph's space after (cells grow by it).
	contentBottom(): number {
		return this.y + this.after
	}

	run(area: Area): FlowStatus {
		if (this.token !== area.token) {
			this.token = area.token
			this.y = area.top
			this.areaPlaced = 0
			this.source = -1
			this.box.open = false
		}

		while (this.index < this.count || this.inParagraph || this.table !== undefined) {
			let status: FlowStatus

			if (this.table !== undefined) status = this.runTable(area)
			else {
				const block = this.blocks[this.index] ?? -1
				const part = this.partAt(this.index)
				const kind = part?.kinds[block]

				if (kind === Item.Table && part !== undefined) {
					this.closeBox(area)
					this.furthest = Math.max(this.furthest, this.index)

					const table = new TableLayout(this, part, block, area)
					const floated = this.floatTable(area, table)

					if (floated === undefined) {
						// A table that flows with the text (one too tall to float starts over).
						this.table = table.floated ? new TableLayout(this, part, block, area) : table
						status = this.runTable(area)
					} else status = floated
				} else status = this.paragraph(area)
			}

			if (status !== FlowStatus.DONE) {
				if (status !== FlowStatus.SUSPEND) this.closeBox(area)

				if ((status === FlowStatus.OVERFLOW || status === FlowStatus.COLUMN_BREAK) && this.rollbackChain(area))
					return FlowStatus.OVERFLOW

				return status
			}
		}

		this.closeBox(area)

		return FlowStatus.DONE
	}

	// A floating table (w:tblpPr) in the body: laid out whole at its position, the text after it wraps around
	// it. Undefined when it flows with the text instead (not floating, in a cell or note, or taller than a
	// page); OVERFLOW when it starts on the next page.
	private floatTable(area: Area, table: TableLayout): FlowStatus | undefined {
		const position = table.position
		const exclusions = area.exclusions

		if (position === undefined || exclusions === null || area.inTable || area.region !== Region.BODY || !area.breaks) return undefined

		const page = area.page
		const frame = area.frame
		const mark = page.mark(new OutputMark())
		const footnotes = area.footnotes?.mark() ?? 0
		const anchor = position.vertAnchor
		const top =
			(anchor === "page" ? 0 : anchor === "margin" ? frame.marginTop : this.y) +
			(position.ySpec === undefined ? twips(position.y) : 0)
		const scratch = this.floatArea.copy(area)

		scratch.top = top
		scratch.bottom = Infinity
		scratch.breaks = false
		scratch.empty = true
		scratch.token = nextToken()
		table.floated = true

		const status = table.run(scratch, top)

		if (status === FlowStatus.SUSPEND) {
			page.rollback(mark)
			area.footnotes?.rollback(footnotes)

			return FlowStatus.SUSPEND
		}

		const height = table.y - top
		let y = top

		if (position.ySpec !== undefined && position.ySpec !== "inline") {
			const from = anchor === "page" ? 0 : anchor === "margin" ? frame.marginTop : this.y
			const to = anchor === "page" ? frame.height : anchor === "margin" ? frame.height - frame.marginBottom : area.bottom
			const align = position.ySpec

			y = align === "center" ? (from + to - height) / 2 : align === "bottom" || align === "outside" ? to - height : from
		}

		const left = table.left - twips(position.leftFromText ?? 0)
		const right = table.left + table.width + twips(position.rightFromText ?? 0)

		// w:tblOverlap never: below the floating objects and tables it would cover.
		if (table.overlap === "never") {
			for (let k = 0; k < MAX_OVERLAP_MOVES; k++) {
				const below = exclusions.overlapBottom(left, y, right, y + height)

				if (below === -Infinity) break

				y = below + twips(position.topFromText ?? 0)
			}
		}

		// Anchored to the text and not fitting below it: the table starts on the next page, or (taller than a
		// page) flows and splits like any other.
		if (anchor === "text" && y + height > area.bottom - (area.footnotes?.reserved ?? 0) + EPSILON) {
			page.rollback(mark)
			area.footnotes?.rollback(footnotes)

			return !area.empty || this.areaPlaced > 0 ? FlowStatus.OVERFLOW : undefined
		}

		if (y !== top) page.move(mark, null, 0, y - top)

		exclusions.add(left, y - twips(position.topFromText ?? 0), right, y + height + twips(position.bottomFromText ?? 0), Wrap.BOTH)
		area.empty = false
		this.areaPlaced++
		this.index++
		this.endChain()

		return FlowStatus.DONE
	}

	private runTable(area: Area): FlowStatus {
		const table = this.table

		if (table === undefined) return FlowStatus.DONE

		const before = area.page.text.lineCount
		const status = table.run(area, this.y)

		this.y = table.y

		// A row waiting for fonts is laid out again as if it had not begun (repeated header rows stay).
		if (status !== FlowStatus.SUSPEND && (area.page.text.lineCount > before || table.rowsPlaced > 0)) {
			area.empty = false
			this.areaPlaced++

			if (table.rowsPlaced > 0) this.endChain()
		}

		if (status === FlowStatus.DONE) {
			this.table = undefined
			this.index++
			this.after = 0
			this.contextual = false
			this.styleId = undefined
			this.hasPrevious = true
		}

		return status
	}

	// Saves the state at the start of a keep-with-next chain.
	private startChain(area: Area): void {
		const footnotes = area.footnotes

		this.chain = {
			index: this.index,
			y: this.y,
			after: this.after,
			contextual: this.contextual,
			style: this.styleId,
			hasPrevious: this.hasPrevious,
			empty: area.empty,
			mark: area.page.mark(),
			exclusions: area.exclusions?.count ?? 0,
			footnotes: footnotes === null ? 0 : footnotes.mark(),
			token: area.token
		}
	}

	// A block not kept with the next one placed content: the chain ends with it.
	private endChain(): void {
		this.chain = undefined
	}

	// The page closes: moved chains whose block is laid out are done, the others are a page further on.
	leavePage(): void {
		const moves = this.moves
		let kept = 0

		for (let k = 0; k < moves.length; k += 2) {
			const block = moves[k + 1] ?? 0

			if (block < this.index) continue

			moves[kept++] = (moves[k] ?? 0) + 1
			moves[kept++] = block
		}

		moves.length = kept
	}

	// Pages back to the earliest page a chain moved off whose block is still to come (0: none).
	get keepBack(): number {
		let back = 0

		for (let k = 0; k < this.moves.length; k += 2) back = Math.max(back, this.moves[k] ?? 0)

		return back
	}

	// Moves an unfinished chain that started in this area to the next one; false when there is none or it
	// started at the area's top (it cannot move anywhere better).
	private rollbackChain(area: Area): boolean {
		const chain = this.chain

		this.chain = undefined

		if (chain === undefined || chain.token !== area.token || chain.empty) return false

		this.moves.push(0, this.index)
		area.page.rollback(chain.mark)

		if (area.exclusions !== null) area.exclusions.count = Math.min(area.exclusions.count, chain.exclusions)

		area.footnotes?.rollback(chain.footnotes)
		this.index = chain.index
		this.y = chain.y
		this.after = chain.after
		this.contextual = chain.contextual
		this.styleId = chain.style
		this.hasPrevious = chain.hasPrevious
		this.inParagraph = false
		this.table = undefined
		this.box.open = false

		return true
	}

	private group(): { items: number[]; parts: StoryPart[]; formats: ParagraphFormat[] } {
		const content = this.engine.content
		const items = SCRATCH_ITEMS
		const parts = SCRATCH_PARTS
		const formats = SCRATCH_FORMATS

		items.length = 0
		parts.length = 0
		formats.length = 0

		for (let k = this.index; k < this.count; k++) {
			const item = this.blocks[k] ?? -1
			const part = this.partAt(k)

			if (part === undefined || part.kinds[item] !== Item.Paragraph) break

			const format = content.formatOf(part, item, this.cell)

			items.push(item)
			parts.push(part)
			formats.push(format)
			this.furthest = Math.max(this.furthest, k)

			if (!content.markHidden(part, item, format) || k + 1 >= this.count) break
		}

		return { items, parts, formats }
	}

	private paragraph(area: Area): FlowStatus {
		const engine = this.engine
		const built = this.built
		const page = area.page

		if (!this.inParagraph) {
			const framed = this.frame(area)

			if (framed !== undefined) return framed

			const { items, parts, formats } = this.group()

			engine.content.build(this.story, parts, items, formats, area.width, area.values, built)
			this.builtValues = area.values

			const props = built.format?.props

			if (props === undefined) {
				this.index += Math.max(1, items.length)

				return FlowStatus.DONE
			}

			if (this.suspendable && engine.styles.hasMissing) return FlowStatus.SUSPEND

			engine.layouter.prepare(engine.content.builder, this.prepareOptions(props), this.prepared)

			// Text boxes are built ahead (after the paragraph's own text is prepared: building reuses the builder).
			if (this.suspendable && (built.floats.length > 0 || built.boxedObjects.length > 0)) {
				this.scanBoxes(area)

				if (engine.styles.hasMissing) return FlowStatus.SUSPEND
			}
			this.inParagraph = true
			this.groupEnd = this.index + items.length
			this.cursor.reset()
			this.breakBeforeDone = false
			this.floatsPlaced = false
			this.source = -1
		}

		const props = built.format?.props

		if (props === undefined) return FlowStatus.DONE

		const cursor = this.cursor
		const first = cursor.rows === 0

		if (!first && built.pageFields.length > 0 && this.builtValues !== undefined && this.builtValues.page !== area.values.page)
			this.renumber(area)

		if (first && props.pageBreakBefore && area.breaks && !this.breakBeforeDone) {
			this.breakBeforeDone = true

			if (this.areaPlaced > 0 || !area.empty) return FlowStatus.PAGE_BREAK
		}

		const g = paragraphGeometry(this.geometry, props, area.left, area.width, engine.compat, engine.gridPitch)
		let top = this.y

		if (first) {
			// A chain starting at the area's top cannot move anywhere better: it is kept as a marker so the rest
			// of a run of keep-with-next paragraphs flows on instead of starting chains of its own.
			if (this.chain === undefined && props.keepNext && area.breaks) this.startChain(area)

			top += this.gap(area, props)
			top = this.openBox(area, props, top)

			if (!this.floatsPlaced) {
				this.floatsPlaced = true
				area.page.mark(this.floatMark)
				this.floatExclusions = area.exclusions?.count ?? 0
				this.placeFloats(area, top)
			}

			if (built.dropCap !== undefined) return this.dropCap(area, top)

			if (built.markerAlignment === "right" || built.markerAlignment === "end") g.firstLineIndent -= built.markerWidth
			else if (built.markerAlignment === "center") g.firstLineIndent -= built.markerWidth / 2
		}

		g.exclusions = area.exclusions !== null && area.exclusions.count > 0 && this.noRoom < 2 ? area.exclusions : null

		if (this.source < 0) {
			this.source = page.addSource(
				this.region,
				this.story.story,
				built.first,
				(area.inTable ? SourceFlag.IN_TABLE : 0) |
					(built.fieldFlags & 1 ? SourceFlag.PAGE_FIELDS : 0) |
					(built.fieldFlags & 2 ? SourceFlag.TOTAL_FIELDS : 0),
				built.text,
				this.region === Region.BODY && built.renderedBreaks.length > 0 ? built.renderedBreaks.slice() : undefined,
				built.firstPart,
				built.map
			)
		}

		this.prepared.source = this.source

		const footnotes = area.footnotes
		const force = this.areaPlaced === 0 && area.empty
		const start = page.text.lineCount
		const footnoteMark = footnotes === null ? 0 : footnotes.mark()
		const rowsBefore = cursor.rows

		this.copyCursor(this.saved, cursor)

		const bottom = area.bottom - (footnotes?.reserved ?? 0)
		let status: LayoutStatus

		if (first && force && area.breaks && engine.compat.suppressTopSpacing && g.lineRule !== LineRule.AUTO) {
			// w:suppressTopSpacing: the page's first line takes its natural height.
			const rule = g.lineRule
			const spacing = g.lineSpacing

			g.lineRule = LineRule.AUTO
			g.lineSpacing = 1
			status = engine.layouter.layout(this.prepared, g, page.text, cursor, top, top, true)
			g.lineRule = rule
			g.lineSpacing = spacing

			if (status === LayoutStatus.OVERFLOW)
				status = engine.layouter.layout(this.prepared, g, page.text, cursor, cursor.y, bottom, false)
		} else {
			status = engine.layouter.layout(this.prepared, g, page.text, cursor, top, bottom, force)
		}

		let placed = page.text.lineCount - start

		if (status === LayoutStatus.NO_ROOM) {
			this.noRoom++
			this.copyCursor(cursor, this.saved)

			if (first) this.withdraw(area)

			return FlowStatus.OVERFLOW
		}

		// Notes referenced on the placed lines must fit at the page bottom with them.
		if (footnotes !== null && built.noteOffsets.length > 0 && placed > 0) {
			const keep = this.fitNotes(area, footnotes, start, placed, force)

			if (keep < placed) {
				footnotes.rollback(footnoteMark)
				placed = this.cut(area, g, start, keep, top, force)
				this.fitNotes(area, footnotes, start, placed, force)
				status = LayoutStatus.OVERFLOW
			}
		}

		// Above footnotes, a paragraph's space after must fit too: its last line moves on when it does not.
		if (
			status === LayoutStatus.DONE &&
			area.breaks &&
			!force &&
			placed > 0 &&
			footnotes !== null &&
			footnotes.reserved > 0 &&
			engine.tuning.noteSpaceAfter === "fit" &&
			!footnotes.split
		) {
			if (cursor.y + this.spaceAfter(props) > area.bottom - footnotes.reserved + EPSILON) {
				footnotes.rollback(footnoteMark)
				placed = this.cut(area, g, start, placed - 1, top, force)

				if (built.noteOffsets.length > 0) this.fitNotes(area, footnotes, start, placed, force)

				status = LayoutStatus.OVERFLOW
			}
		}

		if (status === LayoutStatus.OVERFLOW && area.breaks) {
			let keep = placed

			if (props.keepLines && rowsBefore === 0 && !force) keep = 0
			else if (props.widowControl) {
				const remaining = engine.remainingRows(this.prepared, g, cursor)

				if (remaining === 1 && keep >= 1 && !(force && keep === 1)) keep--
				if (rowsBefore === 0 && keep === 1 && !force) keep = 0
			}

			if (keep < placed) {
				if (footnotes !== null) footnotes.rollback(footnoteMark)

				if (keep === 0) {
					page.text.truncate(start)
					this.copyCursor(cursor, this.saved)
					placed = 0
				} else {
					placed = this.cut(area, g, start, keep, top, force)

					if (footnotes !== null && built.noteOffsets.length > 0) this.fitNotes(area, footnotes, start, placed, force)
				}
			}
		}

		if (placed > 0) {
			this.noRoom = 0
			area.empty = false
			this.areaPlaced++
			this.y = cursor.y
			this.afterPlaced(area, start, placed, props)

			// The paragraph after a keep-with-next run has its first lines here (widow control keeps two): the
			// run stays on this page.
			if (!props.keepNext) this.endChain()
		}

		if (status === LayoutStatus.OVERFLOW) {
			// Nothing fitted: the paragraph starts over in the next area.
			if (placed === 0 && first) this.withdraw(area)

			return FlowStatus.OVERFLOW
		}

		if (status === LayoutStatus.PAGE_BREAK || status === LayoutStatus.COLUMN_BREAK) {
			this.source = -1

			if (!area.breaks) return FlowStatus.DONE

			if (cursor.item >= this.prepared.count && !cursor.pendingMark) this.finishParagraph(props)

			return status === LayoutStatus.PAGE_BREAK ? FlowStatus.PAGE_BREAK : FlowStatus.COLUMN_BREAK
		}

		this.finishParagraph(props)

		return FlowStatus.DONE
	}

	// Takes back what a paragraph that placed no line here put into the area (anchored objects, text boxes,
	// its source), so it starts over cleanly in the next one.
	private withdraw(area: Area): void {
		this.inParagraph = false
		this.box.open = false

		if (this.floatsPlaced) {
			area.page.rollback(this.floatMark)

			if (area.exclusions !== null) area.exclusions.count = Math.min(area.exclusions.count, this.floatExclusions)
		}

		this.floatsPlaced = false
		this.source = -1
	}

	// Auto spacing applies between neighbouring paragraphs only: none after the last.
	private spaceAfter(props: ResolvedParagraphProps): number {
		return props.spacing.afterAutospacing && this.groupEnd >= this.count ? 0 : spaceAfter(props, this.engine.autoSpacing)
	}

	private finishParagraph(props: ResolvedParagraphProps): void {
		this.inParagraph = false
		this.index = this.groupEnd
		this.after = this.spaceAfter(props)
		this.contextual = props.contextualSpacing
		this.styleId = props.styleId
		this.hasPrevious = true

		if (!props.keepNext) this.endChain()
	}

	private prepareOptions(props: ResolvedParagraphProps): PrepareOptions {
		const language = this.built.format?.runBase.lang?.eastAsia
		const lists = props.kinsoku ? kinsokuLists(language, false) : null

		return {
			rtl: props.bidi,
			source: 0,
			kinsoku: lists === null ? null : kinsokuOf(lists),
			hangingPunctuation: props.overflowPunct && lists !== null,
			wordWrap: props.wordWrap
		}
	}

	// Space between the previous block and this paragraph.
	private gap(area: Area, props: ResolvedParagraphProps): number {
		const before = spaceBefore(props, this.engine.autoSpacing)
		const sameStyle = this.styleId === props.styleId
		const after = this.contextual && sameStyle ? 0 : this.after
		const own = props.contextualSpacing && sameStyle && this.hasPrevious ? 0 : before

		const tuning = this.engine.tuning
		const collapse = tuning.paragraphSpacing === "max" || (tuning.paragraphSpacing === "auto" && this.engine.compat.htmlSpacing)
		const join = (a: number, b: number): number => (collapse ? Math.max(a, b) : a + b)

		if (this.areaPlaced === 0) {
			if (area.breaks && area.empty) {
				const compat = this.engine.compat
				const afterBreaks =
					tuning.dropSpaceAtPageTop === "breaks" ||
					(tuning.dropSpaceAtPageTop === "compat" && (compat.compatibilityMode >= 15 || compat.suppressSpaceAfterBreak))
				const drop =
					tuning.dropSpaceAtPageTop === "never"
						? false
						: afterBreaks
							? area.start !== AreaStart.DOCUMENT
							: area.start === AreaStart.OVERFLOW && !this.breakBeforeDone

				if (drop) return 0
			}

			return this.hasPrevious && !area.empty ? join(after, own) : own
		}

		return join(after, own)
	}

	// Reserves notes referenced on lines [start, start + placed); returns how many lines keep their notes
	// in the area.
	private fitNotes(area: Area, footnotes: FootnoteSink, start: number, placed: number, force: boolean): number {
		const text = area.page.text
		const built = this.built
		const limit = area.bottom
		let ref = 0

		for (let k = 0; k < placed; k++) {
			const line = start + k
			const textStart = text.lineInts[line * LINE_INTS + LineInt.TEXT_START] ?? 0
			const textEnd = text.lineInts[line * LINE_INTS + LineInt.TEXT_END] ?? 0
			const bottom = (text.lineFloats[line * LINE_FLOATS] ?? 0) + (text.lineFloats[line * LINE_FLOATS + 1] ?? 0)
			const lineForced = force && k === 0
			const mark = footnotes.mark()

			while (ref < built.noteOffsets.length && (built.noteOffsets[ref] ?? 0) < textStart) ref++

			for (; ref < built.noteOffsets.length && (built.noteOffsets[ref] ?? 0) < Math.max(textEnd, textStart + 1); ref++) {
				const story = built.noteStories[ref] ?? -1

				if (built.noteEndnote[ref] === 1) {
					this.engine.addEndnote(story)

					continue
				}

				if (!footnotes.add(story, bottom, limit, lineForced)) {
					footnotes.rollback(mark)

					return k
				}
			}

			if (!lineForced && bottom > limit - footnotes.reserved + EPSILON) {
				footnotes.rollback(mark)

				return k
			}
		}

		return placed
	}

	// Keeps `keep` of the lines laid out from `start`, laying the paragraph out again up to there so the
	// cursor resumes after them.
	private cut(area: Area, g: Geometry, start: number, keep: number, top: number, force: boolean): number {
		const page = area.page
		const text = page.text

		if (keep <= 0) {
			text.truncate(start)
			this.copyCursor(this.cursor, this.saved)

			return 0
		}

		const limit = (text.lineFloats[(start + keep) * LINE_FLOATS] ?? 0) + 0.01

		text.truncate(start)
		this.copyCursor(this.cursor, this.saved)
		this.engine.layouter.layout(this.prepared, g, text, this.cursor, top, limit, force)

		return text.lineCount - start
	}

	private copyCursor(to: LayoutCursor, from: LayoutCursor): void {
		to.item = from.item
		to.offset = from.offset
		to.rows = from.rows
		to.pendingMark = from.pendingMark
		to.y = from.y
	}

	private afterPlaced(area: Area, start: number, placed: number, _props: ResolvedParagraphProps): void {
		const box = this.box
		const text = area.page.text

		if (box.open)
			box.bottom = Math.max(
				box.bottom,
				(text.lineFloats[(start + placed - 1) * LINE_FLOATS] ?? 0) + (text.lineFloats[(start + placed - 1) * LINE_FLOATS + 1] ?? 0)
			)

		if (this.built.boxedObjects.length > 0) this.inlineBoxes(area, start, placed)
	}

	// Text boxes inside inline objects on lines [start, start + placed).
	private inlineBoxes(area: Area, start: number, placed: number): void {
		const engine = this.engine
		const objects = engine.objects
		const text = area.page.text

		for (let line = start; line < start + placed; line++) {
			const first = text.lineInts[line * LINE_INTS + LineInt.FIRST_FRAGMENT] ?? 0
			const count = text.lineInts[line * LINE_INTS + LineInt.FRAGMENT_COUNT] ?? 0
			const baseline = text.lineFloats[line * LINE_FLOATS + 2] ?? 0

			for (let k = first; k < first + count; k++) {
				if (((text.fragmentInts[k * FRAGMENT_INTS + 3] ?? 0) & FRAGMENT_KIND_MASK) !== FragmentKind.OBJECT) continue

				const id = text.fragmentInts[k * FRAGMENT_INTS + 4] ?? -1

				if (!this.built.boxedObjects.includes(id)) continue

				const x = text.fragmentFloats[k * FRAGMENT_FLOATS] ?? 0

				this.textboxes(area, id, x, baseline - (objects.ascents[id] ?? 0))
			}
		}
	}

	private textboxes(area: Area, id: number, x: number, y: number): void {
		const engine = this.engine
		const objects = engine.objects
		const stories = objects.textboxes[id]
		const insets = objects.insets[id] ?? [0, 0, 0, 0]

		if (stories === undefined) return

		const width = (objects.widths[id] ?? 0) - (insets[0] ?? 0) - (insets[2] ?? 0)
		let top = y + (insets[1] ?? 0)

		for (const index of stories) {
			const story = engine.model.stories[index]

			if (story === undefined) continue

			const part = engine.model.parts[story.part]

			if (part === undefined) continue

			top += engine.layoutBox(
				{ part, story: index, counters: this.story.counters, noteMark: "", breaksAsLines: true, recount: this.story.recount },
				Region.TEXTBOX,
				area.page,
				area.values,
				area.frame,
				x + (insets[0] ?? 0),
				top,
				width
			)
		}
	}

	// Builds text box contents ahead so their fonts are prepared with the paragraph's.
	private scanBoxes(area: Area): void {
		const engine = this.engine
		const objects = engine.objects
		const built = this.built

		for (const list of [built.floats, built.boxedObjects]) {
			for (const id of list) {
				for (const index of objects.textboxes[id] ?? []) {
					const story = engine.model.stories[index]
					const part = story === undefined ? undefined : engine.model.parts[story.part]

					if (story === undefined || part === undefined) continue

					engine.scan(
						{ part, story: index, counters: undefined, noteMark: "", breaksAsLines: true, recount: false },
						story.item,
						area.values
					)
				}
			}
		}
	}

	// Positions the paragraph's anchored objects and makes those that wrap text into obstacles.
	private placeFloats(area: Area, paragraphTop: number): void {
		const engine = this.engine
		const objects = engine.objects

		for (const id of this.built.floats) {
			const spec = objects.floats[id]

			if (spec === undefined) continue

			const width = objects.widths[id] ?? 0
			const height = objects.heights[id] ?? 0
			const x = horizontalPosition(spec, area, width)

			if (byParity(spec.horizontalAlign) || byParity(spec.verticalAlign) || marginByParity(spec.horizontalFrom))
				this.engine.parityReads++

			const layer = spec.behindDoc ? Layer.BEHIND : spec.wrap === "none" ? Layer.FRONT : Layer.TEXT
			const exclusions = area.exclusions
			const wraps = exclusions !== null && spec.wrap !== "none" && !spec.behindDoc && spec.wrap !== "unknown"
			let y = verticalPosition(spec, area, paragraphTop, height)

			// An object that may not overlap others moves down below the anchored drawings it would cover.
			if (wraps && !spec.allowOverlap) y = clearOfDrawings(area.page, objects, x, y, width, height)

			area.page.addObject(id, layer, -1, x, y, width, height)

			if (wraps) {
				const top = y - spec.distTop
				const bottom = y + height + spec.distBottom

				if (spec.wrap === "topAndBottom") exclusions.add(-1e7, top, 1e7, bottom, Wrap.BOTH)
				else if (spec.polygon !== undefined) wrapAround(exclusions, spec, x, y, width, height)
				else exclusions.add(x - spec.distLeft, top, x + width + spec.distRight, bottom, wrapSide(spec.wrapText))
			}

			if (objects.textboxes[id] !== undefined) this.textboxes(area, id, x, y)
		}
	}

	// Paragraphs in a frame (w:framePr other than drop caps; neighbours with the same frame share it): laid
	// out in a box at the frame's position, the text after them wrapping around it. Undefined when the next
	// paragraph has no frame.
	private frame(area: Area): FlowStatus | undefined {
		const engine = this.engine
		const part = this.partAt(this.index)
		const content = engine.content

		if (part === undefined || this.inFrame) return undefined

		const first = this.blocks[this.index] ?? -1
		// Word does not float frames inside table cells.
		const frame =
			part.kinds[first] === Item.Paragraph && !area.inTable ? content.formatOf(part, first, this.cell).props.frame : undefined

		if (frame === undefined || (frame.dropCap !== undefined && frame.dropCap !== "none")) return undefined

		let end = this.index + 1

		while (end < this.count) {
			const item = this.blocks[end] ?? -1
			const other = this.partAt(end)

			if (
				other === undefined ||
				other.kinds[item] !== Item.Paragraph ||
				!sameFrame(content.formatOf(other, item, this.cell).props.frame, frame)
			)
				break

			end++
		}

		this.furthest = Math.max(this.furthest, end - 1)

		// The fonts the frame's text needs, loaded before it is laid out.
		if (this.suspendable) {
			for (let k = this.index; k < end; k++) {
				const item = this.blocks[k] ?? -1
				const other = this.partAt(k) ?? part

				content.build(this.story, [other], [item], [content.formatOf(other, item, this.cell)], area.width, area.values, this.built)
			}

			if (engine.styles.hasMissing) return FlowStatus.SUSPEND
		}

		const page = area.page
		const pageFrame = area.frame
		const mark = page.mark(new OutputMark())
		const inner = new BlockFlow(engine, { ...this.story, breaksAsLines: true }, this.region, false).copyBlocks(this, this.index, end)

		inner.inFrame = true
		const box = this.floatArea.copy(area)
		const fixed = frame.width !== undefined && frame.width > 0
		// Without anchors a frame stays where its paragraph is: Word anchors it to the text both ways
		// (MS-OI29500 §17.3.1.11 e, g).
		const hAnchor = frame.hAnchor ?? "text"
		const vAnchor = frame.vAnchor ?? "text"
		const hStart = hAnchor === "page" ? 0 : hAnchor === "margin" ? pageFrame.marginLeft : area.left
		const hSize =
			hAnchor === "page"
				? pageFrame.width
				: hAnchor === "margin"
					? pageFrame.width - pageFrame.marginLeft - pageFrame.marginRight
					: area.width
		const vStart = vAnchor === "page" ? 0 : vAnchor === "margin" ? pageFrame.marginTop : this.y
		const vSize =
			vAnchor === "page"
				? pageFrame.height
				: vAnchor === "margin"
					? pageFrame.height - pageFrame.marginTop - pageFrame.marginBottom
					: 0

		box.left = hStart
		box.width = fixed ? twips(frame.width ?? 0) : area.width
		box.top = vStart
		box.bottom = Infinity
		box.breaks = false
		box.empty = true
		box.inTable = false
		box.exclusions = null
		box.footnotes = null
		box.token = nextToken()
		inner.run(box)

		let width = box.width

		// A frame without a width takes its widest line.
		if (!fixed) {
			const text = page.text
			let right = -Infinity

			for (let line = mark.lines; line < text.lineCount; line++)
				right = Math.max(right, (text.lineFloats[line * LINE_FLOATS + 3] ?? 0) + (text.lineFloats[line * LINE_FLOATS + 4] ?? 0))

			width = right === -Infinity ? 0 : right - hStart
		}

		const natural = inner.contentBottom() - vStart
		// A height without a rule is a minimum; no height is the content's whatever the rule (MS-OI29500
		// §17.3.1.11 d, f).
		const given = frame.height ?? 0
		const rule = given === 0 ? "auto" : (frame.hRule ?? "atLeast")
		const height = rule === "exact" ? twips(given) : rule === "atLeast" ? Math.max(natural, twips(given)) : natural

		if (byParity(frame.xAlign) || (vAnchor !== "text" && byParity(frame.yAlign))) this.engine.parityReads++

		const x = alignWithin(frame.xAlign, hStart, hSize, width, pageFrame.odd) ?? hStart + twips(frame.x ?? 0)
		const y =
			(vAnchor === "text" ? undefined : alignWithin(frame.yAlign, vStart, vSize, height, pageFrame.odd)) ??
			vStart + twips(frame.y ?? 0)

		// Anchored to the text and not fitting below it: the frame starts on the next page.
		if (
			vAnchor === "text" &&
			y + height > area.bottom - (area.footnotes?.reserved ?? 0) + EPSILON &&
			(!area.empty || this.areaPlaced > 0)
		) {
			page.rollback(mark)

			return FlowStatus.OVERFLOW
		}

		page.move(mark, null, x - hStart, y - vStart)

		if (rule === "exact") {
			const end = page.mark(new OutputMark())

			page.addClip(x, y, Math.max(width, 1), height, mark.lines, end.lines, mark.objects, end.objects)
		}

		const exclusions = area.exclusions
		const wrap = frame.wrap ?? "around"

		if (exclusions !== null && wrap !== "none") {
			const hSpace = twips(frame.hSpace ?? 0)
			const vSpace = twips(frame.vSpace ?? 0)

			if (wrap === "notBeside") exclusions.add(-1e7, y - vSpace, 1e7, y + height + vSpace, Wrap.BOTH)
			else exclusions.add(x - hSpace, y - vSpace, x + width + hSpace, y + height + vSpace, Wrap.BOTH)
		}

		area.empty = false
		this.areaPlaced++
		this.index = end
		this.endChain()

		return FlowStatus.DONE
	}

	// A drop cap: the paragraph's text in a frame at the left of the next paragraph's first lines.
	private dropCap(area: Area, top: number): FlowStatus {
		const engine = this.engine
		const page = area.page
		const g = this.geometry
		const props = this.built.format?.props
		const frame = props?.frame
		const scratch = new TextLayout()
		const cursor = new LayoutCursor()

		g.width = 10000
		g.startIndent = 0
		g.endIndent = 0
		g.firstLineIndent = 0
		g.exclusions = null
		engine.layouter.layout(this.prepared, g, scratch, cursor, 0, Infinity, true)

		let width = 0

		for (let line = 0; line < scratch.lineCount; line++) width = Math.max(width, scratch.lineFloats[line * LINE_FLOATS + 4] ?? 0)

		const height = cursor.y
		const space = (frame?.hSpace ?? 0) / 15
		const x = this.built.dropCap === "margin" ? area.left - width - space : area.left

		g.left = x
		g.width = width + 1
		cursor.reset()
		this.prepared.source = page.addSource(
			this.region,
			this.story.story,
			this.built.first,
			0,
			this.built.text,
			undefined,
			this.built.firstPart,
			this.built.map
		)
		engine.layouter.layout(this.prepared, g, page.text, cursor, top, Infinity, true)

		if (this.built.dropCap === "drop" && area.exclusions !== null)
			area.exclusions.add(x, top, x + width + space, top + height, Wrap.RIGHT)

		area.empty = false
		this.inParagraph = false
		this.index = this.groupEnd

		return FlowStatus.DONE
	}

	// Starts or continues a border box; returns the paragraph's top moved below its top border.
	private openBox(area: Area, props: ResolvedParagraphProps, top: number): number {
		const box = this.box
		const borders = props.borders
		const shading = this.engine.decorations.shading(props.shading)

		if (
			box.open &&
			box.borders === borders &&
			box.shading === shading &&
			box.left === area.left + paragraphLeft(props) &&
			borders !== undefined
		) {
			box.joins.push(box.bottom)

			return top
		}

		if (box.open && box.borders === borders && box.shading === shading && borders === undefined && shading !== 0) {
			box.joins.push(box.bottom)

			return top
		}

		const before = this.y

		this.closeBox(area)
		// The closed box's bottom border pushes this paragraph down.
		top += this.y - before

		if (borders === undefined && shading === 0) return top

		box.open = true
		box.borders = borders
		box.shading = shading
		box.left = area.left + paragraphLeft(props)
		box.right = area.left + area.width - paragraphRight(props)
		box.joins.length = 0

		const topBorder = borders?.top

		box.top = top
		box.bottom = top

		return topBorder === undefined ? top : top + borderWidth(topBorder) + points(topBorder.space)
	}

	// Draws the open border box.
	closeBox(area: Area): void {
		const box = this.box

		if (!box.open) return

		box.open = false

		const page = area.page
		const borders = box.borders
		const bottomBorder = borders?.bottom
		const bottom = bottomBorder === undefined ? box.bottom : box.bottom + points(bottomBorder.space)
		const decorations = this.engine.decorations
		const left = box.left - (borders?.start === undefined ? 0 : points(borders.start.space))
		const right = box.right + (borders?.end === undefined ? 0 : points(borders.end.space))

		if (box.shading !== 0)
			page.addGraphic(PageGraphic.SHADING, box.shading, Layer.BEHIND, left, box.top, right - left, bottom - box.top)

		if (borders === undefined) return

		const line = (border: ResolvedBorder | undefined, x: number, y: number, horizontal: boolean, length: number): void => {
			if (border === undefined) return

			const width = borderWidth(border)

			page.addGraphic(
				PageGraphic.BORDER,
				decorations.border(border),
				Layer.TEXT,
				x,
				y,
				horizontal ? length : width,
				horizontal ? width : length
			)
		}

		line(borders.top, left, box.top, true, right - left)
		line(borders.bottom, left, bottom, true, right - left)
		line(borders.start, left - (borders.start === undefined ? 0 : borderWidth(borders.start)), box.top, false, bottom - box.top)
		line(borders.end, right, box.top, false, bottom - box.top)

		for (const join of box.joins) line(borders.between, left, join, true, right - left)

		this.y = Math.max(this.y, bottomBorder === undefined ? this.y : bottom + borderWidth(bottomBorder))
	}
}

function points(value: number): number {
	return (value * 4) / 3
}

function paragraphLeft(props: ResolvedParagraphProps): number {
	return (props.bidi ? props.indent.end : props.indent.start) / 15
}

function paragraphRight(props: ResolvedParagraphProps): number {
	return (props.bidi ? props.indent.start : props.indent.end) / 15
}

// The top at or below `y` where a box clears the anchored, wrapping drawings already on the page.
function clearOfDrawings(page: PageOutput, objects: ObjectTable, x: number, y: number, width: number, height: number): number {
	for (let move = 0; move < MAX_OVERLAP_MOVES; move++) {
		let below = -Infinity

		for (let o = 0; o < page.objectCount; o++) {
			const id = page.objectInts[o * OBJECT_INTS] ?? -1
			const spec = objects.floats[id]

			if (spec === undefined || objects.kinds[id] !== ObjectKind.DRAWING || spec.wrap === "none" || spec.behindDoc) continue

			const f = o * OBJECT_FLOATS
			const left = page.objectFloats[f] ?? 0
			const top = page.objectFloats[f + 1] ?? 0
			const bottom = top + (page.objectFloats[f + 3] ?? 0)

			if (left < x + width && left + (page.objectFloats[f + 2] ?? 0) > x && top < y + height && bottom > y)
				below = Math.max(below, bottom)
		}

		if (below === -Infinity) break

		y = below
	}

	return y
}

// Moves of an object that may not overlap others, at most (each clears one obstacle).
const MAX_OVERLAP_MOVES = 32
// Height of the bands a wrap polygon is cut into (px), and their most.
const POLYGON_BAND = 4
const MAX_POLYGON_BANDS = 128

// Tight and through wrapping: the polygon's horizontal extent per band, widened by the side distances.
function wrapAround(exclusions: Exclusions, spec: FloatSpec, x: number, y: number, width: number, height: number): void {
	const polygon = spec.polygon ?? []
	const bands = Math.max(1, Math.min(MAX_POLYGON_BANDS, Math.ceil(height / POLYGON_BAND)))
	const side = wrapSide(spec.wrapText)
	const count = polygon.length / 2

	for (let band = 0; band < bands; band++) {
		const top = band / bands
		const bottom = (band + 1) / bands
		let min = Infinity
		let max = -Infinity

		for (let k = 0; k < count; k++) {
			const x0 = polygon[2 * k] ?? 0
			const y0 = polygon[2 * k + 1] ?? 0
			const x1 = polygon[(2 * k + 2) % (2 * count)] ?? 0
			const y1 = polygon[(2 * k + 3) % (2 * count)] ?? 0
			const low = Math.max(top, Math.min(y0, y1))
			const high = Math.min(bottom, Math.max(y0, y1))

			if (low > high) continue

			// The edge's x where it enters and leaves the band (both ends when it runs along it).
			const a = y1 === y0 ? x0 : x0 + ((x1 - x0) * (low - y0)) / (y1 - y0)
			const b = y1 === y0 ? x1 : x0 + ((x1 - x0) * (high - y0)) / (y1 - y0)

			min = Math.min(min, a, b)
			max = Math.max(max, a, b)
		}

		if (min > max) continue

		exclusions.add(x + min * width - spec.distLeft, y + top * height, x + max * width + spec.distRight, y + bottom * height, side)
	}
}

function wrapSide(wrapText: string): Wrap {
	switch (wrapText) {
		case "left":
			return Wrap.LEFT
		case "right":
			return Wrap.RIGHT
		case "largest":
			return Wrap.LARGEST
		default:
			return Wrap.BOTH
	}
}

// Whether alignment `align` places by the page's parity (inside and outside).
export function byParity(align: string | undefined): boolean {
	return align === "inside" || align === "outside"
}

function marginByParity(from: string | undefined): boolean {
	return from === "insideMargin" || from === "outsideMargin"
}

function alignWithin(align: string | undefined, start: number, size: number, extent: number, odd: boolean): number | undefined {
	switch (align) {
		case "left":
		case "top":
			return start
		case "center":
			return start + (size - extent) / 2
		case "right":
		case "bottom":
			return start + size - extent
		case "inside":
			return odd ? start : start + size - extent
		case "outside":
			return odd ? start + size - extent : start
		default:
			return undefined
	}
}

// Whether two paragraphs' frames are the same frame (§17.3.1.11: identical frame properties).
function sameFrame(a: FrameProps | undefined, b: FrameProps): boolean {
	return (
		a !== undefined &&
		(a === b ||
			(a.dropCap === b.dropCap &&
				a.width === b.width &&
				a.height === b.height &&
				a.hRule === b.hRule &&
				a.x === b.x &&
				a.y === b.y &&
				a.xAlign === b.xAlign &&
				a.yAlign === b.yAlign &&
				a.hAnchor === b.hAnchor &&
				a.vAnchor === b.vAnchor &&
				a.wrap === b.wrap &&
				a.hSpace === b.hSpace &&
				a.vSpace === b.vSpace))
	)
}

function horizontalPosition(spec: FloatSpec, area: Area, width: number): number {
	const frame = area.frame
	let start: number
	let size: number

	switch (spec.horizontalFrom) {
		case "page":
			start = 0
			size = frame.width
			break
		case "margin":
			start = frame.marginLeft
			size = frame.width - frame.marginLeft - frame.marginRight
			break
		case "leftMargin":
			start = 0
			size = frame.marginLeft
			break
		case "rightMargin":
			start = frame.width - frame.marginRight
			size = frame.marginRight
			break
		case "insideMargin":
			start = frame.odd ? 0 : frame.width - frame.marginRight
			size = frame.odd ? frame.marginLeft : frame.marginRight
			break
		case "outsideMargin":
			start = frame.odd ? frame.width - frame.marginRight : 0
			size = frame.odd ? frame.marginRight : frame.marginLeft
			break
		default:
			start = area.left
			size = area.width
	}

	if (area.inTable && spec.layoutInCell && (spec.horizontalFrom === "page" || spec.horizontalFrom === "margin")) {
		start = area.left
		size = area.width
	}

	return alignWithin(spec.horizontalAlign, start, size, width, frame.odd) ?? start + spec.horizontalOffset
}

function verticalPosition(spec: FloatSpec, area: Area, paragraphTop: number, height: number): number {
	const frame = area.frame
	let start: number
	let size: number

	switch (spec.verticalFrom) {
		case "page":
			start = 0
			size = frame.height
			break
		case "margin":
			start = frame.marginTop
			size = frame.height - frame.marginTop - frame.marginBottom
			break
		case "topMargin":
			start = 0
			size = frame.marginTop
			break
		case "bottomMargin":
			start = frame.height - frame.marginBottom
			size = frame.marginBottom
			break
		case "insideMargin":
		case "outsideMargin":
			start = 0
			size = frame.marginTop
			break
		default:
			start = paragraphTop
			size = 0
	}

	if (area.inTable && spec.layoutInCell && (spec.verticalFrom === "page" || spec.verticalFrom === "margin")) {
		start = paragraphTop
		size = 0
	}

	return alignWithin(spec.verticalAlign, start, size, height, frame.odd) ?? start + spec.verticalOffset
}

export type { NumberingCounters }
