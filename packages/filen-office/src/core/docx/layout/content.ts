import { BreakKind, ParagraphBuilder, TabAlignment, TabLeader, type StyleTable } from "../../text/index"
import type { DocxModel } from "../model/document"
import { BreakClear, BreakType, FieldCharType, Item, RevisionView } from "../model/kinds"
import { Ns } from "../model/namespaces"
import type { StoryPart } from "../model/part"
import type { ListMarker, NumberingCounters } from "../numbering/counter"
import { EMPTY_RUN } from "../props/parse"
import type { PropsStore } from "../props/store"
import type { ResolvedRunProps } from "../props/types"
import type { CellFormat, ParagraphFormat, PropertyResolver } from "../styles/resolve"
import { FieldKind, fieldInfo, numberText } from "./fields"
import { objectAt, ObjectKind, type ObjectSizer, type ObjectTable } from "./objects"
import { TrackedLook, revisionLook, type RunStyles } from "./styles"
import { halfPoints, twips } from "./units"

// Builds the text a paragraph lays out from the model: runs resolved to styles, fields shown as their
// results (page fields computed), hidden text and tracked changes per view, list markers, note references,
// symbols, positional tabs and objects.

// Field values for the page being laid out.
export interface PageValues {
	page: number
	// The section's w:pgNumType format.
	pageFormat: string | undefined
	numPages: number
	sectionPages: number
}

// Numbers of note references, assigned in layout order.
export interface NoteNumbers {
	// The mark text of a reference (empty for w:customMarkFollows).
	reference(endnote: boolean, part: number, item: number, story: number): string
	// A paragraph's references are about to be numbered.
	beginParagraph(): void
}

// The story a paragraph belongs to, and its own counters.
export interface StoryContext {
	readonly part: StoryPart
	readonly story: number
	readonly counters: NumberingCounters | undefined
	// The mark text of the note being laid out (w:footnoteRef / w:endnoteRef).
	readonly noteMark: string
	// Page and column breaks do not apply (headers, footers, notes, cells, text boxes).
	readonly breaksAsLines: boolean
}

export const FieldFlag = { PAGE: 1, TOTALS: 2 } as const

export class BuiltParagraph {
	first = -1
	// The part of the first paragraph.
	firstPart = -1
	last = -1
	format: ParagraphFormat | undefined = undefined
	markStyle = 0
	text = ""
	// The list marker's width and lvlJc ("left", "center", "right"), 0 / "" without a marker.
	markerWidth = 0
	markerAlignment = ""
	// Anchored objects (ids) in text order.
	readonly floats: number[] = []
	// Note references: text offset, story index and 1 for endnotes.
	readonly noteOffsets: number[] = []
	readonly noteStories: number[] = []
	readonly noteEndnote: number[] = []
	// Inline objects with text boxes inside.
	readonly boxedObjects: number[] = []
	// w:lastRenderedPageBreak positions (text offsets).
	readonly renderedBreaks: number[] = []
	fieldFlags = 0
	// PAGE field results: start and end text offsets.
	readonly pageFields: number[] = []
	// The paragraph has only a mark (no text, tabs or objects).
	empty = true
	dropCap: "drop" | "margin" | undefined = undefined
	// Where each item's content landed in `text` (SOURCE_MAP_INTS per item), when the layout records it.
	map: Int32Array | undefined = undefined

	reset(): void {
		this.first = -1
		this.firstPart = -1
		this.last = -1
		this.format = undefined
		this.markStyle = 0
		this.text = ""
		this.markerWidth = 0
		this.markerAlignment = ""
		this.floats.length = 0
		this.noteOffsets.length = 0
		this.noteStories.length = 0
		this.noteEndnote.length = 0
		this.boxedObjects.length = 0
		this.renderedBreaks.length = 0
		this.fieldFlags = 0
		this.pageFields.length = 0
		this.empty = true
		this.dropCap = undefined
		this.map = undefined
	}
}

const FIELD_CODE = 0
const FIELD_RESULT = 1
// A computed field: its stored result is replaced.
const FIELD_REPLACED = 2

const LEADERS: Readonly<Record<string, TabLeader>> = {
	dot: TabLeader.DOT,
	hyphen: TabLeader.HYPHEN,
	underscore: TabLeader.UNDERSCORE,
	heavy: TabLeader.HEAVY,
	middleDot: TabLeader.MIDDLE_DOT
}

export interface ContentServices {
	readonly model: DocxModel
	readonly resolver: PropertyResolver
	readonly store: PropsStore
	readonly styles: RunStyles
	readonly objects: ObjectTable
	readonly sizer: ObjectSizer | undefined
	readonly notes: NoteNumbers
	// Line spacing multiples stretch inline objects too.
	readonly stretchObjects: boolean
	readonly view: RevisionView
	readonly showHidden: boolean
	// List markers already counted, by paragraph: a paragraph laid out again must not count twice.
	readonly markers: Map<number, ListMarker | null>
	// Record where each item's content lands in the paragraph text (editing maps positions through it).
	readonly recordMaps: boolean
	// The key list markers are counted under for paragraph `item` of part `part`.
	readonly memoKey: (part: number, item: number) => number
	// The index of a tracked change's author (-1 for none), numbered as reviewers are.
	readonly authors: (name: string | undefined) => number
}

// Decoded text of text items by part, while editing: a relayout lays the paragraphs around an edit out again
// keystroke after keystroke, and their text is decoded once. Bounded by characters (cleared when full); a part
// that goes (an edited block's earlier overlay) takes its entries with it.
const TEXT_CACHE_CHARS = 1 << 18
let textCache = new WeakMap<StoryPart, Map<number, string>>()
let textCacheChars = 0

function cachedText(part: StoryPart, item: number): string {
	let texts = textCache.get(part)
	let text = texts?.get(item)

	if (text !== undefined) return text

	text = part.text(item)

	if (textCacheChars + text.length > TEXT_CACHE_CHARS) {
		textCache = new WeakMap()
		textCacheChars = 0
		texts = undefined
	}

	if (texts === undefined) {
		texts = new Map()
		textCache.set(part, texts)
	}

	texts.set(item, text)
	textCacheChars += text.length

	return text
}

export class ContentBuilder {
	readonly builder = new ParagraphBuilder()
	private readonly services: ContentServices
	private readonly table: StyleTable
	private readonly fieldKinds: number[] = []
	private readonly fieldModes: number[] = []
	private readonly fieldFormats: (string | undefined)[] = []
	private hidden = 0
	private out: BuiltParagraph = new BuiltParagraph()
	private page: PageValues = { page: 1, pageFormat: undefined, numPages: 1, sectionPages: 1 }
	private story: StoryContext | undefined = undefined
	private format: ParagraphFormat | undefined = undefined
	private columnWidth = 0
	private readonly separators = new Map<number, number>()
	private readonly map: number[] = []

	constructor(services: ContentServices) {
		this.services = services
		this.table = services.styles.table
	}

	// A text item's decoded text (remembered while editing, see cachedText).
	private textOf(part: StoryPart, item: number): string {
		return this.services.recordMaps ? cachedText(part, item) : part.text(item)
	}

	// The paragraph format of an item (the paragraph's direct properties in its cell context).
	formatOf(part: StoryPart, item: number, cell: CellFormat | undefined): ParagraphFormat {
		const row = part.paragraphRow(item)
		const pPr = row < 0 ? -1 : (part.paragraphs.pPr[row] ?? -1)

		return this.services.resolver.paragraph(this.services.store.paragraph(part.document, pPr), cell)
	}

	// Whether the paragraph's mark is absent in the view, joining it to the next paragraph.
	markHidden(part: StoryPart, item: number, format: ParagraphFormat): boolean {
		if (part.paragraphMarkHidden(item, this.services.view)) return true

		if (this.services.showHidden) return false

		const mark = this.services.resolver.paragraphMark(format)

		return mark.vanish || mark.specVanish
	}

	// Builds paragraphs first..last (joined: every mark but the last hidden) into the paragraph builder;
	// `parts` and `formats` are their parts and formats in order.
	build(
		story: StoryContext,
		parts: readonly StoryPart[],
		items: readonly number[],
		formats: readonly ParagraphFormat[],
		columnWidth: number,
		page: PageValues,
		out: BuiltParagraph
	): void {
		const services = this.services
		const lastFormat = formats[formats.length - 1]
		const first = items[0] ?? -1
		const firstPart = parts[0] ?? story.part

		out.reset()

		if (lastFormat === undefined) return

		this.out = out
		this.page = page
		this.story = story
		this.columnWidth = columnWidth
		this.map.length = 0
		out.first = first
		out.firstPart = firstPart.index
		out.last = items[items.length - 1] ?? first
		out.format = lastFormat
		out.markStyle = services.styles.style(services.resolver.paragraphMark(lastFormat), "ascii")

		const dropCap = lastFormat.props.frame?.dropCap

		out.dropCap = dropCap === "drop" || dropCap === "margin" ? dropCap : undefined

		const builder = this.builder.reset(out.markStyle)

		services.notes.beginParagraph()

		this.marker(story, firstPart, first, formats[0] ?? lastFormat)

		for (let k = 0; k < items.length; k++) {
			const format = formats[k] ?? lastFormat

			this.format = format
			this.paragraph(parts[k] ?? firstPart, items[k] ?? -1, format)
		}

		out.text = builder.finish().text

		if (services.recordMaps) out.map = Int32Array.from(this.map)

		this.story = undefined
		this.format = undefined
	}

	private marker(story: StoryContext, part: StoryPart, item: number, format: ParagraphFormat): void {
		const services = this.services
		const numbering = format.props.numbering

		if (numbering === undefined || format.level === undefined || story.counters === undefined) return

		const key = services.memoKey(part.index, item)
		let marker = services.markers.get(key)

		if (marker === undefined) {
			marker = story.counters.next(numbering.numId, numbering.ilvl) ?? null
			services.markers.set(key, marker)
		}

		if (marker === null) return

		const run = services.resolver.listMarker(format)

		if (run === undefined || (run.vanish && !services.showHidden)) return

		const builder = this.builder
		const text = marker.pictureBullet !== undefined && marker.text === "" ? "•" : marker.text
		const start = builder.length

		if (text !== "") {
			services.styles.forEachStyle(text, run, TrackedLook.NONE, (from, to, style) => {
				const drawn = services.styles.drawn(style, text.slice(from, to))

				builder.append(drawn, style)
				this.out.markerWidth += this.width(style, drawn, 0, drawn.length)
			})
		}

		const style = services.styles.style(run, "ascii")

		if (marker.suffix === "tab") builder.tab(style)
		else if (marker.suffix === "space") {
			builder.append(" ", style)
		}

		this.out.markerAlignment = marker.alignment
		this.out.empty = builder.length === start && marker.suffix === "nothing"
	}

	private width(style: number, text: string, from: number, to: number): number {
		const table = this.table

		return (
			(table.cache.width(table.fontKey[style] ?? 0, text, from, to) + (table.letterSpacing[style] ?? 0) * (to - from)) *
			(table.scaleX[style] ?? 1)
		)
	}

	private paragraph(part: StoryPart, paragraph: number, format: ParagraphFormat): void {
		const services = this.services
		const view = services.view
		const store = services.store
		const resolver = services.resolver
		const builder = this.builder
		const kinds = part.kinds
		const end = part.end(paragraph)
		const base = resolver.run(format, EMPTY_RUN)
		let run: ResolvedRunProps = base
		let runEnd = -1
		let revision: TrackedLook = TrackedLook.NONE
		let runHidden = false
		const map = services.recordMaps ? this.map : undefined

		this.fieldKinds.length = 0
		this.fieldModes.length = 0
		this.fieldFormats.length = 0
		this.hidden = 0

		for (let item = paragraph + 1; item < end;) {
			const kind = kinds[item] ?? 0
			let next = part.end(item)

			if (item >= runEnd && runEnd >= 0) {
				run = base
				runEnd = -1
				revision = TrackedLook.NONE
				runHidden = false
			}

			const before = builder.length

			switch (kind) {
				case Item.Run: {
					const row = part.runs.rPr[part.aux[item] ?? -1] ?? -1

					run = resolver.run(format, store.run(part.document, row))
					runEnd = next
					revision = view === RevisionView.Markup ? this.revisionOf(part, item, paragraph) : TrackedLook.NONE
					runHidden = (run.vanish || run.specVanish) && !services.showHidden
					next = item + 1

					break
				}
				case Item.Sdt:
				case Item.CustomXml:
				case Item.SmartTag:
				case Item.Hyperlink:
				case Item.Bidi:
				case Item.AlternateContent:
					next = item + 1

					break
				case Item.Insertion:
				case Item.MoveTo:
					next = view === RevisionView.Original ? next : item + 1

					break
				case Item.Deletion:
				case Item.MoveFrom:
					next = view === RevisionView.Final ? next : item + 1

					break
				case Item.FieldSimple: {
					const info = fieldInfo(part.fieldInstruction(part.aux[item] ?? -1, view))

					if (info.kind === FieldKind.OTHER) {
						next = item + 1
					} else if (this.hidden === 0) {
						const firstRun = this.firstRun(part, item)
						const fieldRun =
							firstRun < 0
								? run
								: resolver.run(format, store.run(part.document, part.runs.rPr[part.aux[firstRun] ?? -1] ?? -1))

						this.computedField(info.kind, info.format, fieldRun, revision)
					}

					break
				}
				case Item.Text:
				case Item.DeletedText:
					if (this.hidden === 0 && !runHidden) this.text(this.textOf(part, item), run, revision)

					break
				case Item.Tab:
					if (this.hidden === 0 && !runHidden) this.tab(run, revision)

					break
				case Item.PositionalTab:
					if (this.hidden === 0 && !runHidden) this.positionalTab(part, item, run, revision)

					break
				case Item.Break:
					if (this.hidden === 0 && !runHidden) this.lineBreak(part, item, run, revision)

					break
				case Item.CarriageReturn:
					if (this.hidden === 0 && !runHidden) {
						builder.lineBreak(this.services.styles.style(run, "ascii", revision))
						this.out.empty = false
					}

					break
				case Item.LastRenderedPageBreak:
					this.out.renderedBreaks.push(builder.length)

					break
				case Item.Symbol:
					if (this.hidden === 0 && !runHidden) {
						const symbol = part.symbol(item)

						if (symbol.char > 0 && symbol.char <= 0x10ffff) {
							const style = services.styles.style(run, "ascii", revision, symbol.font)

							builder.append(services.styles.drawn(style, String.fromCodePoint(symbol.char)), style)
							this.out.empty = false
						}
					}

					break
				case Item.SoftHyphen:
					if (this.hidden === 0 && !runHidden) builder.append("­", services.styles.style(run, "ascii", revision))

					break
				case Item.NoBreakHyphen:
					if (this.hidden === 0 && !runHidden) this.text("‑", run, revision)

					break
				case Item.FieldChar:
					this.fieldChar(part, item, view, run, revision)

					break
				case Item.FootnoteReference:
				case Item.EndnoteReference:
					if (this.hidden === 0 && !runHidden) this.noteReference(part, item, kind === Item.EndnoteReference, run, revision)

					break
				case Item.FootnoteMark:
				case Item.EndnoteMark:
					if (this.hidden === 0 && !runHidden && this.story !== undefined) this.text(this.story.noteMark, run, revision)

					break
				case Item.Separator:
				case Item.ContinuationSeparator:
					if (this.hidden === 0) this.separator(kind === Item.ContinuationSeparator, run, revision)

					break
				case Item.PageNumber:
					if (this.hidden === 0 && !runHidden) this.computedField(FieldKind.PAGE, undefined, run, revision)

					break
				case Item.Drawing:
				case Item.VmlPicture:
				case Item.EmbeddedObject:
				case Item.Math:
				case Item.MathParagraph:
					if (this.hidden === 0 && !runHidden) this.object(part, item, kind, run, revision)

					break
				case Item.Ruby:
					if (this.hidden === 0 && !runHidden) this.ruby(part, item, run, revision)

					break
				case Item.Paragraph:
					// A paragraph inside a paragraph only comes from an AlternateContent branch; its text joins in.
					next = item + 1

					break
			}

			if (map !== undefined && builder.length > before) map.push(part.index, item, before, builder.length)

			item = next
		}
	}

	private firstRun(part: StoryPart, container: number): number {
		for (let item = container + 1; item < part.end(container); item++) if (part.kinds[item] === Item.Run) return item

		return -1
	}

	// The tracked change a run sits in, with its author (see revisionLook).
	private revisionOf(part: StoryPart, item: number, paragraph: number): TrackedLook {
		for (let at = part.parent(item); at > paragraph; at = part.parent(at)) {
			const look =
				part.kinds[at] === Item.Insertion
					? TrackedLook.INSERTED
					: part.kinds[at] === Item.Deletion
						? TrackedLook.DELETED
						: part.kinds[at] === Item.MoveFrom
							? TrackedLook.MOVED_FROM
							: part.kinds[at] === Item.MoveTo
								? TrackedLook.MOVED_TO
								: TrackedLook.NONE

			if (look === TrackedLook.NONE) continue

			const node = part.revisions.node[part.aux[at] ?? -1] ?? -1
			const author = node < 0 ? undefined : part.reader.attr(node, Ns.W, "author")

			return revisionLook(look, this.services.authors(author))
		}

		return TrackedLook.NONE
	}

	private text(text: string, run: ResolvedRunProps, revision: TrackedLook): void {
		if (text.length === 0) return

		const builder = this.builder

		const styles = this.services.styles

		styles.forEachStyle(text, run, revision, (from, to, style) => {
			builder.append(styles.drawn(style, from === 0 && to === text.length ? text : text.slice(from, to)), style)
		})
		this.out.empty = false
	}

	private tab(run: ResolvedRunProps, revision: TrackedLook): void {
		this.builder.tab(this.services.styles.style(run, "ascii", revision))
		this.out.empty = false
	}

	// w:ptab: a tab to a position relative to the margins or indents, ignoring the paragraph's stops.
	private positionalTab(part: StoryPart, item: number, run: ResolvedRunProps, revision: TrackedLook): void {
		const reader = part.reader
		const node = part.node(item)
		const alignment = reader.attr(node, Ns.W, "alignment") ?? "left"
		const relative = reader.attr(node, Ns.W, "relativeTo") ?? "margin"
		const leader = LEADERS[reader.attr(node, Ns.W, "leader") ?? "none"] ?? TabLeader.NONE
		const indent = this.format?.props.indent
		const start = relative === "indent" && indent !== undefined ? twips(indent.start) : 0
		const end = this.columnWidth - (relative === "indent" && indent !== undefined ? twips(indent.end) : 0)
		const position = alignment === "center" ? (start + end) / 2 : alignment === "right" ? end : start

		this.builder.tab(this.services.styles.style(run, "ascii", revision), {
			position,
			alignment: alignment === "center" ? TabAlignment.CENTER : alignment === "right" ? TabAlignment.END : TabAlignment.START,
			leader
		})
		this.out.empty = false
	}

	private lineBreak(part: StoryPart, item: number, run: ResolvedRunProps, revision: TrackedLook): void {
		const type = part.breakType(item)
		const style = this.services.styles.style(run, "ascii", revision)
		let kind: BreakKind = BreakKind.LINE

		if (type === BreakType.Page && this.story?.breaksAsLines !== true) kind = BreakKind.PAGE
		else if (type === BreakType.Column && this.story?.breaksAsLines !== true) kind = BreakKind.COLUMN
		else if (type === BreakType.TextWrapping) {
			const clear = part.breakClear(item)

			kind =
				clear === BreakClear.Left
					? BreakKind.LINE_CLEAR_LEFT
					: clear === BreakClear.Right
						? BreakKind.LINE_CLEAR_RIGHT
						: clear === BreakClear.All
							? BreakKind.LINE_CLEAR_ALL
							: BreakKind.LINE
		}

		this.builder.lineBreak(style, kind)
		this.out.empty = false
	}

	private fieldChar(part: StoryPart, item: number, view: RevisionView, run: ResolvedRunProps, revision: TrackedLook): void {
		const type = part.fieldCharType(item)
		const kinds = this.fieldKinds
		const modes = this.fieldModes

		if (type === FieldCharType.Begin) {
			const field = part.fieldOf(item)
			const info = field < 0 ? undefined : fieldInfo(part.fieldInstruction(field, view))

			kinds.push(info?.kind ?? FieldKind.OTHER)
			this.fieldFormats.push(info?.format)
			modes.push(FIELD_CODE)
			this.hidden++

			return
		}

		const depth = modes.length

		if (depth === 0) return

		const mode = modes[depth - 1] ?? FIELD_RESULT
		const kind = (kinds[depth - 1] ?? FieldKind.OTHER) as FieldKind

		if (type === FieldCharType.Separate) {
			if (mode !== FIELD_CODE) return

			if (kind === FieldKind.OTHER) {
				modes[depth - 1] = FIELD_RESULT
				this.hidden--
			} else {
				modes[depth - 1] = FIELD_REPLACED

				if (this.hidden === 1) this.computedField(kind, this.fieldFormats[depth - 1], run, revision)
			}

			return
		}

		if (type === FieldCharType.End) {
			kinds.pop()
			modes.pop()

			const format = this.fieldFormats.pop()

			if (mode === FIELD_RESULT) return

			this.hidden--

			// A computed field without a stored result still shows its value.
			if (mode === FIELD_CODE && kind !== FieldKind.OTHER && this.hidden === 0) this.computedField(kind, format, run, revision)
		}
	}

	private computedField(kind: FieldKind, format: string | undefined, run: ResolvedRunProps, revision: TrackedLook): void {
		const page = this.page

		if (kind === FieldKind.PAGE) {
			const start = this.builder.length

			this.out.fieldFlags |= 1
			this.text(numberText(page.page, format ?? page.pageFormat), run, revision)
			this.out.pageFields.push(start, this.builder.length)

			return
		}

		this.out.fieldFlags |= 2
		this.text(numberText(kind === FieldKind.NUMPAGES ? page.numPages : page.sectionPages, format), run, revision)
	}

	private noteReference(part: StoryPart, item: number, endnote: boolean, run: ResolvedRunProps, revision: TrackedLook): void {
		const services = this.services
		const id = String(part.aux[item] ?? -1)
		const story = (endnote ? services.model.endnotes : services.model.footnotes).get(id)
		const offset = this.builder.length

		if (story === undefined) return

		const custom = part.reader.attr(part.node(item), Ns.W, "customMarkFollows")
		const mark =
			custom === "1" || custom === "true" || custom === "on" ? "" : services.notes.reference(endnote, part.index, item, story.index)

		this.text(mark, run, revision)
		this.out.noteOffsets.push(offset)
		this.out.noteStories.push(story.index)
		this.out.noteEndnote.push(endnote ? 1 : 0)
		this.out.empty = false
	}

	private separator(continuation: boolean, run: ResolvedRunProps, revision: TrackedLook): void {
		const services = this.services
		const kind = continuation ? ObjectKind.CONTINUATION_SEPARATOR : ObjectKind.SEPARATOR
		// Word's separator is 144 pt wide (2 in); the continuation separator spans the column.
		const width = continuation ? this.columnWidth : Math.min(this.columnWidth, 192)
		const key = Math.round(width * 8) * 2 + (continuation ? 1 : 0)
		let id = this.separators.get(key)

		if (id === undefined) {
			id = services.objects.add(kind, -1, -1, -1, width, 1, 1, undefined, undefined, undefined)
			this.separators.set(key, id)
		}

		this.builder.object(services.styles.style(run, "ascii", revision), width, halfPoints(run.size) * 0.35, 0, id)
		this.out.empty = false
	}

	private object(part: StoryPart, item: number, kind: number, run: ResolvedRunProps, revision: TrackedLook): void {
		const services = this.services
		const style = services.styles.style(run, "ascii", revision)
		const size = halfPoints(run.size)
		const placed = objectAt(
			services.model,
			services.objects,
			services.sizer,
			part,
			item,
			size,
			text => this.width(style, text, 0, text.length),
			services.styles.hasMissing
		)

		if (placed === undefined) return

		if (placed.anchored) {
			this.out.floats.push(placed.id)

			return
		}

		const display = kind === Item.MathParagraph
		const builder = this.builder

		// Display equations take a line of their own.
		if (display && builder.length > 0) builder.lineBreak(style)

		// Word's line spacing multiple stretches the text of a line, not the objects on it: the box is shrunk
		// by the multiple the layout applies to the whole line.
		const spacing = this.format?.props.spacing
		const multiple =
			!services.stretchObjects && spacing !== undefined && spacing.lineRule === "auto" && spacing.line > 0 ? spacing.line / 240 : 1

		builder.object(style, placed.width, placed.ascent / multiple, (placed.height - placed.ascent) / multiple, placed.id)

		if ((services.objects.textboxes[placed.id]?.length ?? 0) > 0) this.out.boxedObjects.push(placed.id)

		this.out.empty = false
	}

	// Ruby: the base text (the annotation above it is not laid out yet).
	private ruby(part: StoryPart, item: number, run: ResolvedRunProps, revision: TrackedLook): void {
		for (let child = item + 1; child < part.end(item); child = part.end(child)) {
			if (part.kinds[child] !== Item.RubyBase) continue

			for (let inner = child + 1; inner < part.end(child); inner++) {
				if (part.kinds[inner] === Item.Text) this.text(this.textOf(part, inner), run, revision)
			}
		}
	}
}
