import {
	DECORATION_FLOATS,
	DECORATION_INTS,
	FRAGMENT_FLOATS,
	FRAGMENT_INTS,
	LINE_FLOATS,
	LINE_INTS,
	LineInt,
	TextLayout,
	type DecorationKind
} from "../../text/index"

// A page's display list: the text lines of every story on it (core/text layout), plus graphics (fills,
// borders, rules), placed objects and the paragraphs the lines come from, all in typed arrays. Coordinates
// are px from the page's top left corner.

// graphicFloats: x, y, width, height. graphicInts: kind, key, layer.
export const GRAPHIC_FLOATS = 4
export const GRAPHIC_INTS = 3

export const PageGraphic = {
	// A fill: key is a DecorationTable colour key.
	SHADING: 1,
	// A border line drawn as its box (width or height is the line's thickness): key is a border key.
	BORDER: 2,
	// The rule between columns (w:cols w:sep).
	COLUMN_SEPARATOR: 3,
	// A note separator drawn without a separator story.
	NOTE_SEPARATOR: 4,
	// A changed-line bar in the margin (tracked changes shown as markup).
	REVISION_BAR: 5
} as const

export type PageGraphic = (typeof PageGraphic)[keyof typeof PageGraphic]

export const Layer = {
	// Behind the text: fills, floating objects with behindDoc.
	BEHIND: 0,
	// With the text: borders, inline and wrapped objects.
	TEXT: 1,
	// In front of the text.
	FRONT: 2
} as const

export type Layer = (typeof Layer)[keyof typeof Layer]

// Clip rectangles (rows of exact height cut their cells' content): clipFloats x, y, width, height; clipInts
// the lines [first, end) and objects [first, end) drawn through the clip.
export const CLIP_FLOATS = 4
export const CLIP_INTS = 4

// objectFloats: x, y, width, height. objectInts: object id (DocumentLayout.objects), layer, source.
export const OBJECT_FLOATS = 4
export const OBJECT_INTS = 3

// Where a source paragraph lives on the page.
export const Region = {
	BODY: 0,
	HEADER: 1,
	FOOTER: 2,
	FOOTNOTE: 3,
	ENDNOTE: 4,
	TEXTBOX: 5,
	// Line numbers in the margin (w:lnNumType); the source names no item.
	LINE_NUMBER: 6,
	// A list marker, a page number or other generated text with no paragraph of its own.
	GENERATED: 7
} as const

export type Region = (typeof Region)[keyof typeof Region]

export const SourceFlag = {
	// The paragraph sits in a table cell.
	IN_TABLE: 1,
	// The paragraph's text contains page-dependent fields.
	PAGE_FIELDS: 2,
	// The text depends on the document's page count (NUMPAGES, SECTIONPAGES).
	TOTAL_FIELDS: 4
} as const

// sourceInts: region, story, item (the first paragraph of a joined group), flags. Line SOURCE values index
// these; the source's text is what the line's text offsets index.
export const SOURCE_INTS = 4

export const SourceInt = { REGION: 0, STORY: 1, ITEM: 2, FLAGS: 3 } as const

// What append() copies from: a growing output or a finished page.
export interface OutputView {
	readonly text: TextLayout
	readonly graphicCount: number
	readonly graphicFloats: Float32Array
	readonly graphicInts: Int32Array
	readonly objectCount: number
	readonly objectFloats: Float32Array
	readonly objectInts: Int32Array
	readonly sourceCount: number
	readonly sourceInts: Int32Array
	readonly sourceTexts: readonly string[]
	readonly clipCount: number
	readonly clipFloats: Float32Array
	readonly clipInts: Int32Array
}

// A position in an output to roll back to or move content after.
export class OutputMark {
	lines = 0
	graphics = 0
	objects = 0
	sources = 0
	clips = 0

	copy(from: OutputMark): this {
		this.lines = from.lines
		this.graphics = from.graphics
		this.objects = from.objects
		this.sources = from.sources
		this.clips = from.clips

		return this
	}
}

function growF32(a: Float32Array, size: number): Float32Array<ArrayBuffer> {
	const out = new Float32Array(Math.max(size, a.length * 2))

	out.set(a)

	return out
}

function growI32(a: Int32Array, size: number): Int32Array<ArrayBuffer> {
	const out = new Int32Array(Math.max(size, a.length * 2))

	out.set(a)

	return out
}

// The copied range [first, last + 1) of source indices [from, to) under `map` (offset by `base`), or null.
function remap(map: Int32Array, base: number, from: number, to: number): [number, number] | null {
	let first = -1
	let last = -1

	for (let k = Math.max(from, base); k < to && k - base < map.length; k++) {
		const index = map[k - base] ?? -1

		if (index < 0) continue

		if (first < 0) first = index

		last = index
	}

	return first < 0 ? null : [first, last + 1]
}

// The growing output of one page (or of a story laid out off-page, then appended).
export class PageOutput {
	readonly text = new TextLayout()
	graphicCount = 0
	graphicFloats = new Float32Array(GRAPHIC_FLOATS * 32)
	graphicInts = new Int32Array(GRAPHIC_INTS * 32)
	objectCount = 0
	objectFloats = new Float32Array(OBJECT_FLOATS * 8)
	objectInts = new Int32Array(OBJECT_INTS * 8)
	sourceCount = 0
	sourceInts = new Int32Array(SOURCE_INTS * 64)
	sourceTexts: string[] = []
	clipCount = 0
	clipFloats = new Float32Array(CLIP_FLOATS * 4)
	clipInts = new Int32Array(CLIP_INTS * 4)
	// w:lastRenderedPageBreak offsets per source (body paragraphs only), for checks against Word's breaks.
	sourceBreaks: (readonly number[] | undefined)[] = []

	reset(): void {
		this.text.reset()
		this.graphicCount = 0
		this.objectCount = 0
		this.sourceCount = 0
		this.sourceTexts.length = 0
		this.sourceBreaks.length = 0
		this.clipCount = 0
	}

	addClip(
		x: number,
		y: number,
		width: number,
		height: number,
		lineFrom: number,
		lineTo: number,
		objectFrom: number,
		objectTo: number
	): void {
		if (lineTo <= lineFrom && objectTo <= objectFrom) return

		const index = this.clipCount

		if ((index + 1) * CLIP_INTS > this.clipInts.length) {
			this.clipInts = growI32(this.clipInts, (index + 1) * CLIP_INTS)
			this.clipFloats = growF32(this.clipFloats, (index + 1) * CLIP_FLOATS)
		}

		const f = index * CLIP_FLOATS
		const i = index * CLIP_INTS

		this.clipFloats[f] = x
		this.clipFloats[f + 1] = y
		this.clipFloats[f + 2] = width
		this.clipFloats[f + 3] = height
		this.clipInts[i] = lineFrom
		this.clipInts[i + 1] = lineTo
		this.clipInts[i + 2] = objectFrom
		this.clipInts[i + 3] = objectTo
		this.clipCount++
	}

	addSource(region: Region, story: number, item: number, flags: number, text: string, breaks?: readonly number[]): number {
		const source = this.sourceCount

		if ((source + 1) * SOURCE_INTS > this.sourceInts.length) this.sourceInts = growI32(this.sourceInts, (source + 1) * SOURCE_INTS)

		const at = source * SOURCE_INTS

		this.sourceInts[at] = region
		this.sourceInts[at + 1] = story
		this.sourceInts[at + 2] = item
		this.sourceInts[at + 3] = flags
		this.sourceTexts[source] = text
		this.sourceBreaks[source] = breaks
		this.sourceCount++

		return source
	}

	sourceFlags(source: number): number {
		return this.sourceInts[source * SOURCE_INTS + SourceInt.FLAGS] ?? 0
	}

	addGraphic(kind: PageGraphic, key: number, layer: Layer, x: number, y: number, width: number, height: number): void {
		if (!(width > 0 && height > 0)) return

		const index = this.graphicCount

		if ((index + 1) * GRAPHIC_INTS > this.graphicInts.length) {
			this.graphicInts = growI32(this.graphicInts, (index + 1) * GRAPHIC_INTS)
			this.graphicFloats = growF32(this.graphicFloats, (index + 1) * GRAPHIC_FLOATS)
		}

		const f = index * GRAPHIC_FLOATS
		const i = index * GRAPHIC_INTS

		this.graphicFloats[f] = x
		this.graphicFloats[f + 1] = y
		this.graphicFloats[f + 2] = width
		this.graphicFloats[f + 3] = height
		this.graphicInts[i] = kind
		this.graphicInts[i + 1] = key
		this.graphicInts[i + 2] = layer
		this.graphicCount++
	}

	addObject(id: number, layer: Layer, source: number, x: number, y: number, width: number, height: number): void {
		const index = this.objectCount

		if ((index + 1) * OBJECT_INTS > this.objectInts.length) {
			this.objectInts = growI32(this.objectInts, (index + 1) * OBJECT_INTS)
			this.objectFloats = growF32(this.objectFloats, (index + 1) * OBJECT_FLOATS)
		}

		const f = index * OBJECT_FLOATS
		const i = index * OBJECT_INTS

		this.objectFloats[f] = x
		this.objectFloats[f + 1] = y
		this.objectFloats[f + 2] = width
		this.objectFloats[f + 3] = height
		this.objectInts[i] = id
		this.objectInts[i + 1] = layer
		this.objectInts[i + 2] = source
		this.objectCount++
	}

	mark(into: OutputMark = new OutputMark()): OutputMark {
		into.lines = this.text.lineCount
		into.graphics = this.graphicCount
		into.objects = this.objectCount
		into.sources = this.sourceCount
		into.clips = this.clipCount

		return into
	}

	rollback(mark: OutputMark): void {
		this.text.truncate(mark.lines)
		this.graphicCount = Math.min(this.graphicCount, mark.graphics)
		this.objectCount = Math.min(this.objectCount, mark.objects)
		this.clipCount = Math.min(this.clipCount, mark.clips)

		if (mark.sources < this.sourceCount) {
			this.sourceCount = mark.sources
			this.sourceTexts.length = mark.sources
			this.sourceBreaks.length = mark.sources
		}
	}

	// Moves what was added between `from` and `to` (default: now) by (dx, dy).
	move(from: OutputMark, to: OutputMark | null, dx: number, dy: number): void {
		const text = this.text
		const lineEnd = to?.lines ?? text.lineCount

		if (dx === 0 && dy === 0) return

		for (let line = from.lines; line < lineEnd; line++) {
			const f = line * LINE_FLOATS
			const i = line * LINE_INTS

			text.lineFloats[f] = (text.lineFloats[f] ?? 0) + dy
			text.lineFloats[f + 2] = (text.lineFloats[f + 2] ?? 0) + dy
			text.lineFloats[f + 3] = (text.lineFloats[f + 3] ?? 0) + dx

			if (dx !== 0) {
				const first = text.lineInts[i + LineInt.FIRST_FRAGMENT] ?? 0
				const count = text.lineInts[i + LineInt.FRAGMENT_COUNT] ?? 0

				for (let k = first; k < first + count; k++)
					text.fragmentFloats[k * FRAGMENT_FLOATS] = (text.fragmentFloats[k * FRAGMENT_FLOATS] ?? 0) + dx
			}
		}

		if (from.lines < lineEnd) {
			const firstDecoration = text.lineInts[from.lines * LINE_INTS + LineInt.FIRST_DECORATION] ?? text.decorationCount
			const lastDecoration =
				lineEnd < text.lineCount ? (text.lineInts[lineEnd * LINE_INTS + LineInt.FIRST_DECORATION] ?? 0) : text.decorationCount

			for (let d = firstDecoration; d < lastDecoration; d++) {
				const f = d * DECORATION_FLOATS

				text.decorationFloats[f] = (text.decorationFloats[f] ?? 0) + dx
				text.decorationFloats[f + 1] = (text.decorationFloats[f + 1] ?? 0) + dy
			}
		}

		const graphicEnd = to?.graphics ?? this.graphicCount

		for (let g = from.graphics; g < graphicEnd; g++) {
			const f = g * GRAPHIC_FLOATS

			this.graphicFloats[f] = (this.graphicFloats[f] ?? 0) + dx
			this.graphicFloats[f + 1] = (this.graphicFloats[f + 1] ?? 0) + dy
		}

		const objectEnd = to?.objects ?? this.objectCount

		for (let o = from.objects; o < objectEnd; o++) {
			const f = o * OBJECT_FLOATS

			this.objectFloats[f] = (this.objectFloats[f] ?? 0) + dx
			this.objectFloats[f + 1] = (this.objectFloats[f + 1] ?? 0) + dy
		}

		const clipEnd = to?.clips ?? this.clipCount

		for (let c = from.clips; c < clipEnd; c++) {
			const f = c * CLIP_FLOATS

			this.clipFloats[f] = (this.clipFloats[f] ?? 0) + dx
			this.clipFloats[f + 1] = (this.clipFloats[f + 1] ?? 0) + dy
		}
	}

	// Bottom of the content added since `from` (lines, graphics and objects), or -Infinity.
	bottomSince(from: OutputMark): number {
		let bottom = -Infinity
		const text = this.text

		for (let line = from.lines; line < text.lineCount; line++) {
			bottom = Math.max(bottom, (text.lineFloats[line * LINE_FLOATS] ?? 0) + (text.lineFloats[line * LINE_FLOATS + 1] ?? 0))
		}

		for (let g = from.graphics; g < this.graphicCount; g++) {
			bottom = Math.max(bottom, (this.graphicFloats[g * GRAPHIC_FLOATS + 1] ?? 0) + (this.graphicFloats[g * GRAPHIC_FLOATS + 3] ?? 0))
		}

		return bottom
	}

	// Appends what `source` holds between `from` and `to` (default: all of it) whose top lies in
	// [clipTop, clipBottom), moved by (dx, dy). Sources come along whole (up to `to`), renumbered.
	append(
		source: OutputView,
		dx: number,
		dy: number,
		clipTop = -Infinity,
		clipBottom = Infinity,
		from: OutputMark | null = null,
		to: OutputMark | null = null
	): void {
		const sourceOffset = this.sourceCount
		const sourceEnd = Math.min(source.sourceCount, to?.sources ?? source.sourceCount)

		for (let s = 0; s < sourceEnd; s++) {
			const at = s * SOURCE_INTS

			this.addSource(
				(source.sourceInts[at] ?? 0) as Region,
				source.sourceInts[at + 1] ?? -1,
				source.sourceInts[at + 2] ?? -1,
				source.sourceInts[at + 3] ?? 0,
				source.sourceTexts[s] ?? ""
			)
		}

		const text = source.text
		const out = this.text
		const lineStart = from?.lines ?? 0
		const lineEnd = Math.min(text.lineCount, to?.lines ?? text.lineCount)
		// New index of each copied line (-1 when clipped away), for clip rectangles.
		const lineMap = source.clipCount > 0 ? new Int32Array(Math.max(0, lineEnd - lineStart)).fill(-1) : null

		for (let line = lineStart; line < lineEnd; line++) {
			const f = line * LINE_FLOATS
			const i = line * LINE_INTS
			const top = text.lineFloats[f] ?? 0

			if (top < clipTop - 1e-3 || top >= clipBottom - 1e-3) continue

			if (lineMap !== null) lineMap[line - lineStart] = out.lineCount

			out.addLine(
				top + dy,
				text.lineFloats[f + 1] ?? 0,
				(text.lineFloats[f + 2] ?? 0) + dy,
				(text.lineFloats[f + 3] ?? 0) + dx,
				text.lineFloats[f + 4] ?? 0,
				text.lineInts[i + LineInt.TEXT_START] ?? 0,
				text.lineInts[i + LineInt.TEXT_END] ?? 0,
				text.lineInts[i + LineInt.FLAGS] ?? 0,
				(text.lineInts[i + LineInt.SOURCE] ?? 0) + sourceOffset
			)

			const first = text.lineInts[i + LineInt.FIRST_FRAGMENT] ?? 0
			const count = text.lineInts[i + LineInt.FRAGMENT_COUNT] ?? 0

			for (let k = first; k < first + count; k++) {
				const ff = k * FRAGMENT_FLOATS
				const fi = k * FRAGMENT_INTS

				out.addFragment(
					(text.fragmentFloats[ff] ?? 0) + dx,
					text.fragmentFloats[ff + 1] ?? 0,
					text.fragmentFloats[ff + 2] ?? 0,
					text.fragmentFloats[ff + 3] ?? 0,
					text.fragmentInts[fi] ?? 0,
					text.fragmentInts[fi + 1] ?? 0,
					text.fragmentInts[fi + 2] ?? 0,
					text.fragmentInts[fi + 3] ?? 0,
					text.fragmentInts[fi + 4] ?? 0
				)
			}

			const firstDecoration = text.lineInts[i + LineInt.FIRST_DECORATION] ?? 0
			const lastDecoration =
				line + 1 < text.lineCount ? (text.lineInts[i + LINE_INTS + LineInt.FIRST_DECORATION] ?? 0) : text.decorationCount

			for (let d = firstDecoration; d < lastDecoration; d++) {
				const df = d * DECORATION_FLOATS
				const di = d * DECORATION_INTS

				out.addDecoration(
					(text.decorationInts[di] ?? 0) as DecorationKind,
					text.decorationInts[di + 1] ?? 0,
					(text.decorationFloats[df] ?? 0) + dx,
					(text.decorationFloats[df + 1] ?? 0) + dy,
					text.decorationFloats[df + 2] ?? 0,
					text.decorationFloats[df + 3] ?? 0
				)
			}
		}

		const graphicEnd = Math.min(source.graphicCount, to?.graphics ?? source.graphicCount)

		for (let g = from?.graphics ?? 0; g < graphicEnd; g++) {
			const f = g * GRAPHIC_FLOATS
			const i = g * GRAPHIC_INTS
			const y = source.graphicFloats[f + 1] ?? 0

			if (y < clipTop - 1e-3 || y >= clipBottom - 1e-3) continue

			this.addGraphic(
				(source.graphicInts[i] ?? 0) as PageGraphic,
				source.graphicInts[i + 1] ?? 0,
				(source.graphicInts[i + 2] ?? 0) as Layer,
				(source.graphicFloats[f] ?? 0) + dx,
				y + dy,
				source.graphicFloats[f + 2] ?? 0,
				source.graphicFloats[f + 3] ?? 0
			)
		}

		const objectStart = from?.objects ?? 0
		const objectEnd = Math.min(source.objectCount, to?.objects ?? source.objectCount)
		const objectMap = source.clipCount > 0 ? new Int32Array(Math.max(0, objectEnd - objectStart)).fill(-1) : null

		for (let o = objectStart; o < objectEnd; o++) {
			const f = o * OBJECT_FLOATS
			const i = o * OBJECT_INTS
			const y = source.objectFloats[f + 1] ?? 0

			if (y < clipTop - 1e-3 || y >= clipBottom - 1e-3) continue

			const objectSource = source.objectInts[i + 2] ?? -1

			if (objectMap !== null) objectMap[o - objectStart] = this.objectCount

			this.addObject(
				source.objectInts[i] ?? 0,
				(source.objectInts[i + 1] ?? 0) as Layer,
				objectSource < 0 ? -1 : objectSource + sourceOffset,
				(source.objectFloats[f] ?? 0) + dx,
				y + dy,
				source.objectFloats[f + 2] ?? 0,
				source.objectFloats[f + 3] ?? 0
			)
		}

		if (lineMap === null || objectMap === null) return

		const clipEnd = Math.min(source.clipCount, to?.clips ?? source.clipCount)

		for (let c = from?.clips ?? 0; c < clipEnd; c++) {
			const f = c * CLIP_FLOATS
			const i = c * CLIP_INTS
			const lines = remap(lineMap, lineStart, source.clipInts[i] ?? 0, source.clipInts[i + 1] ?? 0)
			const objects = remap(objectMap, objectStart, source.clipInts[i + 2] ?? 0, source.clipInts[i + 3] ?? 0)

			if (lines === null && objects === null) continue

			this.addClip(
				(source.clipFloats[f] ?? 0) + dx,
				(source.clipFloats[f + 1] ?? 0) + dy,
				source.clipFloats[f + 2] ?? 0,
				source.clipFloats[f + 3] ?? 0,
				lines?.[0] ?? this.text.lineCount,
				lines?.[1] ?? this.text.lineCount,
				objects?.[0] ?? this.objectCount,
				objects?.[1] ?? this.objectCount
			)
		}
	}
}

// A finished page: tight typed arrays, transferable to another thread.
export class PageDisplayList implements OutputView {
	readonly index: number
	// The page number fields show, and its text in the section's format.
	readonly number: number
	readonly numberText: string
	readonly section: number
	readonly width: number
	readonly height: number
	readonly text: TextLayout
	readonly graphicCount: number
	readonly graphicFloats: Float32Array
	readonly graphicInts: Int32Array
	readonly objectCount: number
	readonly objectFloats: Float32Array
	readonly objectInts: Int32Array
	readonly sourceCount: number
	readonly sourceInts: Int32Array
	readonly sourceTexts: readonly string[]
	readonly clipCount: number
	readonly clipFloats: Float32Array
	readonly clipInts: Int32Array
	// Body blocks on the page: the first and last body item (paragraph or table) with content here.
	readonly bodyFirst: number
	readonly bodyLast: number

	constructor(
		output: OutputView,
		info: {
			index: number
			number: number
			numberText: string
			section: number
			width: number
			height: number
			bodyFirst: number
			bodyLast: number
		}
	) {
		this.index = info.index
		this.number = info.number
		this.numberText = info.numberText
		this.section = info.section
		this.width = info.width
		this.height = info.height
		this.bodyFirst = info.bodyFirst
		this.bodyLast = info.bodyLast
		this.text = output.text.compact()
		this.graphicCount = output.graphicCount
		this.graphicFloats = output.graphicFloats.slice(0, output.graphicCount * GRAPHIC_FLOATS)
		this.graphicInts = output.graphicInts.slice(0, output.graphicCount * GRAPHIC_INTS)
		this.objectCount = output.objectCount
		this.objectFloats = output.objectFloats.slice(0, output.objectCount * OBJECT_FLOATS)
		this.objectInts = output.objectInts.slice(0, output.objectCount * OBJECT_INTS)
		this.sourceCount = output.sourceCount
		this.sourceInts = output.sourceInts.slice(0, output.sourceCount * SOURCE_INTS)
		this.sourceTexts = output.sourceTexts.slice(0, output.sourceCount)
		this.clipCount = output.clipCount
		this.clipFloats = output.clipFloats.slice(0, output.clipCount * CLIP_FLOATS)
		this.clipInts = output.clipInts.slice(0, output.clipCount * CLIP_INTS)
	}

	// Bytes of the typed arrays. The source texts are not included: each is the laid-out text of a whole
	// paragraph (built once per layout of it, list markers and field results included), shared by the pages it
	// spans in this realm but copied with every page that is sent to another thread; see textLength.
	get byteLength(): number {
		return (
			this.text.byteLength +
			this.graphicFloats.byteLength +
			this.graphicInts.byteLength +
			this.objectFloats.byteLength +
			this.objectInts.byteLength +
			this.sourceInts.byteLength +
			this.clipFloats.byteLength +
			this.clipInts.byteLength
		)
	}

	// UTF-16 code units of the source texts (two bytes each when copied).
	get textLength(): number {
		let length = 0

		for (const text of this.sourceTexts) length += text.length

		return length
	}

	sourceRegion(source: number): Region {
		return (this.sourceInts[source * SOURCE_INTS] ?? 0) as Region
	}

	sourceItem(source: number): number {
		return this.sourceInts[source * SOURCE_INTS + SourceInt.ITEM] ?? -1
	}

	sourceStory(source: number): number {
		return this.sourceInts[source * SOURCE_INTS + SourceInt.STORY] ?? -1
	}
}
