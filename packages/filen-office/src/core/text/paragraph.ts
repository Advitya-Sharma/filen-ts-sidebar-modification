// A paragraph's content as the layout reads it: one string in logical order plus typed side tables.
// Special characters stand for non-text content: TAB_CHAR a tab, BREAK_CHAR a break inside the paragraph
// (its kind in the break table), OBJECT_CHAR an inline object (its box in the object table).

export const TAB_CHAR = "\t"
// VT, as Word itself encodes a text-wrapping break: Line_Break BK, Bidi_Class S (not a paragraph end).
export const BREAK_CHAR = "\u000b"
export const OBJECT_CHAR = "￼"

export const BreakKind = {
	LINE: 0,
	PAGE: 1,
	COLUMN: 2,
	// Line breaks that continue below floating objects on the left, right or both sides (w:br clear).
	LINE_CLEAR_LEFT: 3,
	LINE_CLEAR_RIGHT: 4,
	LINE_CLEAR_ALL: 5
} as const

export type BreakKind = (typeof BreakKind)[keyof typeof BreakKind]

export const TabAlignment = { START: 0, CENTER: 1, END: 2, DECIMAL: 3, BAR: 4 } as const

export type TabAlignment = (typeof TabAlignment)[keyof typeof TabAlignment]

export const TabLeader = { NONE: 0, DOT: 1, HYPHEN: 2, UNDERSCORE: 3, HEAVY: 4, MIDDLE_DOT: 5 } as const

export type TabLeader = (typeof TabLeader)[keyof typeof TabLeader]

// A tab's own stop (Word's w:ptab), overriding the paragraph's stops: position in px from the column's
// start edge.
export interface ExplicitTab {
	readonly position: number
	readonly alignment: TabAlignment
	readonly leader: TabLeader
}

export interface ParagraphText {
	readonly text: string
	readonly length: number
	// Style index of each run; runs are [runStart[i], runStart[i + 1]) with the last ending at length.
	readonly runCount: number
	readonly runStart: Int32Array
	readonly runStyle: Int32Array
	// Objects in text order: width, ascent, descent (px) and the caller's id.
	readonly objectCount: number
	readonly objectBox: Float64Array
	readonly objectId: Int32Array
	// Breaks in text order.
	readonly breakCount: number
	readonly breakKind: Uint8Array
	// Tabs in text order: position (NaN for the paragraph's stops), alignment, leader.
	readonly tabCount: number
	readonly tabPosition: Float64Array
	readonly tabAlignment: Uint8Array
	readonly tabLeader: Uint8Array
	// Style of the paragraph mark: the height of an empty paragraph and part of its last line's.
	readonly markStyle: number
}

// Builds ParagraphText; reusable, so a document's paragraphs share its buffers. The result is a view of
// the builder valid until the next reset().
export class ParagraphBuilder implements ParagraphText {
	text = ""
	length = 0
	runCount = 0
	runStart = new Int32Array(16)
	runStyle = new Int32Array(16)
	objectCount = 0
	objectBox = new Float64Array(12)
	objectId = new Int32Array(4)
	breakCount = 0
	breakKind = new Uint8Array(4)
	tabCount = 0
	tabPosition = new Float64Array(8)
	tabAlignment = new Uint8Array(8)
	tabLeader = new Uint8Array(8)
	markStyle = 0

	private parts: string[] = []

	reset(markStyle: number): this {
		this.parts.length = 0
		this.text = ""
		this.length = 0
		this.runCount = 0
		this.objectCount = 0
		this.breakCount = 0
		this.tabCount = 0
		this.markStyle = markStyle

		// Run tables grown by a huge paragraph are released.
		if (this.runStart.length > 1 << 16) {
			this.runStart = new Int32Array(16)
			this.runStyle = new Int32Array(16)
		}

		return this
	}

	// Text of one style. Tabs and VT become tabs and line breaks, CR and LF spaces, and U+FFFC (reserved
	// for objects) U+FFFD.
	append(text: string, style: number): this {
		if (text.length === 0) return this

		if (SPECIAL.test(text)) {
			let from = 0

			for (let i = 0; i < text.length; i++) {
				const c = text.charCodeAt(i)

				if (c !== 0x09 && c !== 0x0a && c !== 0x0b && c !== 0x0d && c !== 0xfffc) continue

				this.raw(text.slice(from, i), style)

				if (c === 0x09) this.tab(style)
				else if (c === 0x0b) this.lineBreak(style)
				else this.raw(c === 0xfffc ? "\ufffd" : " ", style)

				from = i + 1
			}

			return this.raw(text.slice(from), style)
		}

		return this.raw(text, style)
	}

	private raw(text: string, style: number): this {
		if (text.length === 0) return this

		this.run(style)
		this.parts.push(text)
		this.length += text.length

		return this
	}

	tab(style: number, explicit?: ExplicitTab): this {
		if (this.tabCount === this.tabPosition.length) {
			this.tabPosition = growFloat64(this.tabPosition)
			this.tabAlignment = growUint8(this.tabAlignment)
			this.tabLeader = growUint8(this.tabLeader)
		}

		this.tabPosition[this.tabCount] = explicit?.position ?? Number.NaN
		this.tabAlignment[this.tabCount] = explicit?.alignment ?? TabAlignment.START
		this.tabLeader[this.tabCount] = explicit?.leader ?? TabLeader.NONE
		this.tabCount++

		return this.raw(TAB_CHAR, style)
	}

	lineBreak(style: number, kind: BreakKind = BreakKind.LINE): this {
		if (this.breakCount === this.breakKind.length) this.breakKind = growUint8(this.breakKind)

		this.breakKind[this.breakCount++] = kind

		return this.raw(BREAK_CHAR, style)
	}

	// An inline box sitting on the baseline: `ascent` above it, `descent` below.
	object(style: number, width: number, ascent: number, descent: number, id: number): this {
		if (this.objectCount === this.objectId.length) {
			this.objectBox = growFloat64(this.objectBox)
			this.objectId = growInt32(this.objectId)
		}

		const at = this.objectCount * 3

		this.objectBox[at] = width
		this.objectBox[at + 1] = ascent
		this.objectBox[at + 2] = descent
		this.objectId[this.objectCount++] = id

		return this.raw(OBJECT_CHAR, style)
	}

	finish(): ParagraphText {
		this.text = this.parts.length === 1 ? (this.parts[0] ?? "") : this.parts.join("")
		this.parts.length = 0

		return this
	}

	private run(style: number): void {
		if (this.runCount > 0 && this.runStyle[this.runCount - 1] === style) return

		if (this.runCount === this.runStart.length) {
			this.runStart = growInt32(this.runStart)
			this.runStyle = growInt32(this.runStyle)
		}

		this.runStart[this.runCount] = this.length
		this.runStyle[this.runCount++] = style
	}
}

const SPECIAL = /[\t\n\v\r\ufffc]/

function growInt32(a: Int32Array): Int32Array<ArrayBuffer> {
	const out = new Int32Array(a.length * 2)

	out.set(a)

	return out
}

function growFloat64(a: Float64Array): Float64Array<ArrayBuffer> {
	const out = new Float64Array(a.length * 2)

	out.set(a)

	return out
}

function growUint8(a: Uint8Array): Uint8Array<ArrayBuffer> {
	const out = new Uint8Array(a.length * 2)

	out.set(a)

	return out
}
