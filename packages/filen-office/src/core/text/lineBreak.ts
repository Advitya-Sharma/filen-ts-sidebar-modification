import { EAST_ASIAN_FLAG, LINE_BREAK_CLASS_MASK, LineBreakClass as C, PICTOGRAPHIC_UNASSIGNED_FLAG } from "./classes"
import { lineBreakTable } from "./unicode"

// Line break opportunities by the Unicode Line Breaking Algorithm (UAX #14, rules LB2-LB31 with LB1's
// default resolution), plus East Asian line-start/line-end prohibition (kinsoku) as a tailoring.

export const BREAK_NONE = 0
export const BREAK_ALLOWED = 1
export const BREAK_MANDATORY = 2
// Not produced here: a hyphenation opportunity (soft hyphen or hyphenator) marked by the layout.
export const BREAK_HYPHEN = 3

const SOT = -1
const EOT = -1

const DOTTED_CIRCLE = 0x25cc

// Class sets: one bit per set in a per-class word.
const SETS = new Uint32Array(64)
let nextSet = 1

function mask(...classes: number[]): number {
	const bit = nextSet

	nextSet <<= 1

	for (const c of classes) SETS[c] = (SETS[c] ?? 0) | bit

	return bit
}

function has(set: number, c: number): boolean {
	return c >= 0 && ((SETS[c] ?? 0) & set) !== 0
}

const HARD = mask(C.BK, C.CR, C.LF, C.NL)
const NO_BASE = mask(C.BK, C.CR, C.LF, C.NL, C.SP, C.ZW)
const QUOTES = mask(C.QU, C.QU_PI, C.QU_PF)
const QUOTE_START_BEFORE = mask(C.BK, C.CR, C.LF, C.NL, C.OP, C.QU, C.QU_PI, C.QU_PF, C.GL, C.SP, C.ZW)
const QUOTE_END_AFTER = mask(C.SP, C.GL, C.WJ, C.CL, C.QU, C.QU_PI, C.QU_PF, C.CP, C.EX, C.IS, C.SY, C.BK, C.CR, C.LF, C.NL, C.ZW)
const CLOSING = mask(C.CL, C.CP, C.EX, C.SY)
const WORD_START_BEFORE = mask(C.BK, C.CR, C.LF, C.NL, C.SP, C.ZW, C.CB, C.GL)
const NON_STARTERS = mask(C.BA, C.HH, C.HY, C.NS)
const HYPHENS = mask(C.HY, C.HH)
const LETTERS = mask(C.AL, C.HL)
const AFFIXES = mask(C.PR, C.PO)
const IDEOGRAPHIC = mask(C.ID, C.EB, C.EM)
const JAMO = mask(C.JL, C.JV, C.JT, C.H2, C.H3)
const AKSARA = mask(C.AK, C.AS)

export interface BreakOptions {
	// CJ (small kana and the like) as non-starters (strict breaking) rather than ideographs (normal).
	readonly strict: boolean
	// Kinsoku lists applied on top of UAX #14, or null.
	readonly kinsoku: Kinsoku | null
}

export const DEFAULT_BREAK_OPTIONS: BreakOptions = { strict: true, kinsoku: null }

// Characters that may not start or end a line. Applied only where the prohibited character or the one
// across the break is East Asian (East_Asian_Width F, W or H), as Word applies kinsoku to Asian text.
export class Kinsoku {
	readonly noStart: ReadonlySet<number>
	readonly noEnd: ReadonlySet<number>

	constructor(noStart: string, noEnd: string) {
		this.noStart = codePoints(noStart)
		this.noEnd = codePoints(noEnd)
	}
}

function codePoints(text: string): Set<number> {
	const set = new Set<number>()

	for (const char of text) set.add(char.codePointAt(0) ?? 0)

	return set
}

// Finds break opportunities in text[start, end). Reusable; the per-code-unit scratch only grows.
export class BreakFinder {
	private props = new Uint8Array(256)

	// Table values per code unit from the last find() or classify() (Line_Break class in the low six bits).
	get classes(): Uint8Array {
		return this.props
	}

	// out[k] is the break before text[start + k] for k in [0, end - start]: BREAK_NONE, BREAK_ALLOWED or
	// BREAK_MANDATORY. out[0] is BREAK_NONE (LB2), out[end - start] BREAK_MANDATORY (LB3); the second unit
	// of a surrogate pair is always BREAK_NONE.
	find(text: string, start: number, end: number, out: Uint8Array, options: BreakOptions = DEFAULT_BREAK_OPTIONS): void {
		const n = end - start

		if (out.length < n + 1) throw new RangeError("The break buffer is too small")

		out.fill(BREAK_NONE, 0, n + 1)

		if (n <= 0) {
			out[0] = BREAK_MANDATORY

			return
		}

		const props = this.classify(text, start, end, options.strict)

		this.resolve(text, start, n, props, out)

		if (options.kinsoku !== null) this.kinsoku(text, start, n, props, out, options.kinsoku)

		out[n] = BREAK_MANDATORY
	}

	// Table values per code unit; the second unit of a surrogate pair repeats the first.
	classify(text: string, start: number, end: number, strict: boolean): Uint8Array {
		const n = end - start

		// Grows by doubling; scratch from a huge paragraph is released once paragraphs are small again.
		if (this.props.length < n + 1 || (this.props.length > 1 << 16 && (n + 1) * 4 < this.props.length))
			this.props = new Uint8Array(Math.max(n + 1, 256, Math.min(this.props.length * 2, 2 * (n + 1))))

		const props = this.props
		const table = lineBreakTable()

		for (let i = 0; i < n; i++) {
			const unit = text.charCodeAt(start + i)
			let cp = unit

			if (unit >= 0xd800 && unit < 0xdc00 && i + 1 < n) {
				const next = text.charCodeAt(start + i + 1)

				if (next >= 0xdc00 && next < 0xe000) cp = ((unit - 0xd800) << 10) + (next - 0xdc00) + 0x10000
			}

			let value = table.get(cp)

			if ((value & LINE_BREAK_CLASS_MASK) === C.CJ) value = (value & ~LINE_BREAK_CLASS_MASK) | (strict ? C.NS : C.ID)

			props[i] = value

			if (cp > 0xffff) props[++i] = value
		}

		return props
	}

	private resolve(text: string, start: number, n: number, props: Uint8Array, out: Uint8Array): void {
		// Effective classes after LB9/LB10 of the previous item and the one before it.
		let prev = SOT
		let prevEastAsian = false
		let prevPictographic = false
		let prevDotted = false
		let prev2 = SOT
		let prev2EastAsian = false
		let prev2Dotted = false
		// The actual class of the previous code point (LB4, LB5, LB8a).
		let last = SOT
		// For SP runs: the item before the first space.
		let beforeSpaces = SOT
		let beforeSpacesQuoteStart = false
		// prev is an initial quotation mark in LB15a's left context.
		let prevQuoteStart = false
		// LB25: 1 after NU (SY | IS)*, 2 after NU (SY | IS)* (CL | CP).
		let numeric = 0
		let regional = 0

		for (let i = 0; i < n; i++) {
			const value = props[i] ?? 0
			const actual = value & LINE_BREAK_CLASS_MASK
			const unit = text.charCodeAt(start + i)
			const cp = unit >= 0xd800 && unit < 0xdc00 ? (text.codePointAt(start + i) ?? unit) : unit
			const step = cp > 0xffff ? 1 : 0
			let action = BREAK_NONE
			let cur = actual
			let eastAsian = (value & EAST_ASIAN_FLAG) !== 0
			let pictographic = (value & PICTOGRAPHIC_UNASSIGNED_FLAG) !== 0
			let dotted = cp === DOTTED_CIRCLE

			const mark = actual === C.CM || actual === C.ZWJ

			if (mark && i > 0 && !has(NO_BASE, prev)) {
				// LB9: the mark joins the previous item and changes nothing else (any earlier rule that
				// applies here also forbids the break).
				last = actual
				i += step

				continue
			}

			if (mark) {
				// LB10.
				cur = C.AL
				eastAsian = false
				pictographic = false
				dotted = false
			}

			if (i > 0) {
				if (last === C.BK || last === C.LF || last === C.NL || (last === C.CR && actual !== C.LF)) {
					action = BREAK_MANDATORY
				} else if (has(HARD, actual) || actual === C.SP || actual === C.ZW || last === C.CR) {
					action = BREAK_NONE
				} else if (prev === C.ZW || (prev === C.SP && beforeSpaces === C.ZW)) {
					action = BREAK_ALLOWED
				} else if (last === C.ZWJ) {
					action = BREAK_NONE
				} else if ((prev === C.AL || prev === C.HL) && (cur === C.AL || cur === C.HL)) {
					// Fast paths for the commonest pairs, where no earlier rule applies: LB28 and LB18.
					action = BREAK_NONE
				} else if (
					prev === C.SP &&
					(cur === C.AL || cur === C.HL || cur === C.ID || cur === C.NU) &&
					beforeSpaces !== C.OP &&
					!beforeSpacesQuoteStart
				) {
					action = BREAK_ALLOWED
				} else {
					action = this.pair(
						text,
						start,
						n,
						props,
						i + step + 1,
						prev,
						prevEastAsian,
						prevPictographic,
						prevDotted,
						prev2,
						prev2EastAsian,
						prev2Dotted,
						beforeSpaces,
						beforeSpacesQuoteStart,
						prevQuoteStart,
						numeric,
						regional,
						cur,
						eastAsian,
						dotted
					)
				}
			}

			out[i] = action

			if (cur === C.SP && prev !== C.SP) {
				beforeSpaces = prev
				beforeSpacesQuoteStart = prevQuoteStart
			}

			const quoteStart = cur === C.QU_PI && (prev === SOT || has(QUOTE_START_BEFORE, prev))

			numeric =
				cur === C.NU
					? 1
					: numeric === 1 && (cur === C.SY || cur === C.IS)
						? 1
						: numeric === 1 && (cur === C.CL || cur === C.CP)
							? 2
							: 0
			regional = cur === C.RI ? regional + 1 : 0
			prev2 = prev
			prev2EastAsian = prevEastAsian
			prev2Dotted = prevDotted
			prev = cur
			prevEastAsian = eastAsian
			prevPictographic = pictographic
			prevDotted = dotted
			prevQuoteStart = quoteStart
			last = actual
			i += step
		}
	}

	// Rules LB11-LB31 for the pair (prev, cur); `after` indexes the code unit after cur's code point.
	private pair(
		text: string,
		start: number,
		n: number,
		props: Uint8Array,
		after: number,
		prev: number,
		prevEastAsian: boolean,
		prevPictographic: boolean,
		prevDotted: boolean,
		prev2: number,
		prev2EastAsian: boolean,
		prev2Dotted: boolean,
		beforeSpaces: number,
		beforeSpacesQuoteStart: boolean,
		prevQuoteStart: boolean,
		numeric: number,
		regional: number,
		cur: number,
		eastAsian: boolean,
		dotted: boolean
	): number {
		const spaced = prev === C.SP

		if (cur === C.WJ || prev === C.WJ || prev === C.GL) return BREAK_NONE
		if (cur === C.GL && prev !== C.SP && prev !== C.HY && prev !== C.HH) return BREAK_NONE
		if (has(CLOSING, cur)) return BREAK_NONE
		if (prev === C.OP || (spaced && beforeSpaces === C.OP)) return BREAK_NONE
		if (prevQuoteStart || (spaced && beforeSpacesQuoteStart)) return BREAK_NONE

		if (cur === C.QU_PF) {
			const next = nextBase(props, after, n)

			if (next === EOT || has(QUOTE_END_AFTER, (props[next] ?? 0) & LINE_BREAK_CLASS_MASK)) return BREAK_NONE
		}

		if (cur === C.IS) {
			if (spaced) {
				const next = nextBase(props, after, n)

				if (next !== EOT && ((props[next] ?? 0) & LINE_BREAK_CLASS_MASK) === C.NU) return BREAK_ALLOWED
			}

			return BREAK_NONE
		}

		if (cur === C.NS && (prev === C.CL || prev === C.CP || (spaced && (beforeSpaces === C.CL || beforeSpaces === C.CP))))
			return BREAK_NONE
		if (cur === C.B2 && (prev === C.B2 || (spaced && beforeSpaces === C.B2))) return BREAK_NONE
		if (spaced) return BREAK_ALLOWED
		if (cur === C.QU || cur === C.QU_PF || prev === C.QU || prev === C.QU_PI) return BREAK_NONE

		if (has(QUOTES, cur)) {
			if (!prevEastAsian) return BREAK_NONE

			const next = nextBase(props, after, n)

			if (next === EOT || ((props[next] ?? 0) & EAST_ASIAN_FLAG) === 0) return BREAK_NONE
		}

		if (has(QUOTES, prev) && (!eastAsian || prev2 === SOT || !prev2EastAsian)) return BREAK_NONE
		if (cur === C.CB || prev === C.CB) return BREAK_ALLOWED
		if (has(HYPHENS, prev) && has(LETTERS, cur) && (prev2 === SOT || has(WORD_START_BEFORE, prev2))) return BREAK_NONE
		if (has(NON_STARTERS, cur) || prev === C.BB) return BREAK_NONE
		if (has(HYPHENS, prev) && prev2 === C.HL && cur !== C.HL) return BREAK_NONE
		if (prev === C.SY && cur === C.HL) return BREAK_NONE
		if (cur === C.IN) return BREAK_NONE
		if ((has(LETTERS, prev) && cur === C.NU) || (prev === C.NU && has(LETTERS, cur))) return BREAK_NONE
		if ((prev === C.PR && has(IDEOGRAPHIC, cur)) || (has(IDEOGRAPHIC, prev) && cur === C.PO)) return BREAK_NONE
		if ((has(AFFIXES, prev) && has(LETTERS, cur)) || (has(LETTERS, prev) && has(AFFIXES, cur))) return BREAK_NONE

		// LB25.
		if (has(AFFIXES, cur) && numeric !== 0) return BREAK_NONE
		if (cur === C.NU && (numeric === 1 || has(AFFIXES, prev) || prev === C.HY || prev === C.IS)) return BREAK_NONE

		if (cur === C.OP && has(AFFIXES, prev)) {
			const next = nextBase(props, after, n)
			const nextClass = next === EOT ? EOT : (props[next] ?? 0) & LINE_BREAK_CLASS_MASK

			if (nextClass === C.NU) return BREAK_NONE

			if (nextClass === C.IS) {
				const third = nextBase(props, next + ((text.codePointAt(start + next) ?? 0) > 0xffff ? 2 : 1), n)

				if (third !== EOT && ((props[third] ?? 0) & LINE_BREAK_CLASS_MASK) === C.NU) return BREAK_NONE
			}
		}

		if (prev === C.JL && (cur === C.JL || cur === C.JV || cur === C.H2 || cur === C.H3)) return BREAK_NONE
		if ((prev === C.JV || prev === C.H2) && (cur === C.JV || cur === C.JT)) return BREAK_NONE
		if ((prev === C.JT || prev === C.H3) && cur === C.JT) return BREAK_NONE
		if ((has(JAMO, prev) && cur === C.PO) || (prev === C.PR && has(JAMO, cur))) return BREAK_NONE
		if (has(LETTERS, prev) && has(LETTERS, cur)) return BREAK_NONE

		// LB28a.
		const prevAksara = has(AKSARA, prev) || prevDotted
		const curAksara = has(AKSARA, cur) || dotted

		if (prev === C.AP && curAksara) return BREAK_NONE
		if (prevAksara && (cur === C.VF || cur === C.VI)) return BREAK_NONE
		if (prev === C.VI && (prev2 === C.AK || prev2 === C.AS || prev2Dotted) && (cur === C.AK || dotted)) return BREAK_NONE

		if (prevAksara && curAksara) {
			const next = nextBase(props, after, n)

			if (next !== EOT && ((props[next] ?? 0) & LINE_BREAK_CLASS_MASK) === C.VF) return BREAK_NONE
		}

		if (prev === C.IS && has(LETTERS, cur)) return BREAK_NONE
		if ((has(LETTERS, prev) || prev === C.NU) && cur === C.OP && !eastAsian) return BREAK_NONE
		if (prev === C.CP && !prevEastAsian && (has(LETTERS, cur) || cur === C.NU)) return BREAK_NONE
		if (prev === C.RI && cur === C.RI && (regional & 1) === 1) return BREAK_NONE
		if (cur === C.EM && (prev === C.EB || prevPictographic)) return BREAK_NONE

		return BREAK_ALLOWED
	}

	// Kinsoku: no break before a line-start-prohibited or after a line-end-prohibited character.
	private kinsoku(text: string, start: number, n: number, props: Uint8Array, out: Uint8Array, lists: Kinsoku): void {
		for (let i = 1; i < n; i++) {
			if (out[i] !== BREAK_ALLOWED) continue

			const next = text.codePointAt(start + i) ?? 0
			const before = previousCodePoint(text, start, i)
			const beforeIndex = i - (before > 0xffff ? 2 : 1)
			const nextEast = ((props[i] ?? 0) & EAST_ASIAN_FLAG) !== 0
			const beforeEast = ((props[beforeIndex] ?? 0) & EAST_ASIAN_FLAG) !== 0

			if ((lists.noStart.has(next) && (nextEast || beforeEast)) || (lists.noEnd.has(before) && (beforeEast || nextEast))) {
				out[i] = BREAK_NONE
			}
		}
	}
}

function previousCodePoint(text: string, start: number, i: number): number {
	const unit = text.charCodeAt(start + i - 1)

	if (unit >= 0xdc00 && unit < 0xe000 && i >= 2) {
		const lead = text.charCodeAt(start + i - 2)

		if (lead >= 0xd800 && lead < 0xdc00) return ((lead - 0xd800) << 10) + (unit - 0xdc00) + 0x10000
	}

	return unit
}

// The next code unit index at or after `from` that is not a combining mark or ZWJ, or EOT.
function nextBase(props: Uint8Array, from: number, n: number): number {
	for (let i = from; i < n; i++) {
		const c = (props[i] ?? 0) & LINE_BREAK_CLASS_MASK

		if (c !== C.CM && c !== C.ZWJ) return i
	}

	return EOT
}
