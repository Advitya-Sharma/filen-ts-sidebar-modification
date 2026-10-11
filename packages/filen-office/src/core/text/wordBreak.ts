import type { CodePointTable } from "./packed"
import { LOWERCASE_FLAG, PICTOGRAPHIC_FLAG, UPPERCASE_FLAG, WORD_BREAK_CLASS_MASK, WordBreakClass as C } from "./classes"
import { wordBreakTable } from "./unicode"

// Word boundaries by the Unicode Text Segmentation algorithm (UAX #29, default rules WB1-WB999), over
// UTF-16 indices. Positions inside a surrogate pair are never boundaries.

// What a segment contains (wordSegments' flags). A segment without any is space, punctuation, a symbol, an
// ideograph or other unsegmented script.
export const WORD_LETTER = 1
export const WORD_NUMBER = 2
export const WORD_KANA = 4
// A cased letter (General_Category Lu or Lt, and Ll): all-caps words have WORD_UPPER without WORD_LOWER.
export const WORD_UPPER = 8
export const WORD_LOWER = 16

export type WordVisitor = (start: number, end: number, flags: number) => void

// Outside the class range: no element yet (sot), or none ahead (eot).
const NONE = 30

const bit = (...classes: number[]): number => classes.reduce((set, c) => set | (1 << c), 0)

const NEWLINES = bit(C.CR, C.LF, C.Newline)
// WB4: ignored after anything but sot and newlines.
const IGNORED = bit(C.Extend, C.Format, C.ZWJ)
const UNABSORBING = NEWLINES | (1 << NONE)
const AH_LETTER = bit(C.ALetter, C.Hebrew_Letter)
const MID_LETTER = bit(C.MidLetter, C.MidNumLet, C.Single_Quote)
const MID_NUM = bit(C.MidNum, C.MidNumLet, C.Single_Quote)

// Pair outcomes by (collapsed previous class, class). The context rules need the next non-ignored class
// (AHEAD_*) or the one before the previous (BEHIND_*) to be in RULE_SETS[outcome].
const BREAK = 0
const KEEP = 1
const PAIR_RI = 2
const AHEAD_LETTER = 3
const AHEAD_HEBREW = 4
const AHEAD_NUMBER = 5
const BEHIND_LETTER = 6
const BEHIND_HEBREW = 7
const BEHIND_NUMBER = 8

const RULE_SETS = Int32Array.of(0, 0, 0, AH_LETTER, 1 << C.Hebrew_Letter, 1 << C.Numeric, AH_LETTER, 1 << C.Hebrew_Letter, 1 << C.Numeric)

function pairRule(left: number, right: number): number {
	const l = 1 << left
	const r = 1 << right

	if (left === C.CR && right === C.LF) return KEEP // WB3
	if ((l & NEWLINES) !== 0 || (r & NEWLINES) !== 0) return BREAK // WB3a, WB3b
	if ((r & IGNORED) !== 0) return KEEP // WB4
	if ((l & AH_LETTER) !== 0 && (r & AH_LETTER) !== 0) return KEEP // WB5
	if (left === C.Hebrew_Letter && right === C.Single_Quote) return KEEP // WB7a (WB6 failing)
	if ((l & AH_LETTER) !== 0 && (r & MID_LETTER) !== 0) return AHEAD_LETTER // WB6
	if ((l & MID_LETTER) !== 0 && (r & AH_LETTER) !== 0) return BEHIND_LETTER // WB7
	if (left === C.Hebrew_Letter && right === C.Double_Quote) return AHEAD_HEBREW // WB7b
	if (left === C.Double_Quote && right === C.Hebrew_Letter) return BEHIND_HEBREW // WB7c
	if (left === C.Numeric && (r & (AH_LETTER | (1 << C.Numeric))) !== 0) return KEEP // WB8, WB10
	if ((l & AH_LETTER) !== 0 && right === C.Numeric) return KEEP // WB9
	if ((l & MID_NUM) !== 0 && right === C.Numeric) return BEHIND_NUMBER // WB11
	if (left === C.Numeric && (r & MID_NUM) !== 0) return AHEAD_NUMBER // WB12
	if (left === C.Katakana && right === C.Katakana) return KEEP // WB13
	if ((l & (AH_LETTER | bit(C.Numeric, C.Katakana, C.ExtendNumLet))) !== 0 && right === C.ExtendNumLet) return KEEP // WB13a
	if (left === C.ExtendNumLet && (r & (AH_LETTER | bit(C.Numeric, C.Katakana))) !== 0) return KEEP // WB13b
	if (left === C.Regional_Indicator && right === C.Regional_Indicator) return PAIR_RI // WB15, WB16

	return BREAK // WB999
}

const PAIRS = new Uint8Array(32 * 32)
const FLAGS = new Uint8Array(256)

for (let left = 0; left < 32; left++) for (let right = 0; right < 32; right++) PAIRS[(left << 5) | right] = pairRule(left, right)

for (let value = 0; value < 256; value++) {
	const c = value & WORD_BREAK_CLASS_MASK

	FLAGS[value] =
		(c === C.ALetter || c === C.Hebrew_Letter ? WORD_LETTER : 0) |
		(c === C.Numeric ? WORD_NUMBER : 0) |
		(c === C.Katakana ? WORD_KANA : 0) |
		(value & UPPERCASE_FLAG ? WORD_UPPER : 0) |
		(value & LOWERCASE_FLAG ? WORD_LOWER : 0)
}

let asciiValues: Uint8Array | null = null

// Table values for ASCII, which skips the two-stage lookup on Latin text.
function ascii(table: CodePointTable): Uint8Array {
	if (asciiValues !== null) return asciiValues

	asciiValues = new Uint8Array(128)

	for (let cp = 0; cp < 128; cp++) asciiValues[cp] = table.get(cp)

	return asciiValues
}

// Class of the next element in text[i, to) that WB4 does not ignore, or NONE.
function ahead(text: string, i: number, to: number, table: CodePointTable): number {
	while (i < to) {
		let cp = text.charCodeAt(i++)

		if (cp >= 0xd800 && cp < 0xdc00 && i < to) {
			const low = text.charCodeAt(i)

			if (low >= 0xdc00 && low < 0xe000) {
				cp = ((cp - 0xd800) << 10) + low - 0xdc00 + 0x10000
				i++
			}
		}

		const c = table.get(cp) & WORD_BREAK_CLASS_MASK

		if (((IGNORED >> c) & 1) === 0) return c
	}

	return NONE
}

// Walks text[start, to) from a known context: `last` is the collapsed (WB4) class before start, `beforeLast`
// the one before it, `ri` the run of regional indicators ending at `last`. The position start itself is
// never reported. Without a visitor, returns the first boundary after start (to at the end).
function scan(text: string, start: number, to: number, last: number, beforeLast: number, ri: number, visit: WordVisitor | null): number {
	const table = wordBreakTable()
	const asciiTable = ascii(table)
	let prev = NONE
	let segment = start
	let flags = 0

	for (let i = start; i < to;) {
		let cp = text.charCodeAt(i)
		let next = i + 1
		let value: number

		if (cp < 0x80) {
			value = asciiTable[cp] ?? 0
		} else {
			if (cp >= 0xd800 && cp < 0xdc00 && next < to) {
				const low = text.charCodeAt(next)

				if (low >= 0xdc00 && low < 0xe000) {
					cp = ((cp - 0xd800) << 10) + low - 0xdc00 + 0x10000
					next++
				}
			}

			value = table.get(cp)
		}

		const c = value & WORD_BREAK_CLASS_MASK

		if (i !== start) {
			let rule = PAIRS[(last << 5) | c] ?? BREAK

			if (rule === BREAK) {
				// WB3c and WB3d look at the raw previous code point, before WB4.
				if ((prev === C.ZWJ && value >= PICTOGRAPHIC_FLAG) || (prev === C.WSegSpace && c === C.WSegSpace)) rule = KEEP
			} else if (rule === PAIR_RI) {
				rule = ri & 1 ? KEEP : BREAK
			} else if (rule !== KEEP) {
				const context = rule < BEHIND_LETTER ? ahead(text, next, to, table) : beforeLast

				rule = ((RULE_SETS[rule] ?? 0) >> context) & 1 ? KEEP : BREAK
			}

			if (rule === BREAK) {
				if (visit === null) return i

				visit(segment, i, flags)
				segment = i
				flags = 0
			}
		}

		flags |= FLAGS[value] ?? 0

		if (((IGNORED >> c) & 1) === 0 || ((UNABSORBING >> last) & 1) !== 0) {
			if (c === C.Regional_Indicator) ri = last === C.Regional_Indicator ? ri + 1 : 1

			beforeLast = last
			last = c
		}

		prev = c
		i = next
	}

	if (visit !== null && segment < to) visit(segment, to, flags)

	return to
}

// Start of the code point read by classBefore.
let back = 0

function classBefore(text: string, j: number, table: CodePointTable): number {
	let i = j - 1
	let cp = text.charCodeAt(i)

	if (cp >= 0xdc00 && cp < 0xe000 && i > 0) {
		const high = text.charCodeAt(i - 1)

		if (high >= 0xd800 && high < 0xdc00) {
			cp = ((high - 0xd800) << 10) + cp - 0xdc00 + 0x10000
			i--
		}
	}

	back = i

	return table.get(cp) & WORD_BREAK_CLASS_MASK
}

// The collapsed (WB4) class of the element ending at j, with `back` at its start; NONE at the text start.
function collapsedBefore(text: string, j: number, table: CodePointTable): number {
	let first = NONE
	let firstAt = j

	while (j > 0) {
		const c = classBefore(text, j, table)

		if (((IGNORED >> c) & 1) === 0) {
			// Ignorables after a newline are not absorbed: the first of them is the element.
			if (first !== NONE && ((NEWLINES >> c) & 1) !== 0) break

			return c
		}

		first = c
		firstAt = back
		j = back
	}

	back = firstAt

	return first
}

// The first word boundary after index (text.length at the end), in the whole of text.
export function nextWordBoundary(text: string, index: number): number {
	const length = text.length

	if (index >= length) return length

	let start = Math.max(0, index)

	if (start > 0) {
		const unit = text.charCodeAt(start)
		const high = text.charCodeAt(start - 1)

		if (unit >= 0xdc00 && unit < 0xe000 && high >= 0xd800 && high < 0xdc00) start--
	}

	if (start === 0) return scan(text, 0, length, NONE, NONE, 0, null)

	const table = wordBreakTable()
	const last = collapsedBefore(text, start, table)
	const lastAt = back
	const beforeLast = last === NONE ? NONE : collapsedBefore(text, lastAt, table)
	let ri = 0

	// Only the parity of the regional indicator run matters, but finding it means walking the run.
	if (last === C.Regional_Indicator) {
		ri = 1

		for (let at = lastAt; at > 0 && collapsedBefore(text, at, table) === C.Regional_Indicator; at = back) ri++
	}

	return scan(text, start, length, last, beforeLast, ri, null)
}

// Calls visit for each segment between consecutive word boundaries of text[from, to), which is segmented as
// if it were the whole text (sot at from, eot at to). flags is a WORD_* mask of what the segment contains.
export function wordSegments(text: string, visit: WordVisitor, from = 0, to = text.length): void {
	const start = Math.max(0, from)
	const end = Math.min(text.length, to)

	if (start < end) scan(text, start, end, NONE, NONE, 0, visit)
}
