// Sets of code points as sorted, disjoint, non-adjacent inclusive ranges [lo0, hi0, lo1, hi1, ...], and the simple
// case folding PCRE2 matches caselessly with in UTF mode.

export const MAX_CODE_POINT = 0x10ffff

export type Ranges = readonly number[]

// Ranges sorted and merged.
export function normalize(ranges: readonly number[]): number[] {
	const pairs: [number, number][] = []

	for (let index = 0; index + 1 < ranges.length; index += 2) pairs.push([ranges[index] ?? 0, ranges[index + 1] ?? 0])

	pairs.sort((a, z) => a[0] - z[0])

	const out: number[] = []

	for (const [lo, hi] of pairs) {
		const last = out.length - 1

		if (last > 0 && lo <= (out[last] ?? 0) + 1) out[last] = Math.max(out[last] ?? 0, hi)
		else out.push(lo, hi)
	}

	return out
}

export function complement(ranges: Ranges): number[] {
	const out: number[] = []
	let next = 0

	for (let index = 0; index + 1 < ranges.length; index += 2) {
		const lo = ranges[index] ?? 0

		if (lo > next) out.push(next, lo - 1)

		next = (ranges[index + 1] ?? 0) + 1
	}

	if (next <= MAX_CODE_POINT) out.push(next, MAX_CODE_POINT)

	return out
}

export function contains(ranges: Ranges, code: number): boolean {
	let low = 0
	let high = (ranges.length >> 1) - 1

	while (low <= high) {
		const middle = (low + high) >> 1

		if (code < (ranges[middle * 2] ?? 0)) high = middle - 1
		else if (code > (ranges[middle * 2 + 1] ?? 0)) low = middle + 1
		else return true
	}

	return false
}

export function size(ranges: Ranges): number {
	let total = 0

	for (let index = 0; index + 1 < ranges.length; index += 2) total += (ranges[index + 1] ?? 0) - (ranges[index] ?? 0) + 1

	return total
}

// The escapes and POSIX classes PCRE2 defines without Unicode properties: ASCII only (\h and \v name their Unicode
// spaces explicitly).
export const DIGIT: Ranges = [0x30, 0x39]
export const WORD: Ranges = [0x30, 0x39, 0x41, 0x5a, 0x5f, 0x5f, 0x61, 0x7a]
export const SPACE: Ranges = [0x09, 0x0d, 0x20, 0x20]
export const HORIZONTAL: Ranges = [
	0x09, 0x09, 0x20, 0x20, 0xa0, 0xa0, 0x1680, 0x1680, 0x180e, 0x180e, 0x2000, 0x200a, 0x202f, 0x202f, 0x205f, 0x205f, 0x3000, 0x3000
]
export const VERTICAL: Ranges = [0x0a, 0x0d, 0x85, 0x85, 0x2028, 0x2029]

export const POSIX: Readonly<Record<string, Ranges>> = {
	alnum: [0x30, 0x39, 0x41, 0x5a, 0x61, 0x7a],
	alpha: [0x41, 0x5a, 0x61, 0x7a],
	ascii: [0x00, 0x7f],
	blank: [0x09, 0x09, 0x20, 0x20],
	cntrl: [0x00, 0x1f, 0x7f, 0x7f],
	digit: DIGIT,
	graph: [0x21, 0x7e],
	lower: [0x61, 0x7a],
	print: [0x20, 0x7e],
	punct: [0x21, 0x2f, 0x3a, 0x40, 0x5b, 0x60, 0x7b, 0x7e],
	space: SPACE,
	upper: [0x41, 0x5a],
	word: WORD,
	xdigit: [0x30, 0x39, 0x41, 0x46, 0x61, 0x66]
}

export function isWord(code: number): boolean {
	return (code >= 0x61 && code <= 0x7a) || (code >= 0x41 && code <= 0x5a) || (code >= 0x30 && code <= 0x39) || code === 0x5f
}

// Simple case foldings (Unicode CaseFolding.txt, status C) that differ from the character's lower case mapping.
const FOLD_EXCEPTIONS = new Map<number, number>([
	[0x00b5, 0x03bc],
	[0x017f, 0x0073],
	[0x0345, 0x03b9],
	[0x03c2, 0x03c3],
	[0x03d0, 0x03b2],
	[0x03d1, 0x03b8],
	[0x03d5, 0x03c6],
	[0x03d6, 0x03c0],
	[0x03f0, 0x03ba],
	[0x03f1, 0x03c1],
	[0x03f5, 0x03b5],
	[0x1c80, 0x0432],
	[0x1c81, 0x0434],
	[0x1c82, 0x043e],
	[0x1c83, 0x0441],
	[0x1c84, 0x0442],
	[0x1c85, 0x0442],
	[0x1c86, 0x044a],
	[0x1c87, 0x0463],
	[0x1c88, 0xa64b],
	[0x1e9b, 0x1e61],
	[0x1fbe, 0x03b9]
])

// Folds of the BMP, filled as met (0: not yet; the fold of 0 is 0 and never looked up).
let bmpFolds: Uint16Array | null = null

function computeFold(code: number): number {
	const exception = FOLD_EXCEPTIONS.get(code)

	if (exception !== undefined) return exception

	const lower = String.fromCodePoint(code).toLowerCase()
	const folded = lower.codePointAt(0) ?? code

	// A lower case of more than one code point (U+0130) has no simple folding.
	return lower.length === (folded > 0xffff ? 2 : 1) ? folded : code
}

// The code point every member of `code`'s caseless class folds to.
export function fold(code: number): number {
	if (code < 0x80) return code >= 0x41 && code <= 0x5a ? code + 0x20 : code
	if (code > 0xffff) return computeFold(code)

	bmpFolds ??= new Uint16Array(0x10000)

	const known = bmpFolds[code] ?? 0

	if (known !== 0) return known

	const folded = computeFold(code)

	bmpFolds[code] = folded

	return folded
}

// The single code point upper case of a code point, or itself.
export function upper(code: number): number {
	const text = String.fromCodePoint(code).toUpperCase()
	const first = text.codePointAt(0) ?? code

	return text.length === (first > 0xffff ? 2 : 1) ? first : code
}

// The most code points of a caseless class enumerated to find their folds; past it, a code point is tested by its
// own case variants only.
const FOLD_ENUMERATION_LIMIT = 0x3000

// A class's members as tested: an ASCII bitmap, its ranges, and for a caseless class the folds of the members it
// names one by one or as a range (PCRE2 leaves the escapes \d \w \s and the POSIX classes out of case closure).
export class CharClass {
	private readonly ascii = new Uint32Array(4)
	private readonly folds: Ranges | null
	private readonly wide: boolean
	private readonly plain: Ranges
	private readonly cased: Ranges
	private readonly negated: boolean
	private readonly caseless: boolean

	constructor(plain: Ranges, cased: Ranges, negated: boolean, caseless: boolean) {
		let folds: number[] | null = null

		this.plain = plain
		this.cased = cased
		this.negated = negated
		this.caseless = caseless

		this.wide = caseless && size(cased) > FOLD_ENUMERATION_LIMIT

		if (caseless && !this.wide) {
			folds = []

			for (let index = 0; index + 1 < cased.length; index += 2) {
				for (let code = cased[index] ?? 0; code <= (cased[index + 1] ?? 0); code++) {
					const folded = fold(code)

					folds.push(folded, folded)
				}
			}

			folds = normalize(folds)
		}

		this.folds = folds

		for (let code = 0; code < 0x80; code++) {
			if (this.test(code)) this.ascii[code >> 5] = (this.ascii[code >> 5] ?? 0) | (1 << (code & 31))
		}
	}

	private test(code: number): boolean {
		return this.inside(code) !== this.negated
	}

	private inside(code: number): boolean {
		if (contains(this.plain, code) || contains(this.cased, code)) return true
		if (!this.caseless) return false

		const folded = fold(code)

		if (this.folds !== null) return contains(this.folds, folded)

		return contains(this.cased, folded) || contains(this.cased, upper(code)) || contains(this.cased, upper(folded))
	}

	matches(code: number): boolean {
		return code < 0x80 ? ((this.ascii[code >> 5] ?? 0) & (1 << (code & 31))) !== 0 : this.test(code)
	}
}
