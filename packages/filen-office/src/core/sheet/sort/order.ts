// Excel's sort order ("Sort data in a workbook in the browser", support.microsoft.com, "Default sort orders"): in an
// ascending sort numbers (dates are numbers) come first, from the smallest negative to the largest positive, then
// text, then logical values (FALSE before TRUE), then errors (all equal); blank cells are last whatever the
// direction, and a descending sort reverses everything else. Text sorts left to right, character by character, in
// the order 0-9, space ! " # $ % & ( ) * , . / : ; ? @ [ \ ] ^ _ ` { | } ~ + < = >, then the letters; apostrophes
// and hyphens are ignored, except that of two texts equal but for a hyphen the one with the hyphen sorts last. A
// sort is case-insensitive unless asked: then a lowercase letter sorts before its uppercase one (a A b B ...).
//
// The page is silent on letters outside A-Z: accented Latin letters sort with their base letter, after it when the
// texts are otherwise equal (as Windows sorts them); letters of other scripts after Z, by code point; other
// symbols after the listed ones, by code point.

export const SortType = {
	NUMBER: 0,
	TEXT: 1,
	BOOLEAN: 2,
	ERROR: 3,
	BLANK: 4
} as const

export type SortType = (typeof SortType)[keyof typeof SortType]

// One cell as a sort reads it: numbers and booleans by `number` (0 or 1), text by its key (textSortKey).
export interface SortValue {
	readonly type: SortType
	readonly number: number
	readonly key: string
}

export const BLANK_SORT_VALUE: SortValue = { type: SortType.BLANK, number: 0, key: "" }

const SYMBOLS = ' !"#$%&()*,./:;?@[\\]^_`{|}~+<=>'
// Weights: the separator 1, digits 2-11, the listed symbols next, other symbols, then letters. Every weight is a
// single UTF-16 unit below the surrogates, so the keys compare as plain strings.
const SEPARATOR = "\u0001"
const DIGIT_BASE = 2
const SYMBOL_BASE = DIGIT_BASE + 10
const OTHER_SYMBOL_BASE = SYMBOL_BASE + SYMBOLS.length
// Other symbols by code point, folded into 0x0800 weights (a collision only ties them).
const OTHER_SYMBOL_SPAN = 0x0800
const LETTER_BASE = OTHER_SYMBOL_BASE + OTHER_SYMBOL_SPAN
// Other scripts' letters by code point, folded into what is left below the surrogates.
const OTHER_LETTER_BASE = LETTER_BASE + 26
const OTHER_LETTER_SPAN = 0xd7ff - OTHER_LETTER_BASE

const ASCII = new Uint16Array(128)

for (let code = 0; code < 128; code++) {
	if (code >= 48 && code <= 57) ASCII[code] = DIGIT_BASE + code - 48
	else if (code >= 65 && code <= 90) ASCII[code] = LETTER_BASE + code - 65
	else if (code >= 97 && code <= 122) ASCII[code] = LETTER_BASE + code - 97
	else {
		const at = SYMBOLS.indexOf(String.fromCharCode(code))

		ASCII[code] = at >= 0 ? SYMBOL_BASE + at : OTHER_SYMBOL_BASE + code
	}
}

// Accented Latin letters (Latin-1 Supplement and Latin Extended-A) by base letter: "a" for à, with the accent's rank.
// (× and ÷ are symbols: "*" and "/" mark them.)
const LATIN_BASES = [
	// U+00C0-U+00FF
	"AAAAAAA",
	"C",
	"EEEE",
	"IIII",
	"DN",
	"OOOOO*O",
	"UUUU",
	"YTs",
	"aaaaaaa",
	"c",
	"eeee",
	"iiii",
	"dn",
	"ooooo/o",
	"uuuu",
	"yty",
	// U+0100-U+017F
	"AaAaAa",
	"CcCcCcCc",
	"DdDd",
	"EeEeEeEeEe",
	"GgGgGgGg",
	"HhHh",
	"IiIiIiIiIiIi",
	"Jj",
	"Kkk",
	"LlLlLlLlLl",
	"NnNnNnn",
	"Nn",
	"OoOoOoOo",
	"RrRrRr",
	"SsSsSsSs",
	"TtTtTt",
	"UuUuUuUuUuUu",
	"Ww",
	"YyY",
	"ZzZzZz",
	"s"
].join("")

interface Folded {
	// The primary weight, 0 for a character the order ignores (apostrophe, hyphen).
	readonly primary: number
	// 0 unaccented, else an accent's rank.
	readonly accent: number
	readonly upper: boolean
}

function isUpper(char: string): boolean {
	return char !== char.toLowerCase() && char === char.toUpperCase()
}

function fold(code: number, char: string): Folded {
	if (code < 128) {
		if (code === 39 || code === 45) return { primary: 0, accent: 0, upper: false }

		return { primary: ASCII[code] ?? 0, accent: 0, upper: code >= 65 && code <= 90 }
	}

	if (code >= 0xc0 && code <= 0x17f) {
		const base = LATIN_BASES.charCodeAt(code - 0xc0)

		// × and ÷ are symbols.
		if (base !== 42 && base !== 47) return { primary: ASCII[base] ?? 0, accent: 1 + ((code - 0xc0) % 31), upper: isUpper(char) }
	}

	// The right single quotation mark and the hyphens of General Punctuation read as the ASCII ones.
	if (code === 0x2019 || code === 0x2010 || code === 0x2011) return { primary: 0, accent: 0, upper: false }

	// A character with no case is a letter when it is outside the punctuation and symbol blocks.
	const cased = char.toLowerCase() !== char.toUpperCase()
	const symbol =
		!cased &&
		((code >= 0x80 && code <= 0xbf) ||
			code === 0xd7 ||
			code === 0xf7 ||
			(code >= 0x2000 && code <= 0x2bff) ||
			(code >= 0x3000 && code <= 0x303f) ||
			(code >= 0xfe30 && code <= 0xfe4f) ||
			(code >= 0xff00 && code <= 0xff0f))

	if (symbol) return { primary: OTHER_SYMBOL_BASE + 128 + ((code - 128) % (OTHER_SYMBOL_SPAN - 128)), accent: 0, upper: false }

	const lower = cased ? (char.toLowerCase().codePointAt(0) ?? code) : code

	return { primary: OTHER_LETTER_BASE + (lower % OTHER_LETTER_SPAN), accent: 0, upper: cased && isUpper(char) }
}

// A text's sort key: compared as plain strings, two keys order as Excel orders their texts (case-insensitive unless
// `caseSensitive`). Primary weights, then accents, then (when case counts) case, then the ignored characters.
export function textSortKey(text: string, caseSensitive: boolean): string {
	let primary = ""
	let accents = ""
	let cases = ""
	let ignored = 0
	let accented = false

	for (let i = 0; i < text.length; i++) {
		let code = text.charCodeAt(i)
		let char = text[i] as string

		// A surrogate pair is one character.
		if (code >= 0xd800 && code <= 0xdbff && i + 1 < text.length) {
			const low = text.charCodeAt(i + 1)

			if (low >= 0xdc00 && low <= 0xdfff) {
				code = 0x10000 + ((code - 0xd800) << 10) + (low - 0xdc00)
				char = text.slice(i, i + 2)
				i++
			}
		}

		const folded = fold(code, char)

		if (folded.primary === 0) {
			ignored++

			continue
		}

		primary += String.fromCharCode(folded.primary)
		accents += String.fromCharCode(1 + folded.accent)
		cases += folded.upper ? "\u0002" : "\u0001"
		accented ||= folded.accent !== 0
	}

	// Each part ends in the separator, so a text that is a prefix of another sorts first.
	return (
		primary +
		SEPARATOR +
		(accented ? accents : "") +
		SEPARATOR +
		(caseSensitive ? cases : "") +
		SEPARATOR +
		String.fromCharCode(1 + Math.min(ignored, 0xd000))
	)
}

// -1, 0 or 1 for a before, equal to, or after b, ascending; blanks last in either direction.
export function compareSortValues(a: SortValue, b: SortValue, descending: boolean): number {
	if (a.type === SortType.BLANK || b.type === SortType.BLANK) return a.type === b.type ? 0 : a.type === SortType.BLANK ? 1 : -1

	let order = 0

	if (a.type !== b.type) order = a.type < b.type ? -1 : 1
	else if (a.type === SortType.TEXT) order = a.key < b.key ? -1 : a.key > b.key ? 1 : 0
	else if (a.type !== SortType.ERROR) order = a.number < b.number ? -1 : a.number > b.number ? 1 : 0

	return descending ? -order : order
}

// The order of `count` items by their keys (the first key first), stable: items equal on every key keep their order.
// `value(key, item)` reads one key of one item.
export function sortOrder(
	count: number,
	keys: readonly { readonly descending: boolean }[],
	value: (key: number, item: number) => SortValue
): Int32Array {
	const values: SortValue[][] = keys.map((_, key) => {
		const column: SortValue[] = new Array<SortValue>(count)

		for (let item = 0; item < count; item++) column[item] = value(key, item)

		return column
	})
	const order: number[] = new Array<number>(count)

	for (let item = 0; item < count; item++) order[item] = item

	order.sort((a, b) => {
		for (let key = 0; key < keys.length; key++) {
			const column = values[key] as SortValue[]
			const compared = compareSortValues(
				column[a] as SortValue,
				column[b] as SortValue,
				(keys[key] as { descending: boolean }).descending
			)

			if (compared !== 0) return compared
		}

		return a - b
	})

	return Int32Array.from(order)
}
