import { CodePointTable, unpackNumbers } from "./packed"
import { BIDI_BRACKET_EQUIVALENTS, BIDI_BRACKET_PAIRS, BIDI_CLASS_RUNS, LINE_BREAK_RUNS } from "./unicode.generated"

// Lazily expanded Unicode property tables: nothing is decoded until text is first segmented.

let lineBreak: CodePointTable | null = null
let bidi: CodePointTable | null = null
let pairs: Map<number, number> | null = null

// Per code point: resolved Line_Break class (bits 0-5, see LineBreakClass), East_Asian_Width F/W/H
// (bit 6), unassigned Extended_Pictographic (bit 7).
export function lineBreakTable(): CodePointTable {
	return (lineBreak ??= new CodePointTable(LINE_BREAK_RUNS))
}

// Per code point: Bidi_Class (see BidiClass).
export function bidiClassTable(): CodePointTable {
	return (bidi ??= new CodePointTable(BIDI_CLASS_RUNS))
}

// Bracket code point -> pair id * 4 + type (1 opening, 2 closing). The id is the opening bracket after
// canonical decomposition, so U+2329 and U+3008 open the same pair, as N0 requires.
export function bidiBrackets(): Map<number, number> {
	if (pairs !== null) return pairs

	const map = new Map<number, number>()
	const list = unpackNumbers(BIDI_BRACKET_PAIRS, true)
	const equivalents = unpackNumbers(BIDI_BRACKET_EQUIVALENTS, true)
	const canonical = new Map<number, number>()

	for (let i = 0; i + 1 < equivalents.length; i += 2) canonical.set(equivalents[i] ?? 0, equivalents[i + 1] ?? 0)

	for (let i = 0; i + 1 < list.length; i += 2) {
		const open = list[i] ?? 0
		const close = list[i + 1] ?? 0
		const id = canonical.get(open) ?? open

		map.set(open, id * 4 + 1)
		map.set(close, id * 4 + 2)
	}

	pairs = map

	return map
}
