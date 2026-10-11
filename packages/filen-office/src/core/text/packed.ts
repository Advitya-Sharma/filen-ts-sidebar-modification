// Code point property tables, shipped as run-length strings and expanded on first use into a two-stage
// table for the BMP (64-code-point blocks, identical blocks shared) plus binary-searched runs above it.
//
// Numbers are written big-endian in base 32: leading digits from CONTINUE, the last from FINAL, so most
// values take one character and the strings need no escaping.

const CONTINUE = "abcdefghijklmnopqrstuvwxyz012345"
const FINAL = "ABCDEFGHIJKLMNOPQRSTUVWXYZ@#$%&*"

// Resolved Line_Break classes (see the generator): the value order is the table's encoding.
export const LINE_BREAK_CLASS_NAMES = [
	"BK",
	"CR",
	"LF",
	"NL",
	"CM",
	"ZWJ",
	"WJ",
	"ZW",
	"GL",
	"SP",
	"B2",
	"BA",
	"BB",
	"HY",
	"HH",
	"CB",
	"CL",
	"CP",
	"EX",
	"IN",
	"NS",
	"OP",
	"QU",
	"QU_PI",
	"QU_PF",
	"IS",
	"NU",
	"PO",
	"PR",
	"SY",
	"AK",
	"AL",
	"AP",
	"AS",
	"CJ",
	"EB",
	"EM",
	"H2",
	"H3",
	"HL",
	"ID",
	"JL",
	"JV",
	"JT",
	"RI",
	"VF",
	"VI"
] as const

export const BIDI_CLASS_NAMES = [
	"L",
	"R",
	"AL",
	"EN",
	"ES",
	"ET",
	"AN",
	"CS",
	"NSM",
	"BN",
	"B",
	"S",
	"WS",
	"ON",
	"LRE",
	"LRO",
	"RLE",
	"RLO",
	"PDF",
	"LRI",
	"RLI",
	"FSI",
	"PDI"
] as const

// Word_Break values, Other (the default) first.
export const WORD_BREAK_CLASS_NAMES = [
	"Other",
	"CR",
	"LF",
	"Newline",
	"Extend",
	"ZWJ",
	"Regional_Indicator",
	"Format",
	"Katakana",
	"Hebrew_Letter",
	"ALetter",
	"Single_Quote",
	"Double_Quote",
	"MidNumLet",
	"MidLetter",
	"MidNum",
	"Numeric",
	"ExtendNumLet",
	"WSegSpace"
] as const

function packNumber(value: number): string {
	let out = FINAL[value & 31] ?? ""

	for (let rest = Math.floor(value / 32); rest > 0; rest = Math.floor(rest / 32)) out = (CONTINUE[rest & 31] ?? "") + out

	return out
}

// Build-time: [length, value] runs covering the code space.
export function packRuns(runs: readonly (readonly [number, number])[]): string {
	let out = ""

	for (const [length, value] of runs) out += packNumber(length) + packNumber(value)

	return out
}

// Build-time: a number list, optionally as zigzag-coded deltas from the previous number.
export function packNumbers(values: readonly number[], deltas: boolean): string {
	let out = ""
	let previous = 0

	for (const value of values) {
		const delta = deltas ? value - previous : value

		out += packNumber(deltas ? (delta < 0 ? -2 * delta - 1 : 2 * delta) : delta)
		previous = value
	}

	return out
}

export function unpackNumbers(packed: string, deltas: boolean): number[] {
	const out: number[] = []
	let value = 0
	let previous = 0

	for (let i = 0; i < packed.length; i++) {
		const char = packed[i] ?? ""
		const final = FINAL.indexOf(char)

		if (final < 0) {
			value = value * 32 + CONTINUE.indexOf(char)

			continue
		}

		value = value * 32 + final

		if (deltas) {
			previous += value & 1 ? -(value + 1) / 2 : value / 2
			out.push(previous)
		} else {
			out.push(value)
		}

		value = 0
	}

	return out
}

const BLOCK_SHIFT = 6
const BLOCK = 1 << BLOCK_SHIFT
const BMP = 0x10000

export class CodePointTable {
	private readonly index: Uint16Array
	private readonly blocks: Uint8Array
	private readonly starts: Int32Array
	private readonly values: Uint8Array
	private last = 0

	constructor(packed: string) {
		const numbers = unpackNumbers(packed, false)
		const index = new Uint16Array(BMP >> BLOCK_SHIFT)
		const unique = new Map<string, number>()
		const blocks: number[] = []
		const block = new Uint8Array(BLOCK)
		const starts: number[] = []
		const values: number[] = []
		let cp = 0
		let filled = 0

		for (let i = 0; i < numbers.length; i += 2) {
			const length = numbers[i] ?? 0
			const value = numbers[i + 1] ?? 0
			const end = cp + length

			for (let at = cp; at < end && at < BMP;) {
				const stop = Math.min(end, BMP, (at & ~(BLOCK - 1)) + BLOCK)

				block.fill(value, at & (BLOCK - 1), stop - (at & ~(BLOCK - 1)))
				filled += stop - at
				at = stop

				if (filled === BLOCK) {
					const key = String.fromCharCode(...block)
					let id = unique.get(key)

					if (id === undefined) {
						id = unique.size
						unique.set(key, id)

						for (let k = 0; k < BLOCK; k++) blocks.push(block[k] ?? 0)
					}

					index[(at - 1) >> BLOCK_SHIFT] = id
					filled = 0
				}
			}

			if (end > BMP) {
				starts.push(Math.max(cp, BMP))
				values.push(value)
			}

			cp = end
		}

		this.index = index
		this.blocks = Uint8Array.from(blocks)
		this.starts = Int32Array.from(starts)
		this.values = Uint8Array.from(values)
	}

	get(cp: number): number {
		if (cp < BMP) return this.blocks[((this.index[cp >> BLOCK_SHIFT] ?? 0) << BLOCK_SHIFT) | (cp & (BLOCK - 1))] ?? 0

		const starts = this.starts
		let low = this.last

		// Text in one supplementary script stays in one run; check the last hit before searching.
		if ((starts[low] ?? Infinity) <= cp && cp < (starts[low + 1] ?? Infinity)) return this.values[low] ?? 0

		low = 0

		let high = starts.length - 1

		while (low < high) {
			const middle = (low + high + 1) >>> 1

			if ((starts[middle] ?? 0) <= cp) low = middle
			else high = middle - 1
		}

		this.last = low

		return this.values[low] ?? 0
	}

	// Bytes held after expansion.
	get byteLength(): number {
		return this.index.byteLength + this.blocks.byteLength + this.starts.byteLength + this.values.byteLength
	}
}
