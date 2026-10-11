import { BidiClass } from "./classes"
import type { CodePointTable } from "./packed"
import { bidiBrackets, bidiClassTable } from "./unicode"

// The Unicode Bidirectional Algorithm (UAX #9) for one paragraph: explicit levels (X1-X10), weak and
// neutral types (W1-W7, N0-N2), implicit levels (I1-I2), and per line L1 and L2. Rules P1, L3 and L4 are
// the caller's: a paragraph is one call, and mirroring and shaping are the renderer's.

export const BIDI_REMOVED = 0xff

export type BidiDirection = "ltr" | "rtl" | "auto"

const { L, R, AL, EN, ES, ET, AN, CS, NSM, BN, B, S, WS, ON, LRE, LRO, RLE, RLO, PDF, LRI, RLI, FSI, PDI } = BidiClass

const MAX_DEPTH = 125
const SHRINK_ABOVE = 1 << 16
const BRACKET_STACK = 63
// reorder() keeps an item's level in the top byte of its index.
const MAX_REORDER_ITEMS = 1 << 24

const bit = (value: number): number => 1 << value

// Removed by X9.
const REMOVED = bit(RLE) | bit(LRE) | bit(RLO) | bit(LRO) | bit(PDF) | bit(BN)
const ISOLATE_INITIATOR = bit(LRI) | bit(RLI) | bit(FSI)
const NEUTRAL_OR_ISOLATE = bit(B) | bit(S) | bit(WS) | bit(ON) | bit(LRI) | bit(RLI) | bit(FSI) | bit(PDI)
// L1 whitespace sequences; X9-removed characters count as whitespace there.
const TRAILING = bit(WS) | bit(FSI) | bit(LRI) | bit(RLI) | bit(PDI) | REMOVED
const RIGHT_TO_LEFT = bit(R) | bit(AL) | bit(AN) | bit(RLE) | bit(RLO) | bit(RLI) | bit(FSI)

let table: CodePointTable | null = null

function classes(): CodePointTable {
	return (table ??= bidiClassTable())
}

function codePointAt(text: string, index: number, end: number): number {
	const high = text.charCodeAt(index)

	if (high < 0xd800 || high > 0xdbff || index + 1 >= end) return high

	const low = text.charCodeAt(index + 1)

	return low >= 0xdc00 && low <= 0xdfff ? ((high - 0xd800) << 10) + (low - 0xdc00) + 0x10000 : high
}

export function hasRightToLeft(text: string, start: number, end: number): boolean {
	const lookup = classes()

	for (let i = start; i < end; i++) {
		const unit = text.charCodeAt(i)

		// Nothing below Hebrew is right-to-left or an explicit formatting character.
		if (unit < 0x590) continue

		const cp = codePointAt(text, i, end)

		if (cp > 0xffff) i++

		if ((RIGHT_TO_LEFT & bit(lookup.get(cp))) !== 0) return true
	}

	return false
}

export class BidiResolver {
	private capacity = 0
	// Per code point of the paragraph.
	private cps = new Int32Array(0)
	private cls = new Uint8Array(0)
	private types = new Uint8Array(0)
	// Embedding levels from X1-X8; sos and eos of every sequence are taken from them.
	private levels = new Uint8Array(0)
	private resolved = new Uint8Array(0)
	private match = new Int32Array(0)
	private runFirst = new Int32Array(0)
	private runLast = new Int32Array(0)
	private runDone = new Uint8Array(0)
	// Per position of the isolating run sequence being resolved.
	private seq = new Int32Array(0)
	private seqTypes = new Uint8Array(0)
	private openings = new Int32Array(0)
	private closing = new Int32Array(0)
	// Isolate matching (BD9), then reused for nothing else.
	private isolates = new Int32Array(0)

	private readonly stackLevel = new Uint8Array(MAX_DEPTH + 2)
	private readonly stackOverride = new Uint8Array(MAX_DEPTH + 2)
	private readonly stackIsolate = new Uint8Array(MAX_DEPTH + 2)
	private readonly bracketId = new Int32Array(BRACKET_STACK)
	private readonly bracketAt = new Int32Array(BRACKET_STACK)

	private grow(count: number): void {
		// Scratch from a huge paragraph is released once paragraphs are back to normal sizes.
		const shrink = this.capacity > SHRINK_ABOVE && count * 4 < this.capacity

		if (count <= this.capacity && !shrink) return

		const capacity = shrink ? Math.max(count * 2, 64) : Math.max(count, this.capacity * 2, 64)

		this.capacity = capacity
		this.cps = new Int32Array(capacity)
		this.cls = new Uint8Array(capacity)
		this.types = new Uint8Array(capacity)
		this.levels = new Uint8Array(capacity)
		this.resolved = new Uint8Array(capacity)
		this.match = new Int32Array(capacity)
		this.runFirst = new Int32Array(capacity)
		this.runLast = new Int32Array(capacity)
		this.runDone = new Uint8Array(capacity)
		this.seq = new Int32Array(capacity)
		this.seqTypes = new Uint8Array(capacity)
		this.openings = new Int32Array(capacity)
		this.closing = new Int32Array(capacity)
		this.isolates = new Int32Array(capacity)
	}

	resolve(text: string, start: number, end: number, direction: BidiDirection, levels: Uint8Array): number {
		const lookup = classes()

		this.grow(end - start)

		const { cps, cls, match, isolates } = this
		let n = 0

		for (let i = start; i < end;) {
			const cp = codePointAt(text, i, end)

			cps[n] = cp
			cls[n] = lookup.get(cp)
			n++
			i += cp > 0xffff ? 2 : 1
		}

		// BD9: matching PDIs. A paragraph separator ends every isolate.
		let depth = 0

		for (let k = 0; k < n; k++) {
			const c = cls[k] ?? 0

			match[k] = -1

			if ((ISOLATE_INITIATOR & bit(c)) !== 0) isolates[depth++] = k
			else if (c === PDI && depth > 0) match[isolates[--depth] ?? 0] = k
			else if (c === B) depth = 0
		}

		const paragraph = direction === "rtl" ? 1 : direction === "auto" && this.firstStrong(0, n) === 1 ? 1 : 0

		this.explicitLevels(n, paragraph)
		this.isolatingRuns(n, paragraph)

		const resolved = this.resolved
		let unit = 0

		for (let k = 0; k < n; k++) {
			const level = (REMOVED & bit(cls[k] ?? 0)) !== 0 ? BIDI_REMOVED : (resolved[k] ?? 0)

			levels[unit++] = level

			if ((cps[k] ?? 0) > 0xffff) levels[unit++] = level
		}

		return paragraph
	}

	// P2-P3 over [from, to): 0 for L, 1 for R or AL, -1 for none. Isolates are skipped whole.
	private firstStrong(from: number, to: number): number {
		const { cls, match } = this

		for (let k = from; k < to; k++) {
			const c = cls[k] ?? 0

			if (c === L) return 0
			if (c === R || c === AL) return 1

			if ((ISOLATE_INITIATOR & bit(c)) !== 0) {
				const pdi = match[k] ?? -1

				if (pdi < 0) return -1

				k = pdi
			}
		}

		return -1
	}

	// X1-X8.
	private explicitLevels(n: number, paragraph: number): void {
		const { cls, types, levels, match, stackLevel, stackOverride, stackIsolate } = this
		let top = 0
		let overflowIsolate = 0
		let overflowEmbedding = 0
		let validIsolate = 0

		stackLevel[0] = paragraph
		stackOverride[0] = 0
		stackIsolate[0] = 0

		for (let k = 0; k < n; k++) {
			const c = cls[k] ?? 0
			let type = c

			switch (c) {
				case RLE:
				case LRE:
				case RLO:
				case LRO: {
					const current = stackLevel[top] ?? 0
					const next = c === RLE || c === RLO ? (current + 1) | 1 : (current + 2) & ~1

					if (next <= MAX_DEPTH && overflowIsolate === 0 && overflowEmbedding === 0) {
						top++
						stackLevel[top] = next
						stackOverride[top] = c === LRO ? 1 : c === RLO ? 2 : 0
						stackIsolate[top] = 0
					} else if (overflowIsolate === 0) {
						overflowEmbedding++
					}

					levels[k] = stackLevel[top] ?? 0
					break
				}
				case RLI:
				case LRI:
				case FSI: {
					const current = stackLevel[top] ?? 0
					const override = stackOverride[top] ?? 0
					const pdi = match[k] ?? -1
					const rtl = c === RLI || (c === FSI && this.firstStrong(k + 1, pdi < 0 ? n : pdi) === 1)
					const next = rtl ? (current + 1) | 1 : (current + 2) & ~1

					levels[k] = current

					if (override !== 0) type = override === 1 ? L : R

					if (next <= MAX_DEPTH && overflowIsolate === 0 && overflowEmbedding === 0) {
						validIsolate++
						top++
						stackLevel[top] = next
						stackOverride[top] = 0
						stackIsolate[top] = 1
					} else {
						overflowIsolate++
					}

					break
				}
				case PDI: {
					if (overflowIsolate > 0) {
						overflowIsolate--
					} else if (validIsolate > 0) {
						overflowEmbedding = 0

						while (stackIsolate[top] === 0) top--

						top--
						validIsolate--
					}

					const override = stackOverride[top] ?? 0

					levels[k] = stackLevel[top] ?? 0

					if (override !== 0) type = override === 1 ? L : R

					break
				}
				case PDF:
					if (overflowIsolate === 0) {
						if (overflowEmbedding > 0) overflowEmbedding--
						else if (stackIsolate[top] === 0 && top > 0) top--
					}

					levels[k] = stackLevel[top] ?? 0
					break
				case B:
					levels[k] = paragraph
					top = 0
					overflowIsolate = 0
					overflowEmbedding = 0
					validIsolate = 0
					break
				case BN:
					levels[k] = stackLevel[top] ?? 0
					break
				default: {
					const override = stackOverride[top] ?? 0

					levels[k] = stackLevel[top] ?? 0

					if (override !== 0) type = override === 1 ? L : R
				}
			}

			types[k] = type
		}
	}

	// X10 (BD13): level runs chained across matched isolates into isolating run sequences.
	private isolatingRuns(n: number, paragraph: number): void {
		const { cls, levels, match, runFirst, runLast, runDone, seq } = this
		let runs = 0
		let previous = -1
		let last = -1

		for (let k = 0; k < n; k++) {
			if ((REMOVED & bit(cls[k] ?? 0)) !== 0) continue

			const level = levels[k] ?? 0

			if (level !== previous) {
				if (runs > 0) runLast[runs - 1] = last

				runFirst[runs++] = k
				previous = level
			}

			last = k
		}

		if (runs > 0) runLast[runs - 1] = last

		runDone.fill(0, 0, runs)

		for (let r = 0; r < runs; r++) {
			if (runDone[r] !== 0) continue

			let length = 0
			let run = r

			for (;;) {
				runDone[run] = 1

				const lastOfRun = runLast[run] ?? 0

				for (let k = runFirst[run] ?? 0; k <= lastOfRun; k++) {
					if ((REMOVED & bit(cls[k] ?? 0)) === 0) seq[length++] = k
				}

				if ((ISOLATE_INITIATOR & bit(cls[lastOfRun] ?? 0)) === 0) break

				const pdi = match[lastOfRun] ?? -1
				const next = pdi < 0 ? -1 : this.runStartingAt(pdi, runs)

				if (next < 0) break

				run = next
			}

			this.resolveSequence(length, n, paragraph)
		}
	}

	private runStartingAt(position: number, runs: number): number {
		const runFirst = this.runFirst
		let low = 0
		let high = runs - 1

		while (low <= high) {
			const middle = (low + high) >>> 1
			const first = runFirst[middle] ?? 0

			if (first === position) return middle
			if (first < position) low = middle + 1
			else high = middle - 1
		}

		return -1
	}

	// W1-W7, N0-N2 and I1-I2 for the sequence in seq[0, m).
	private resolveSequence(m: number, n: number, paragraph: number): void {
		const { cls, types, levels, resolved: out, seq, seqTypes: st } = this
		const first = seq[0] ?? 0
		const lastIndex = seq[m - 1] ?? 0
		const level = levels[first] ?? 0
		let before = paragraph
		let after = paragraph

		for (let k = first - 1; k >= 0; k--) {
			if ((REMOVED & bit(cls[k] ?? 0)) === 0) {
				before = levels[k] ?? 0
				break
			}
		}

		if ((ISOLATE_INITIATOR & bit(cls[lastIndex] ?? 0)) === 0) {
			for (let k = lastIndex + 1; k < n; k++) {
				if ((REMOVED & bit(cls[k] ?? 0)) === 0) {
					after = levels[k] ?? 0
					break
				}
			}
		}

		const sos = Math.max(level, before) & 1 ? R : L
		const eos = Math.max(level, after) & 1 ? R : L
		const embedding = level & 1 ? R : L

		for (let i = 0; i < m; i++) st[i] = types[seq[i] ?? 0] ?? 0

		// W1
		for (let i = 0; i < m; i++) {
			if (st[i] !== NSM) continue

			const prior = i === 0 ? sos : (st[i - 1] ?? 0)

			st[i] = i > 0 && (bit(prior) & (ISOLATE_INITIATOR | bit(PDI))) !== 0 ? ON : prior
		}

		// W2, W3
		let strong: number = sos

		for (let i = 0; i < m; i++) {
			const t = st[i] ?? 0

			if (t === L || t === R) {
				strong = t
			} else if (t === AL) {
				strong = AL
				st[i] = R
			} else if (t === EN && strong === AL) {
				st[i] = AN
			}
		}

		// W4
		for (let i = 1; i + 1 < m; i++) {
			const t = st[i]

			if (t !== ES && t !== CS) continue

			const prior = st[i - 1]
			const next = st[i + 1]

			if (prior === EN && next === EN) st[i] = EN
			else if (t === CS && prior === AN && next === AN) st[i] = AN
		}

		// W5, W6
		for (let i = 0; i < m; i++) {
			const t = st[i]

			if (t === ET) {
				let j = i + 1

				while (j < m && st[j] === ET) j++

				const value = (i > 0 && st[i - 1] === EN) || (j < m && st[j] === EN) ? EN : ON

				st.fill(value, i, j)
				i = j - 1
			} else if (t === ES || t === CS) {
				st[i] = ON
			}
		}

		// W7
		strong = sos

		for (let i = 0; i < m; i++) {
			const t = st[i]

			if (t === L || t === R) strong = t
			else if (t === EN && strong === L) st[i] = L
		}

		this.pairedBrackets(m, sos, embedding)

		// N1, N2
		for (let i = 0; i < m; i++) {
			if ((NEUTRAL_OR_ISOLATE & bit(st[i] ?? 0)) === 0) continue

			let j = i + 1

			while (j < m && (NEUTRAL_OR_ISOLATE & bit(st[j] ?? 0)) !== 0) j++

			const leading = i === 0 ? sos : st[i - 1] === L ? L : R
			const trailing = j === m ? eos : st[j] === L ? L : R

			st.fill(leading === trailing ? leading : embedding, i, j)
			i = j - 1
		}

		// I1, I2
		for (let i = 0; i < m; i++) {
			const t = st[i]
			let resolved = level

			if ((level & 1) === 0) {
				if (t === R) resolved = level + 1
				else if (t === AN || t === EN) resolved = level + 2
			} else if (t === L || t === EN || t === AN) {
				resolved = level + 1
			}

			out[seq[i] ?? 0] = resolved
		}
	}

	// N0 with the bracket pairs of BD16.
	private pairedBrackets(m: number, sos: number, embedding: number): void {
		const { cps, cls, seq, seqTypes: st, openings, closing, bracketId, bracketAt } = this
		const brackets = bidiBrackets()
		let opened = 0
		let depth = 0

		for (let i = 0; i < m; i++) {
			if (st[i] !== ON) continue

			const entry = brackets.get(cps[seq[i] ?? 0] ?? 0)

			if (entry === undefined) continue

			const id = entry >>> 2

			if ((entry & 3) === 1) {
				if (depth === BRACKET_STACK) return

				bracketId[depth] = id
				bracketAt[depth++] = i
				openings[opened++] = i
				closing[i] = -1
			} else {
				for (let d = depth - 1; d >= 0; d--) {
					if (bracketId[d] === id) {
						closing[bracketAt[d] ?? 0] = i
						depth = d
						break
					}
				}
			}
		}

		const opposite = embedding === L ? R : L

		for (let p = 0; p < opened; p++) {
			const open = openings[p] ?? 0
			const close = closing[open] ?? -1

			if (close < 0) continue

			let found = 0

			for (let k = open + 1; k < close; k++) {
				const t = st[k]

				if (t === L ? embedding === L : t === R || t === EN || t === AN ? embedding === R : false) {
					found = 1
					break
				}

				if (t === L || t === R || t === EN || t === AN) found = 2
			}

			if (found === 0) continue

			let direction = embedding

			if (found === 2) {
				let context = sos

				for (let k = open - 1; k >= 0; k--) {
					const t = st[k]

					if (t === L) {
						context = L
						break
					}

					if (t === R || t === EN || t === AN) {
						context = R
						break
					}
				}

				if (context === opposite) direction = opposite
			}

			st[open] = direction
			st[close] = direction

			for (let k = open + 1; k < m && cls[seq[k] ?? 0] === NSM; k++) st[k] = direction
			for (let k = close + 1; k < m && cls[seq[k] ?? 0] === NSM; k++) st[k] = direction
		}
	}
}

export function applyL1(
	text: string,
	textOffset: number,
	levels: Uint8Array,
	lineStart: number,
	lineEnd: number,
	paragraphLevel: number
): void {
	const lookup = classes()
	const end = textOffset + lineEnd
	let trailing = -1

	for (let i = lineStart; i < lineEnd;) {
		const cp = codePointAt(text, textOffset + i, end)
		const units = cp > 0xffff ? 2 : 1
		const c = lookup.get(cp)

		if (c === S || c === B) {
			levels.fill(paragraphLevel, trailing < 0 ? i : trailing, i + units)
			trailing = -1
		} else if ((TRAILING & bit(c)) !== 0) {
			if (trailing < 0) trailing = i
		} else {
			trailing = -1
		}

		i += units
	}

	if (trailing >= 0) levels.fill(paragraphLevel, trailing, lineEnd)
}

export function reorder(levels: Uint8Array, start: number, end: number, out: Int32Array): void {
	const count = end - start

	if (count >= MAX_REORDER_ITEMS) throw new RangeError("Too many items to reorder")

	// Removed items take the level before them; leading ones the first level that follows.
	let previous = 0

	for (let k = start; k < end; k++) {
		const level = levels[k] ?? 0

		if (level !== BIDI_REMOVED) {
			previous = level
			break
		}
	}

	let highest = 0
	let lowestOdd = BIDI_REMOVED

	for (let k = 0; k < count; k++) {
		const raw = levels[start + k] ?? 0
		const level = raw === BIDI_REMOVED ? previous : raw

		previous = level
		out[k] = k | (level << 24)

		if (level > highest) highest = level
		if (level & 1 && level < lowestOdd) lowestOdd = level
	}

	for (let level = highest; level >= lowestOdd; level--) {
		for (let i = 0; i < count;) {
			if ((out[i] ?? 0) >>> 24 < level) {
				i++
				continue
			}

			let j = i + 1

			while (j < count && (out[j] ?? 0) >>> 24 >= level) j++

			for (let a = i, b = j - 1; a < b; a++, b--) {
				const swap = out[a] ?? 0

				out[a] = out[b] ?? 0
				out[b] = swap
			}

			i = j
		}
	}

	for (let k = 0; k < count; k++) out[k] = (out[k] ?? 0) & 0xffffff
}
