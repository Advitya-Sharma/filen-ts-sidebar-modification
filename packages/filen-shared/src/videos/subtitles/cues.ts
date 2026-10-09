// Compact cue store shared by every subtitle parser.
//
// Cue text uses WebVTT cue payload syntax: the only markup is `<i>`, `<b>`, `<u>` (and their closers),
// `&amp;`, `&lt;` and `&nbsp;`; every literal `<` or `&` is escaped, so the text is safe for `VTTCue.text` and innerHTML.

export const MAX_SUBTITLE_CHARS = 8 * 1024 * 1024
export const MAX_SUBTITLE_CUES = 100_000
export const MAX_SUBTITLE_CUE_CHARS = 2048

export interface CueStore {
	starts: Float64Array
	ends: Float64Array
	texts: string[]
	// Numpad alignment 1-9 from `{\anN}` / cue settings, 0 = none. null when no cue has one.
	hints: Uint8Array | null
	// Longest cue, bounds the backwards scan of `activeCueIndices`.
	maxDuration: number
}

export const EMPTY_CUES: CueStore = {
	starts: new Float64Array(0),
	ends: new Float64Array(0),
	texts: [],
	hints: null,
	maxDuration: 0
}

export class CueBuilder {
	private starts = new Float64Array(64)
	private ends = new Float64Array(64)
	private hints = new Uint8Array(64)
	private texts: string[] = []
	private anyHint = false

	public get size(): number {
		return this.texts.length
	}

	// false when the cue was rejected (store full or unusable times / empty text).
	public add(start: number, end: number, text: string, hint = 0): boolean {
		const n = this.texts.length

		if (n >= MAX_SUBTITLE_CUES || text.length === 0 || !Number.isFinite(start) || !Number.isFinite(end) || start < 0 || end < start) {
			return false
		}

		if (n === this.starts.length) {
			const cap = n * 2
			const starts = new Float64Array(cap)
			const ends = new Float64Array(cap)
			const hints = new Uint8Array(cap)

			starts.set(this.starts)
			ends.set(this.ends)
			hints.set(this.hints)

			this.starts = starts
			this.ends = ends
			this.hints = hints
		}

		this.starts[n] = start
		this.ends[n] = end
		this.hints[n] = hint
		this.texts.push(text)

		if (hint !== 0) {
			this.anyHint = true
		}

		return true
	}

	public finish(): CueStore {
		const n = this.texts.length

		if (n === 0) {
			return EMPTY_CUES
		}

		let sorted = true

		for (let i = 1; i < n; i++) {
			if ((this.starts[i] as number) < (this.starts[i - 1] as number)) {
				sorted = false

				break
			}
		}

		let starts: Float64Array
		let ends: Float64Array
		let hints: Uint8Array | null
		let texts: string[]

		if (sorted) {
			starts = this.starts.slice(0, n)
			ends = this.ends.slice(0, n)
			hints = this.anyHint ? this.hints.slice(0, n) : null
			texts = this.texts
		} else {
			const order = new Uint32Array(n)

			for (let i = 0; i < n; i++) {
				order[i] = i
			}

			const source = this.starts

			order.sort((a, b) => (source[a] as number) - (source[b] as number) || a - b)

			starts = new Float64Array(n)
			ends = new Float64Array(n)
			hints = this.anyHint ? new Uint8Array(n) : null
			texts = new Array<string>(n)

			for (let i = 0; i < n; i++) {
				const from = order[i] as number

				starts[i] = source[from] as number
				ends[i] = this.ends[from] as number
				texts[i] = this.texts[from] as string

				if (hints !== null) {
					hints[i] = this.hints[from] as number
				}
			}
		}

		let maxDuration = 0

		for (let i = 0; i < n; i++) {
			const d = (ends[i] as number) - (starts[i] as number)

			if (d > maxDuration) {
				maxDuration = d
			}
		}

		return {
			starts,
			ends,
			texts,
			hints,
			maxDuration
		}
	}
}

// First index with starts[i] >= value.
function lowerBound(starts: Float64Array, value: number): number {
	let lo = 0
	let hi = starts.length

	while (lo < hi) {
		const mid = (lo + hi) >>> 1

		if ((starts[mid] as number) < value) {
			lo = mid + 1
		} else {
			hi = mid
		}
	}

	return lo
}

// Fills `out` with the indices of every cue active at `time` (start <= time < end), in start order.
export function activeCueIndices(store: CueStore, time: number, out: number[] = []): number[] {
	out.length = 0

	const n = store.starts.length

	for (let i = lowerBound(store.starts, time - store.maxDuration); i < n && (store.starts[i] as number) <= time; i++) {
		if ((store.ends[i] as number) > time) {
			out.push(i)
		}
	}

	return out
}

// Index of the first cue starting at or after `time`, `size` when none.
export function firstCueAtOrAfter(store: CueStore, time: number): number {
	return lowerBound(store.starts, time)
}

// Shifted copy (positive = later); the text array and hints are shared. Cues pushed before 0 are clamped.
export function applyCueOffset(store: CueStore, offsetSeconds: number): CueStore {
	if (offsetSeconds === 0 || store.texts.length === 0) {
		return store
	}

	const n = store.starts.length
	const starts = new Float64Array(n)
	const ends = new Float64Array(n)
	let maxDuration = 0

	for (let i = 0; i < n; i++) {
		const s = Math.max(0, (store.starts[i] as number) + offsetSeconds)
		const e = Math.max(s, (store.ends[i] as number) + offsetSeconds)

		starts[i] = s
		ends[i] = e

		if (e - s > maxDuration) {
			maxDuration = e - s
		}
	}

	return {
		starts,
		ends,
		texts: store.texts,
		hints: store.hints,
		maxDuration
	}
}

export type CueTextMode = "srt" | "vtt" | "smi" | "ass"

export interface NormalizedCueText {
	text: string
	hint: number
	// Vector drawing (ASS `\p`), not text.
	drawing: boolean
}

const NAMED_ENTITIES: Readonly<Record<string, string>> = {
	amp: "&amp;",
	lt: "&lt;",
	gt: ">",
	nbsp: "&nbsp;",
	quot: "\"",
	apos: "'",
	lrm: "&lrm;",
	rlm: "&rlm;"
}

// indexOf bounded to `max` chars after `from`, so adversarial runs of openers stay linear.
function findWithin(text: string, code: number, from: number, max: number): number {
	const limit = Math.min(text.length, from + max)

	for (let i = from; i < limit; i++) {
		if (text.charCodeAt(i) === code) {
			return i
		}
	}

	return -1
}

function isLetter(code: number): boolean {
	return (code >= 65 && code <= 90) || (code >= 97 && code <= 122)
}

function overrideHint(inner: string): number {
	const at = inner.indexOf("\\an")

	if (at === -1) {
		return 0
	}

	const digit = inner.charCodeAt(at + 3) - 48

	return digit >= 1 && digit <= 9 ? digit : 0
}

function overrideDraws(inner: string): boolean {
	let at = inner.indexOf("\\p")

	while (at !== -1) {
		const digit = inner.charCodeAt(at + 2) - 48

		if (digit >= 1 && digit <= 9) {
			return true
		}

		at = inner.indexOf("\\p", at + 2)
	}

	return false
}

function decodeEntity(name: string): string | null {
	const named = NAMED_ENTITIES[name]

	if (named !== undefined) {
		return named
	}

	if (name.charCodeAt(0) !== 35) {
		return null
	}

	const hex = name.charCodeAt(1) === 120 || name.charCodeAt(1) === 88
	const digits = name.slice(hex ? 2 : 1)

	if (digits.length === 0 || digits.length > 7 || !(hex ? /^[0-9a-fA-F]+$/ : /^[0-9]+$/).test(digits)) {
		return null
	}

	const code = parseInt(digits, hex ? 16 : 10)

	if (code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) {
		return null
	}

	if (code === 38) {
		return "&amp;"
	}

	if (code === 60) {
		return "&lt;"
	}

	return String.fromCodePoint(code)
}

const MAX_TAG_CHARS = 1024
const MAX_OVERRIDE_CHARS = 2048

// Single pass; unknown tags are dropped, text is escaped into cue payload syntax. `{\anN}` becomes the hint.
export function normalizeCueText(raw: string, mode: CueTextMode): NormalizedCueText {
	const ass = mode === "ass"
	const vtt = mode === "vtt"
	const len = Math.min(raw.length, MAX_SUBTITLE_CUE_CHARS * 2)

	if (raw.length > len) {
		raw = raw.slice(0, len)
	}

	// Closers are searched with a forward-only cache, so unmatched openers cost one scan in total.
	let nextGt = -2
	let nextBrace = -2
	let out = ""
	let hint = 0
	let drawing = false
	let i = 0

	while (i < len && out.length < MAX_SUBTITLE_CUE_CHARS) {
		const code = raw.charCodeAt(i)

		if (code === 60 && !ass) {
			const next = raw.charCodeAt(i + 1)
			const closes = next === 47
			const nameStart = closes ? i + 2 : i + 1
			const first = raw.charCodeAt(nameStart)
			if (nextGt !== -1 && nextGt < i + 1) {
				nextGt = raw.indexOf(">", i + 1)
			}

			const end = nextGt !== -1 && nextGt - i <= MAX_TAG_CHARS ? nextGt : -1

			if (end !== -1 && (isLetter(first) || (vtt && first >= 48 && first <= 57))) {
				let nameEnd = nameStart

				while (nameEnd < end && isLetter(raw.charCodeAt(nameEnd))) {
					nameEnd++
				}

				const name = raw.slice(nameStart, nameEnd).toLowerCase()

				if ((name === "i" || name === "b" || name === "u") && nameEnd === nameStart + 1) {
					out += closes ? `</${name}>` : `<${name}>`
				} else if (name === "br" && !closes) {
					out += "\n"
				}

				i = end + 1

				continue
			}

			out += "&lt;"
			i++

			continue
		}

		if (code === 123) {
			if (nextBrace !== -1 && nextBrace < i + 1) {
				nextBrace = raw.indexOf("}", i + 1)
			}

			const end = nextBrace !== -1 && nextBrace - i <= MAX_OVERRIDE_CHARS ? nextBrace : -1

			if (end !== -1 && (ass || raw.charCodeAt(i + 1) === 92)) {
				const inner = raw.slice(i + 1, end)
				const h = overrideHint(inner)

				if (h !== 0 && hint === 0) {
					hint = h
				}

				if (ass && overrideDraws(inner)) {
					drawing = true
				}

				i = end + 1

				continue
			}
		}

		if (code === 92 && ass) {
			const next = raw.charCodeAt(i + 1)

			if (next === 78 || next === 110) {
				out += "\n"
				i += 2

				continue
			}

			if (next === 104) {
				out += "&nbsp;"
				i += 2

				continue
			}
		}

		if (code === 38) {
			if (!ass) {
				const end = findWithin(raw, 59, i + 1, 10)

				if (end !== -1) {
					const decoded = decodeEntity(raw.slice(i + 1, end))

					if (decoded !== null) {
						out += decoded
						i = end + 1

						continue
					}
				}
			}

			out += "&amp;"
			i++

			continue
		}

		if (code === 13) {
			i++

			continue
		}

		if (code === 60) {
			out += "&lt;"
		} else {
			out += raw[i]
		}

		i++
	}

	const text = trimLines(out)

	return {
		text: BLANK_TEXT.test(text) ? "" : text,
		hint,
		drawing
	}
}

const BLANK_TEXT = /^(?:\s|&nbsp;)*$/

// Trims the whole text and each line, dropping blank lines.
function trimLines(value: string): string {
	if (value.indexOf("\n") === -1) {
		return value.trim()
	}

	let out = ""

	for (const line of value.split("\n")) {
		const t = line.trim()

		if (t.length > 0) {
			out = out.length === 0 ? t : `${out}\n${t}`
		}
	}

	return out
}

// Parses `[H+:]M[M]:SS[.,]fraction` (also `SS` alone) at `from`. Returns seconds and the index after it, or null.
export function parseSubtitleTimestamp(value: string, from = 0): { seconds: number; end: number } | null {
	const len = value.length
	let i = from
	let total = 0
	let groups = 0

	for (;;) {
		let digits = 0
		let n = 0

		while (i < len) {
			const d = value.charCodeAt(i) - 48

			if (d < 0 || d > 9) {
				break
			}

			if (digits < 10) {
				n = n * 10 + d
			}

			digits++
			i++
		}

		if (digits === 0) {
			return null
		}

		total = total * 60 + n
		groups++

		if (value.charCodeAt(i) === 58 && groups < 3) {
			i++

			continue
		}

		break
	}

	let fraction = 0

	const sep = value.charCodeAt(i)

	if (sep === 46 || sep === 44) {
		let scale = 0.1
		let digits = 0

		i++

		while (i < len) {
			const d = value.charCodeAt(i) - 48

			if (d < 0 || d > 9) {
				break
			}

			if (digits < 9) {
				fraction += d * scale
				scale /= 10
			}

			digits++
			i++
		}

		if (digits === 0) {
			return null
		}
	}

	if (groups < 2) {
		return null
	}

	return {
		seconds: total + fraction,
		end: i
	}
}

// Forward-only line iterator over `\n`, `\r\n` and lone `\r`; no intermediate array.
export class SubtitleLineReader {
	public pos = 0

	private readonly text: string

	public constructor(text: string) {
		this.text = text
	}

	public next(): string | null {
		const text = this.text
		const len = text.length

		if (this.pos >= len) {
			return null
		}

		const start = this.pos
		let i = start

		while (i < len) {
			const c = text.charCodeAt(i)

			if (c === 10 || c === 13) {
				break
			}

			i++
		}

		const line = text.slice(start, i)

		if (i < len && text.charCodeAt(i) === 13 && text.charCodeAt(i + 1) === 10) {
			i++
		}

		this.pos = i + 1

		return line
	}

	public peek(): string | null {
		const save = this.pos
		const line = this.next()

		this.pos = save

		return line
	}
}
