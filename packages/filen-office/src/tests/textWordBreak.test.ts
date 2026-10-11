import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { WORD_KANA, WORD_LETTER, WORD_LOWER, WORD_NUMBER, WORD_UPPER, nextWordBoundary, wordSegments } from "../core/text/index"
import { wordBreakTable } from "../core/text/unicode"
import { FIXTURES, readCached } from "../scripts/unicodeTables"

interface Case {
	readonly line: string
	readonly text: string
	// UTF-16 positions of the expected boundaries, ascending, including 0 and the end.
	readonly boundaries: number[]
}

function parse(source: string): Case[] {
	const cases: Case[] = []

	for (const raw of source.split("\n")) {
		const line = raw.replace(/#.*/, "").trim()

		if (line === "") continue

		let text = ""
		const boundaries: number[] = []

		for (const token of line.split(/\s+/)) {
			if (token === "÷") boundaries.push(text.length)
			else if (token !== "×") text += String.fromCodePoint(Number.parseInt(token, 16))
		}

		cases.push({ line, text, boundaries })
	}

	return cases
}

function segmentBoundaries(text: string, from = 0, to = text.length): number[] {
	const out: number[] = []

	wordSegments(
		text,
		(start, end) => {
			if (out.length === 0) out.push(start)

			out.push(end)
		},
		from,
		to
	)

	return out
}

function failures(cases: readonly Case[]): string[] {
	const failed: string[] = []

	for (const test of cases) {
		const expected = test.boundaries.join()

		if (segmentBoundaries(test.text).join() !== expected) {
			failed.push(`segments: ${test.line}`)

			continue
		}

		// nextWordBoundary from every position, which rebuilds the context before it.
		for (let i = 0, k = 0; i < test.text.length; i++) {
			while ((test.boundaries[k] ?? Infinity) <= i) k++

			if (nextWordBoundary(test.text, i) !== test.boundaries[k]) {
				failed.push(`next from ${i}: ${test.line}`)

				break
			}
		}
	}

	return failed
}

describe("UAX #29 word boundary conformance", () => {
	it("passes the sampled WordBreakTest cases", () => {
		const cases = parse(readFileSync(`${FIXTURES}WordBreakTest.sample.txt`, "utf8"))

		expect(cases.length).toBeGreaterThan(900)
		expect(failures(cases)).toEqual([])
	})

	it.runIf(process.env["FILEN_OFFICE_UNICODE_FULL"] === "1")("passes the full WordBreakTest", () => {
		const cases = parse(readCached("auxiliary/WordBreakTest.txt"))

		expect(cases.length).toBe(1944)
		expect(failures(cases)).toEqual([])
	})
})

function segments(text: string, from?: number, to?: number): string[] {
	const out: string[] = []

	wordSegments(text, (start, end) => out.push(text.slice(start, end)), from, to)

	return out
}

function words(text: string): [string, number][] {
	const out: [string, number][] = []

	wordSegments(text, (start, end, flags) => {
		if (flags !== 0) out.push([text.slice(start, end), flags])
	})

	return out
}

describe("word segments for spellchecking", () => {
	it("keeps apostrophes inside words", () => {
		expect(segments("can't stop")).toEqual(["can't", " ", "stop"])
		expect(segments("can’t")).toEqual(["can’t"])
		expect(segments("dogs' toys")).toEqual(["dogs", "'", " ", "toys"])
	})

	it("splits abbreviations and hyphenated words", () => {
		expect(segments("e.g.")).toEqual(["e.g", "."])
		expect(segments("well-known")).toEqual(["well", "-", "known"])
		expect(segments("example.com")).toEqual(["example.com"])
	})

	it("keeps numbers together", () => {
		expect(segments("3.14")).toEqual(["3.14"])
		expect(segments("1,234.5 or 3a")).toEqual(["1,234.5", " ", "or", " ", "3a"])
		expect(words("3.14 A3 snake_case")).toEqual([
			["3.14", WORD_NUMBER],
			["A3", WORD_LETTER | WORD_NUMBER | WORD_UPPER],
			["snake_case", WORD_LETTER | WORD_LOWER]
		])
	})

	it("reports letters and case", () => {
		expect(words("Hello, NASA and ÉCOLE élève!")).toEqual([
			["Hello", WORD_LETTER | WORD_UPPER | WORD_LOWER],
			["NASA", WORD_LETTER | WORD_UPPER],
			["and", WORD_LETTER | WORD_LOWER],
			["ÉCOLE", WORD_LETTER | WORD_UPPER],
			["élève", WORD_LETTER | WORD_LOWER]
		])
		expect(words("  ,;!? — … ")).toEqual([])
	})

	it("keeps combining marks with their base", () => {
		const word = "café naïve"

		expect(segments(word)).toEqual(["café", " ", "naïve"])
		expect(segments("́a")).toEqual(["́", "a"])
		expect(segments("a\ńb")).toEqual(["a", "\n", "́", "b"])
	})

	it("keeps emoji ZWJ sequences and flags whole", () => {
		const family = "\u{1f468}‍\u{1f469}‍\u{1f467}"
		const flags = "\u{1f1e9}\u{1f1ea}\u{1f1eb}\u{1f1f7}\u{1f1ee}"

		expect(segments(`${family} ok`)).toEqual([family, " ", "ok"])
		expect(segments(flags)).toEqual(["\u{1f1e9}\u{1f1ea}", "\u{1f1eb}\u{1f1f7}", "\u{1f1ee}"])
		expect(words(family)).toEqual([])
	})

	it("handles Hebrew geresh and gershayim", () => {
		expect(segments("צה״ל")).toEqual(["צה״ל"])
		expect(segments('צה"ל')).toEqual(['צה"ל'])
		expect(segments("ג'ירפה")).toEqual(["ג'ירפה"])
		expect(segments("ג' ")).toEqual(["ג'", " "])
		// Geresh is a letter, so it never splits a word.
		expect(segments("ג\u05f3ירפה \u05f3")).toEqual(["ג\u05f3ירפה", " ", "\u05f3"])
		expect(words("שלום")).toEqual([["שלום", WORD_LETTER]])
	})

	it("reads surrogate pairs as one code point", () => {
		const bold = "\u{1d400}\u{1d401}"

		expect(segments(`${bold} x`)).toEqual([bold, " ", "x"])
		expect(words(bold)).toEqual([[bold, WORD_LETTER | WORD_UPPER]])
		// Ideographs, including CJK Extension B, are a segment each.
		expect(segments("\u{20000}\u{20001}漢")).toEqual(["\u{20000}", "\u{20001}", "漢"])
		expect(nextWordBoundary("\u{20000}\u{20001}", 1)).toBe(2)
		// A lone surrogate is a code point of its own.
		expect(segments("a\ud800b")).toEqual(["a", "\ud800", "b"])
	})

	it("keeps katakana runs and splits hiragana and ideographs", () => {
		expect(segments("カタカナとひら")).toEqual(["カタカナ", "と", "ひ", "ら"])
		expect(words("カタカナー")).toEqual([["カタカナー", WORD_KANA]])
	})

	it("segments only within from and to", () => {
		const text = "one two-three"

		expect(segments(text, 4, 11)).toEqual(["two", "-", "thr"])
		expect(segments(text, 2, 2)).toEqual([])
		expect(segments(text, -5, 99)).toEqual(segments(text))
		// "a.b" is one word in the whole text, but "a." alone ends before eot.
		expect(segments("a.b", 0, 2)).toEqual(["a", "."])
	})

	it("finds the next boundary in the whole text", () => {
		const text = "can't stop, 3.14"
		const found: number[] = []

		for (let at = 0; at < text.length; at = nextWordBoundary(text, at)) found.push(at)

		expect(found).toEqual([0, 5, 6, 10, 11, 12])
		expect(nextWordBoundary(text, 2)).toBe(5)
		expect(nextWordBoundary(text, text.length)).toBe(text.length)
		expect(nextWordBoundary(text, -1)).toBe(5)
		expect(nextWordBoundary("", 0)).toBe(0)
	})

	it("expands to a small table", () => {
		expect(wordBreakTable().byteLength).toBeLessThan(32 * 1024)
	})
})
