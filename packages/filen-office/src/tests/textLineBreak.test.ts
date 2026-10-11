import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { KINSOKU_JAPANESE, KINSOKU_SIMPLIFIED_CHINESE, KINSOKU_TRADITIONAL_CHINESE, kinsokuLists, kinsokuOf } from "../core/text/kinsoku"
import { BREAK_ALLOWED, BREAK_MANDATORY, BREAK_NONE, BreakFinder, type Kinsoku } from "../core/text/lineBreak"
import { FIXTURES, readCached } from "../scripts/unicodeTables"

interface Case {
	readonly line: string
	readonly text: string
	// Expected break before each code unit and at the end (true = opportunity).
	readonly breaks: boolean[]
}

function parse(source: string): Case[] {
	const cases: Case[] = []

	for (const raw of source.split("\n")) {
		const line = raw.replace(/#.*/, "").trim()

		if (line === "") continue

		let text = ""
		const breaks: boolean[] = []

		for (const token of line.split(/\s+/)) {
			if (token === "÷" || token === "×") {
				breaks.push(token === "÷")

				continue
			}

			const char = String.fromCodePoint(Number.parseInt(token, 16))

			// The second unit of a surrogate pair never breaks.
			if (char.length === 2) breaks.push(false)

			text += char
		}

		cases.push({ line, text, breaks })
	}

	return cases
}

function failures(cases: readonly Case[]): string[] {
	const finder = new BreakFinder()
	const out = new Uint8Array(4096)
	const failed: string[] = []

	for (const test of cases) {
		finder.find(test.text, 0, test.text.length, out)

		// The test file marks sot as "×" and eot as "÷".
		const actual = Array.from(out.subarray(0, test.text.length + 1), value => value !== BREAK_NONE)

		if (actual.join() !== test.breaks.join()) failed.push(test.line)
	}

	return failed
}

describe("UAX #14 conformance", () => {
	it("passes the sampled LineBreakTest cases", () => {
		const cases = parse(readFileSync(`${FIXTURES}LineBreakTest.sample.txt`, "utf8"))

		expect(cases.length).toBeGreaterThan(1000)
		expect(failures(cases)).toEqual([])
	})

	it.runIf(process.env["FILEN_OFFICE_UNICODE_FULL"] === "1")("passes the full LineBreakTest", () => {
		const cases = parse(readCached("auxiliary/LineBreakTest.txt"))

		expect(cases.length).toBe(19346)
		expect(failures(cases)).toEqual([])
	})
})

function breaksOf(text: string, kinsoku: Kinsoku | null = null, strict = true): string {
	const out = new Uint8Array(text.length + 1)

	new BreakFinder().find(text, 0, text.length, out, { strict, kinsoku })

	let marked = ""

	for (let i = 0; i <= text.length; i++) {
		if (i > 0 && i < text.length && out[i] === BREAK_ALLOWED) marked += "|"
		if (i > 0 && i < text.length && out[i] === BREAK_MANDATORY) marked += "!"
		if (i < text.length) marked += text[i]
	}

	return marked
}

describe("break opportunities", () => {
	it("breaks Latin text after spaces and hyphens", () => {
		expect(breaksOf("The well-known fox.")).toBe("The |well-|known |fox.")
	})

	it("keeps numbers, prefixes and closing punctuation together", () => {
		expect(breaksOf("costs $12.50 (approx.) today!")).toBe("costs |$12.50 |(approx.) |today!")
	})

	it("handles soft hyphen, no-break hyphen, no-break space, ZWSP and ZWNJ", () => {
		expect(breaksOf("Silben­trennung")).toBe("Silben­|trennung")
		expect(breaksOf("e‑mail")).toBe("e‑mail")
		expect(breaksOf("10 km more")).toBe("10 km |more")
		expect(breaksOf("ab​cd")).toBe("ab​|cd")
		expect(breaksOf("ab‌cd")).toBe("ab‌cd")
	})

	it("forces breaks after line feeds and keeps CR LF together", () => {
		expect(breaksOf("a\r\nb\nc")).toBe("a\r\n!b\n!c")
	})

	it("breaks between ideographs but not before closing punctuation", () => {
		expect(breaksOf("日本語。テスト")).toBe("日|本|語。|テ|ス|ト")
	})

	it("treats small kana as non-starters only in strict mode", () => {
		expect(breaksOf("キャット", null, true)).toBe("キャッ|ト")
		expect(breaksOf("キャット", null, false)).toBe("キ|ャ|ッ|ト")
	})

	it("applies kinsoku to East Asian text only", () => {
		const traditional = kinsokuOf(KINSOKU_TRADITIONAL_CHINESE)

		// U+2022 (bullet) is AL in UAX #14 but may not start a line in Traditional Chinese.
		expect(breaksOf("字\u2022字")).toBe("字|\u2022|字")
		expect(breaksOf("字\u2022字", traditional)).toBe("字\u2022|字")
		// U+00B7 (middle dot) may neither start nor end a line in Simplified Chinese.
		expect(breaksOf("字\u00b7字")).toBe("字\u00b7|字")
		expect(breaksOf("字\u00b7字", kinsokuOf(KINSOKU_SIMPLIFIED_CHINESE))).toBe("字\u00b7字")
		// Latin text keeps its UAX #14 breaks.
		expect(breaksOf("a \u2022 b \u2035 c", traditional)).toBe(breaksOf("a \u2022 b \u2035 c"))
	})

	it("picks the default lists by language", () => {
		expect(kinsokuLists("ja-JP", false)).toBe(KINSOKU_JAPANESE)
		expect(kinsokuLists("ja-JP", true)?.noStart).toContain("\u3063")
		expect(kinsokuLists("zh-TW", false)).toBe(KINSOKU_TRADITIONAL_CHINESE)
		expect(kinsokuLists("zh-CN", false)?.noStart).toContain("\u3001")
		expect(kinsokuLists("ko-KR", false)?.noEnd).toContain("\uffe6")
		expect(kinsokuLists("de-DE", false)).toBeNull()
	})

	it("reuses its scratch without allocating per code point", () => {
		const finder = new BreakFinder()
		const text = "lorem ipsum dolor sit amet ".repeat(40)
		const out = new Uint8Array(text.length + 1)

		finder.find(text, 0, text.length, out)
		expect(out[6]).toBe(BREAK_ALLOWED)
		expect(out[text.length]).toBe(BREAK_MANDATORY)
	})
})
