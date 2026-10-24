import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { BIDI_REMOVED, BidiResolver, applyL1, hasRightToLeft, reorder, type BidiDirection } from "../core/text/bidi"
import { readCached } from "../scripts/unicodeTables"
import { budget } from "./timing"

const DIRECTIONS: readonly BidiDirection[] = ["ltr", "rtl", "auto"]

// Runs every case of a BidiCharacterTest.txt body; returns the failures (at most `keep` described).
function conformance(source: string, keep = 10): { cases: number; failures: number; first: string[] } {
	const resolver = new BidiResolver()
	const first: string[] = []
	let cases = 0
	let failures = 0
	let levels = new Uint8Array(64)
	let order = new Int32Array(64)
	let compact = new Uint8Array(64)
	let compactIndex = new Int32Array(64)

	for (const line of source.split("\n")) {
		if (line === "" || line.startsWith("#")) continue

		const [points = "", direction = "", paragraph = "", expectedLevels = "", expectedOrder = ""] = line.split(";")
		const cps = points
			.trim()
			.split(" ")
			.map(hex => Number.parseInt(hex, 16))
		const text = String.fromCodePoint(...cps)

		cases++

		if (levels.length < text.length) {
			levels = new Uint8Array(text.length * 2)
			order = new Int32Array(text.length * 2)
			compact = new Uint8Array(text.length * 2)
			compactIndex = new Int32Array(text.length * 2)
		}

		const level = resolver.resolve(text, 0, text.length, DIRECTIONS[Number(direction)] ?? "ltr", levels)
		const offsets: number[] = []
		const removed: boolean[] = []

		for (let unit = 0, k = 0; k < cps.length; k++) {
			offsets.push(unit)
			removed.push(levels[unit] === BIDI_REMOVED)
			unit += (cps[k] ?? 0) > 0xffff ? 2 : 1
		}

		applyL1(text, 0, levels, 0, text.length, level)

		const actualLevels = cps.map((_, k) => (removed[k] ? "x" : String(levels[offsets[k] ?? 0]))).join(" ")
		let count = 0

		for (let k = 0; k < cps.length; k++) {
			if (removed[k]) continue

			compact[count] = levels[offsets[k] ?? 0] ?? 0
			compactIndex[count++] = k
		}

		reorder(compact, 0, count, order)

		const actualOrder = Array.from(order.subarray(0, count), index => compactIndex[index]).join(" ")

		if (String(level) !== paragraph || actualLevels !== expectedLevels.trim() || actualOrder !== expectedOrder.trim()) {
			failures++

			if (first.length < keep) first.push(`${line}\n  got ${level};${actualLevels};${actualOrder}`)
		}
	}

	return { cases, failures, first }
}

// One character of each Bidi_Class, for the class sequences of BidiTest.txt.
const SAMPLES: Record<string, number> = {
	L: 0x61,
	R: 0x5d0,
	AL: 0x627,
	EN: 0x31,
	ES: 0x2b,
	ET: 0x24,
	AN: 0x660,
	CS: 0x2c,
	NSM: 0x300,
	BN: 0xad,
	B: 0x2029,
	S: 0x09,
	WS: 0x20,
	ON: 0x21,
	LRE: 0x202a,
	LRO: 0x202d,
	RLE: 0x202b,
	RLO: 0x202e,
	PDF: 0x202c,
	LRI: 0x2066,
	RLI: 0x2067,
	FSI: 0x2068,
	PDI: 0x2069
}

// Runs BidiTest.txt (class sequences under @Levels/@Reorder headers, each for a bitset of auto 1, LTR 2,
// RTL 4); returns the case count and the first failures.
function classConformance(source: string, keep = 10): { cases: number; failures: number; first: string[] } {
	const resolver = new BidiResolver()
	const levels = new Uint8Array(256)
	const compact = new Uint8Array(256)
	const compactIndex = new Int32Array(256)
	const order = new Int32Array(256)
	const first: string[] = []
	let expectedLevels = ""
	let expectedOrder = ""
	let cases = 0
	let failures = 0

	for (const raw of source.split("\n")) {
		const line = raw.replace(/#.*/, "").trim()

		if (line === "") continue

		if (line.startsWith("@Levels:")) {
			expectedLevels = line.slice(8).trim()

			continue
		}

		if (line.startsWith("@Reorder:")) {
			expectedOrder = line.slice(9).trim()

			continue
		}

		const [classes = "", bits = "0"] = line.split(";")
		const text = String.fromCodePoint(
			...classes
				.trim()
				.split(" ")
				.map(name => SAMPLES[name] ?? 0x21)
		)

		for (const [bit, direction] of [
			[1, "auto"],
			[2, "ltr"],
			[4, "rtl"]
		] as const) {
			if ((Number(bits) & bit) === 0) continue

			cases++

			const level = resolver.resolve(text, 0, text.length, direction, levels)
			const removed = Array.from(levels.subarray(0, text.length), value => value === BIDI_REMOVED)

			applyL1(text, 0, levels, 0, text.length, level)

			let count = 0

			for (let k = 0; k < text.length; k++) {
				if (removed[k] === true) continue

				compact[count] = levels[k] ?? 0
				compactIndex[count++] = k
			}

			reorder(compact, 0, count, order)

			const actualLevels = removed.map((gone, k) => (gone ? "x" : String(levels[k]))).join(" ")
			const actualOrder = Array.from(order.subarray(0, count), index => compactIndex[index]).join(" ")

			if (actualLevels !== expectedLevels || actualOrder !== expectedOrder) {
				failures++

				if (first.length < keep)
					first.push(`${classes} (${direction}) got ${actualLevels} / ${actualOrder}, want ${expectedLevels} / ${expectedOrder}`)
			}
		}
	}

	return { cases, failures, first }
}

describe("UAX #9 conformance", () => {
	it.runIf(process.env["FILEN_OFFICE_UNICODE_FULL"] === "1")("passes the full BidiTest", () => {
		const result = classConformance(readCached("BidiTest.txt"))

		expect(result.cases).toBeGreaterThan(700_000)
		expect(result.first).toEqual([])
		expect(result.failures).toBe(0)
	})

	it("passes the sampled BidiCharacterTest cases", () => {
		const result = conformance(readFileSync(new URL("fixtures/unicode/BidiCharacterTest.sample.txt", import.meta.url), "utf8"))

		expect(result.cases).toBeGreaterThan(1000)
		expect(result.first).toEqual([])
	})

	it.runIf(process.env["FILEN_OFFICE_UNICODE_FULL"] === "1")("passes the full BidiCharacterTest", () => {
		const result = conformance(readCached("BidiCharacterTest.txt"))

		expect(result.cases).toBe(91707)
		expect(result.first).toEqual([])
		expect(result.failures).toBe(0)
	})
})

function resolveAll(text: string, direction: BidiDirection): { level: number; levels: number[] } {
	const levels = new Uint8Array(text.length)
	const level = new BidiResolver().resolve(text, 0, text.length, direction, levels)

	return { level, levels: Array.from(levels) }
}

describe("BidiResolver", () => {
	it("resolves left-to-right text to level 0", () => {
		const text = "Hello, world 123 (ok)."

		expect(hasRightToLeft(text, 0, text.length)).toBe(false)
		expect(resolveAll(text, "auto")).toEqual({ level: 0, levels: new Array<number>(text.length).fill(0) })
	})

	it("detects right-to-left content", () => {
		expect(hasRightToLeft("abc א", 0, 5)).toBe(true)
		expect(hasRightToLeft("abc א", 0, 4)).toBe(false)
		expect(hasRightToLeft("x‫x", 0, 3)).toBe(true)
		expect(hasRightToLeft("١", 0, 1)).toBe(true)
	})

	it("takes an auto paragraph's direction from its first strong character", () => {
		// "shalom abc": Hebrew first, so the paragraph is RTL and the Latin word sits at level 2.
		const { level, levels } = resolveAll("שלום abc", "auto")

		expect(level).toBe(1)
		expect(levels).toEqual([1, 1, 1, 1, 1, 2, 2, 2])
		expect(resolveAll("abc ש", "auto").level).toBe(0)
		expect(resolveAll("⁧ש⁩ abc", "auto").level).toBe(0)
	})

	it("raises numbers in right-to-left text", () => {
		const { levels } = resolveAll("ש 123", "rtl")

		expect(levels).toEqual([1, 1, 2, 2, 2])
	})

	it("gives both units of a surrogate pair the same level", () => {
		// U+10900 PHOENICIAN LETTER ALF is R; U+1D400 MATHEMATICAL BOLD CAPITAL A is L.
		const text = "\u{10900}\u{1d400}"
		const { levels } = resolveAll(text, "rtl")

		expect(text.length).toBe(4)
		expect(levels).toEqual([1, 1, 2, 2])
	})

	it("marks characters removed by X9", () => {
		const { levels } = resolveAll("a‫b‬", "ltr")

		expect(levels).toEqual([0, BIDI_REMOVED, 2, BIDI_REMOVED])
	})

	it("resolves several paragraphs with one instance", () => {
		const resolver = new BidiResolver()
		const levels = new Uint8Array(32)

		expect(resolver.resolve("xxא b", 2, 5, "auto", levels)).toBe(1)
		expect(Array.from(levels.subarray(0, 3))).toEqual([1, 1, 2])
		expect(resolver.resolve("ab", 0, 2, "auto", levels)).toBe(0)
		expect(Array.from(levels.subarray(0, 2))).toEqual([0, 0])
	})
})

describe("applyL1", () => {
	it("resets trailing whitespace and segment separators to the paragraph level", () => {
		const text = "א b\tב c  "
		const levels = new Uint8Array(text.length)
		const level = new BidiResolver().resolve(text, 0, text.length, "rtl", levels)

		applyL1(text, 0, levels, 0, text.length, level)

		expect(Array.from(levels)).toEqual([1, 1, 2, 1, 1, 1, 2, 1, 1])
	})

	it("works on a line inside a longer text", () => {
		const text = "abcא  "
		const levels = Uint8Array.from([0, 0, 1, 1, 1])

		applyL1(text, 2, levels, 0, 5, 1)

		expect(Array.from(levels)).toEqual([0, 0, 1, 1, 1])
	})
})

describe("reorder", () => {
	function visual(levels: number[]): number[] {
		const out = new Int32Array(levels.length)

		reorder(Uint8Array.from(levels), 0, levels.length, out)

		return Array.from(out)
	}

	it("reverses right-to-left runs", () => {
		expect(visual([0, 1, 1, 0])).toEqual([0, 2, 1, 3])
		expect(visual([1, 1, 2, 2, 1])).toEqual([4, 2, 3, 1, 0])
		expect(visual([0, 0, 0])).toEqual([0, 1, 2])
		expect(visual([2, 2])).toEqual([0, 1])
	})

	it("orders removed items with their neighbours", () => {
		expect(visual([1, BIDI_REMOVED, 1, 0])).toEqual([2, 1, 0, 3])
		expect(visual([BIDI_REMOVED, 1, 1])).toEqual([2, 1, 0])
	})

	it("reorders a sub-range", () => {
		const out = new Int32Array(2)

		reorder(Uint8Array.from([0, 1, 1, 0]), 1, 3, out)
		expect(Array.from(out)).toEqual([1, 0])
	})
})

describe("performance", () => {
	it("resolves 10k mixed paragraphs reusing its scratch space", () => {
		const paragraphs: string[] = []

		for (let p = 0; p < 10_000; p++) {
			paragraphs.push(
				`Paragraph ${p}: the word שלום (${p % 97}) means peace, עברית [${p}] and more text here to reach a realistic length.`
			)
		}

		const resolver = new BidiResolver()
		const levels = new Uint8Array(1024)

		for (const text of paragraphs) resolver.resolve(text, 0, text.length, "auto", levels)

		const scratch = Reflect.get(resolver, "cps")
		const started = performance.now()

		for (const text of paragraphs) resolver.resolve(text, 0, text.length, "auto", levels)

		const elapsed = performance.now() - started

		console.log(`bidi: 10k paragraphs in ${elapsed.toFixed(1)} ms`)
		expect(Reflect.get(resolver, "cps")).toBe(scratch)
		expect(elapsed).toBeLessThan(budget(2000))
	})
})
