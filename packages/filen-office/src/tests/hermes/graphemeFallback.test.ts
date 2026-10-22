import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { legacyText, type LegacyField } from "../../core/docx/forms/controls"
import { cpuMs } from "../timing"

// Hermes has no Intl.Segmenter: legacy text fields cut at w:maxLength with the fallback, which must agree with
// Intl.Segmenter's grapheme clusters on what documents hold.

const FIELD: LegacyField = {
	type: "FORMTEXT",
	name: "Text1",
	enabled: true,
	checked: false,
	entries: [],
	result: 0,
	textType: "regular",
	maxLength: 1,
	format: null
}

const TEXTS = [
	"plain ascii",
	"e\u0301a\u0300\u0323b",
	"\u{1F44D}\u{1F3FD}x\u{1F468}\u200D\u{1F469}\u200D\u{1F467}\u200D\u{1F466}y",
	"\u{1F1E9}\u{1F1EA}\u{1F1EB}\u{1F1F7}\u{1F1EE}",
	"1\uFE0F\u20E3#\uFE0F\u20E3",
	"\u{1F3F4}\u{E0067}\u{E0062}\u{E0073}\u{E0063}\u{E0074}\u{E007F} scot",
	"a\r\nb\nc",
	"\u0915\u094D\u0937\u0924\u094D\u0930\u093F\u092F \u0939\u093F\u0928\u094D\u0926\u0940",
	"\u09AC\u09BE\u0982\u09B2\u09BE \u0995\u09CD\u09B7",
	"\uD55C\uAD6D\uC5B4 \uD14D\uC2A4\uD2B8",
	"\u65E5\u672C\u8A9E\u306E\u30C6\u30AD\u30B9\u30C8",
	"\u{10000}\u{1F600}",
	"x\u200Cy\u200Dz",
	"\uFDFA \u0645\u0631\u062D\u0628\u0627",
	"a\u0308\u0308\u0308"
]

const Segmenter = Intl.Segmenter

function expected(text: string, count: number): string {
	let out = ""
	let taken = 0

	for (const { segment } of new Segmenter(undefined, { granularity: "grapheme" }).segment(text)) {
		if (taken++ === count) break

		out += segment
	}

	return out
}

describe("legacy text maxLength without Intl.Segmenter", () => {
	beforeAll(() => {
		Reflect.deleteProperty(Intl, "Segmenter")
	})

	afterAll(() => {
		Reflect.set(Intl, "Segmenter", Segmenter)
	})

	it("cuts at the same grapheme boundaries as Intl.Segmenter", () => {
		expect(Reflect.get(Intl, "Segmenter")).toBeUndefined()

		for (const text of TEXTS) {
			for (let count = 1; count <= [...text].length + 1; count++) {
				expect(legacyText({ ...FIELD, maxLength: count }, text), `${JSON.stringify(text)} ${count}`).toBe(expected(text, count))
			}
		}
	})

	it("cuts a long run of viramas and marks in time linear in it", () => {
		// A virama, marks, then a letter joins a conjunct: a pattern trying that before the marks alone read every
		// virama's marks again, quadratic in the run.
		const text = `\u0915${"\u094D".repeat(40_000)}!`
		const started = cpuMs()

		expect(legacyText({ ...FIELD, maxLength: 1 }, text)).toBe(text.slice(0, -1))
		expect(cpuMs() - started).toBeLessThan(250)
	})

	it("keeps a field without a limit whole and an empty text empty", () => {
		expect(legacyText({ ...FIELD, maxLength: 0 }, "e\u0301x")).toBe("e\u0301x")
		expect(legacyText({ ...FIELD, maxLength: 3 }, "")).toBe("")
		expect(legacyText({ ...FIELD, maxLength: 2, format: "UPPERCASE" }, "a\u{1F600}bc")).toBe("A\u{1F600}")
	})
})
