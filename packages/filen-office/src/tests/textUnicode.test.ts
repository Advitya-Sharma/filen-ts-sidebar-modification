import { describe, expect, it } from "vitest"
import {
	BidiClass,
	LINE_BREAK_CLASS_MASK,
	LineBreakClass,
	ParagraphBuilder,
	TAB_CHAR,
	BREAK_CHAR,
	OBJECT_CHAR,
	UNICODE_VERSION
} from "../core/text/index"
import { CodePointTable, LINE_BREAK_CLASS_NAMES, packNumbers, packRuns, unpackNumbers } from "../core/text/packed"
import { bidiBrackets, bidiClassTable, lineBreakTable } from "../core/text/unicode"
import { canMeasureOffscreen, canvasMeasurer, offscreenMeasurer, type CanvasContextLike } from "../dom/index"
import { UNICODE_VERSION as PINNED } from "../scripts/unicodeTables"

describe("packed property tables", () => {
	it("round-trips numbers and runs", () => {
		const numbers = [0, 1, 31, 32, 1023, 1024, 0x10ffff]

		expect(unpackNumbers(packNumbers(numbers, false), false)).toEqual(numbers)
		expect(unpackNumbers(packNumbers([5, 3, 70000, 2], true), true)).toEqual([5, 3, 70000, 2])

		const table = new CodePointTable(
			packRuns([
				[0x41, 1],
				[0x10000 - 0x41, 2],
				[0x20000, 3],
				[0x10ffff + 1 - 0x30000, 4]
			])
		)

		expect([
			table.get(0),
			table.get(0x41),
			table.get(0xffff),
			table.get(0x10000),
			table.get(0x2ffff),
			table.get(0x30000),
			table.get(0x10ffff)
		]).toEqual([1, 2, 2, 3, 3, 4, 4])
	})

	it("ships the pinned Unicode version", () => {
		expect(UNICODE_VERSION).toBe(PINNED)
	})

	it("expands to a small table", () => {
		expect(lineBreakTable().byteLength).toBeLessThan(96 * 1024)
		expect(bidiClassTable().byteLength).toBeLessThan(96 * 1024)
	})

	it("knows representative properties", () => {
		const lb = (cp: number): string => LINE_BREAK_CLASS_NAMES[lineBreakTable().get(cp) & LINE_BREAK_CLASS_MASK] ?? "?"

		expect([
			lb(0x41),
			lb(0x20),
			lb(0x30),
			lb(0x4e00),
			lb(0x3001),
			lb(0x05d0),
			lb(0x0e01),
			lb(0x0e31),
			lb(0x1f600),
			lb(0x201c),
			lb(0x201d)
		]).toEqual(["AL", "SP", "NU", "ID", "CL", "HL", "AL", "CM", "ID", "QU_PI", "QU_PF"])
		expect(lineBreakTable().get(0x4e00) & 0x40).toBe(0x40)
		expect(LineBreakClass.CJ).toBe(LINE_BREAK_CLASS_NAMES.indexOf("CJ"))

		const bidi = bidiClassTable()

		expect([
			bidi.get(0x41),
			bidi.get(0x5d0),
			bidi.get(0x627),
			bidi.get(0x31),
			bidi.get(0x661),
			bidi.get(0x9),
			bidi.get(0x2067)
		]).toEqual([BidiClass.L, BidiClass.R, BidiClass.AL, BidiClass.EN, BidiClass.AN, BidiClass.S, BidiClass.RLI])
		// Pairs share an id; U+2329 pairs like U+3008 (canonical equivalents).
		const brackets = bidiBrackets()

		expect((brackets.get(0x28) ?? 0) >> 2).toBe((brackets.get(0x29) ?? 0) >> 2)
		expect((brackets.get(0x2329) ?? 0) >> 2).toBe((brackets.get(0x3009) ?? 0) >> 2)
	})
})

describe("paragraph builder", () => {
	it("records runs, tabs, breaks and objects and sanitises special characters in text", () => {
		const builder = new ParagraphBuilder()
		const paragraph = builder.reset(9).append("a\tb\u000bc\nd￼", 1).append("e", 1).append("f", 2).object(2, 10, 8, 2, 77).finish()

		expect(paragraph.text).toBe(`a${TAB_CHAR}b${BREAK_CHAR}c d�ef${OBJECT_CHAR}`)
		expect([paragraph.runCount, [...paragraph.runStart.subarray(0, 2)], [...paragraph.runStyle.subarray(0, 2)]]).toEqual([
			2,
			[0, 9],
			[1, 2]
		])
		expect([paragraph.tabCount, paragraph.breakCount, paragraph.objectCount, paragraph.objectId[0]]).toEqual([1, 1, 1, 77])
		expect(paragraph.markStyle).toBe(9)
	})
})

describe("canvas measurer", () => {
	class FakeContext implements CanvasContextLike {
		font = ""
		fontKerning = "auto"
		fontSets = 0
		kerningSets = 0

		measureText(text: string): { width: number; fontBoundingBoxAscent: number; fontBoundingBoxDescent: number } {
			const size = Number(/([\d.]+)px/.exec(this.font)?.[1] ?? "0")

			return { width: text.length * size * 0.5, fontBoundingBoxAscent: size * 0.9, fontBoundingBoxDescent: size * 0.3 }
		}
	}

	it("switches font and kerning only when they change and reads extents at a large size", () => {
		const context = new FakeContext()
		const counting = new Proxy(context, {
			set(target, key, value) {
				if (key === "font") target.fontSets++
				if (key === "fontKerning") target.kerningSets++

				return Reflect.set(target, key, value)
			}
		})
		const measurer = canvasMeasurer(counting)
		const widths = new Float64Array(4)

		expect(measurer.kerningControl).toBe(true)

		measurer.measure("400 20px x", true, ["ab", "abc"], 2, widths)
		measurer.measure("400 20px x", true, ["abcd"], 1, widths)
		measurer.measure("400 20px x", false, ["a"], 1, widths)

		expect([...widths.subarray(0, 1)]).toEqual([10])
		expect(context.fontSets).toBe(1)
		expect(context.kerningSets).toBe(2)

		const extents = measurer.fontExtents("400 11px x")

		expect(extents.ascent).toBeCloseTo(9.9, 9)
		expect(extents.descent).toBeCloseTo(3.3, 9)
	})

	it("probes for OffscreenCanvas and fails descriptively without it", () => {
		expect(canMeasureOffscreen({})).toBe(false)
		expect(canMeasureOffscreen(null)).toBe(false)
		expect(() => offscreenMeasurer({})).toThrow(/OffscreenCanvas is not available/)

		const scope = {
			OffscreenCanvas: class {
				getContext(): CanvasContextLike {
					return { font: "", fontKerning: "auto", measureText: text => ({ width: text.length }) }
				}
			}
		}

		expect(canMeasureOffscreen(scope)).toBe(true)
		expect(offscreenMeasurer(scope).kerningControl).toBe(true)
	})

	it("reports no kerning control without fontKerning (WebKit)", () => {
		const context: CanvasContextLike = { font: "", measureText: text => ({ width: text.length }) }

		expect(canvasMeasurer(context).kerningControl).toBe(false)
	})
})
