import { createReadStream, readFileSync } from "node:fs"
import { Readable } from "node:stream"
import { gunzipSync } from "node:zlib"
import { describe, expect, it } from "vitest"
import type { PdfFaceSource } from "../core/fonts/manager"
import { readSfnt } from "../core/fonts/sfnt"
import { PdfWriter } from "../core/pdfwrite"
import type { PdfByteStream, PdfFontAssets } from "../core/pdfwrite/fonts/assets"
import { PdfFontSet, type PdfFontFaceRequest, type PdfShapedText } from "../core/pdfwrite/fonts/fontSet"
import { directoryRecord, gsubTable, multiply, recurse, withTable } from "./pdfFontLimitsFixtures"
import { cpuMs } from "./timing"

// Fonts whose layout tables or glyph data are hostile: shaping stays within HarfBuzz's buffer limits and a font
// that cannot be subset gives way to its substitute, so the export finishes either way.

const ASSETS = new URL("../../assets/fonts/", import.meta.url)
const assets: PdfFontAssets = {
	open: file => Promise.resolve(Readable.toWeb(createReadStream(new URL(file, ASSETS))) as unknown as PdfByteStream)
}
const CALADEA = new Uint8Array(gunzipSync(readFileSync(new URL("caladea/Caladea-Regular.ttf.gz", ASSETS))))
const GLYPH_A = readSfnt(CALADEA).glyphId(0x61)
const REQUEST: PdfFontFaceRequest = { family: "Embedded", bold: false, italic: false, slot: "hAnsi" }

// Caladea with its GSUB replaced.
function withGsub(table: Uint8Array): Uint8Array {
	return withTable(CALADEA, "GSUB", table)
}

function patched(mutate: (view: DataView, record: (tag: string) => number) => void): Uint8Array {
	const out = CALADEA.slice()
	const view = new DataView(out.buffer)

	mutate(view, tag => directoryRecord(out, tag))

	return out
}

function source(bytes: Uint8Array): PdfFaceSource {
	return {
		kind: "embedded",
		key: "hostile",
		read: () => Promise.resolve(bytes),
		fsType: 0,
		weight: 400,
		italic: false,
		substitute: "arimo"
	}
}

async function shaped(
	bytes: Uint8Array,
	text: string
): Promise<{ readonly fonts: PdfFontSet; readonly shaped: PdfShapedText; readonly ms: number }> {
	const fonts = new PdfFontSet({ assets, resolve: () => source(bytes) })
	const pending = fonts.prepare(text, 0, text.length, REQUEST)

	if (pending !== undefined) await pending

	const started = cpuMs()
	const result = fonts.shape(text, 0, text.length, REQUEST, { level: 0, flags: 0 })

	return { fonts, shaped: result, ms: cpuMs() - started }
}

function glyphCount(text: PdfShapedText): number {
	return text.runs.reduce((sum, run) => sum + run.glyphs.length, 0)
}

async function written(text: PdfShapedText): Promise<void> {
	const writer = await PdfWriter.create({ deflate: null })
	const page = writer.beginPage(200, 200)

	page.text(text.runs, { x: 10, y: 10, size: 12 })

	await writer.endPage(page)
	await writer.finish()
}

describe("PDF font limits", () => {
	it("keeps a font's multiplying substitutions that stay within the limits", async () => {
		const { fonts, shaped: result } = await shaped(
			withGsub(gsubTable(Array.from({ length: 12 }, () => [2, multiply(GLYPH_A, 2)] as const))),
			"a"
		)

		expect(glyphCount(result)).toBe(4096)
		expect(fonts.unshaped()).toBe(0)
		fonts.close()
	})

	it("draws a run unshaped, quickly, when a font's substitutions would multiply it past the length limit", async () => {
		const {
			fonts,
			shaped: result,
			ms
		} = await shaped(withGsub(gsubTable(Array.from({ length: 24 }, () => [2, multiply(GLYPH_A, 2)] as const))), "aaaaaaaaaa")

		expect(glyphCount(result)).toBe(10)
		expect(fonts.unshaped()).toBe(1)
		expect(ms).toBeLessThan(500)
		fonts.close()
	})

	it("draws a run unshaped when a font's contextual lookups recurse past the operations limit", async () => {
		const { fonts, shaped: result, ms } = await shaped(withGsub(gsubTable([[5, recurse(GLYPH_A, 4)]])), "aaaa")

		expect(glyphCount(result)).toBe(4)
		expect(fonts.unshaped()).toBe(1)
		expect(ms).toBeLessThan(500)
		fonts.close()
	})

	it("shapes the next run of the same font again after one went past the limits", async () => {
		// 4096 glyphs for each character: five go past the length limit's floor of 16384.
		const { fonts, shaped: long } = await shaped(
			withGsub(gsubTable(Array.from({ length: 12 }, () => [2, multiply(GLYPH_A, 2)] as const))),
			"aaaaa"
		)
		const short = fonts.shape("a", 0, 1, REQUEST, { level: 0, flags: 0 })

		expect(glyphCount(long)).toBe(5)
		expect(glyphCount(short)).toBe(4096)
		expect(fonts.unshaped()).toBe(1)
		fonts.close()
	})

	it("shapes a long run with many default ignorables in linear time", async () => {
		const text = "a​".repeat(20000)
		const { fonts, shaped: result, ms } = await shaped(CALADEA, text)

		expect(glyphCount(result)).toBe(20000)
		expect(ms).toBeLessThan(1000)
		fonts.close()
	})

	const unsubsettable: readonly (readonly [string, Uint8Array])[] = [
		["without loca", patched((view, record) => view.setUint8(record("loca"), 0x78))],
		["with a loca too short for its glyphs", patched((view, record) => view.setUint32(record("loca") + 12, 2))],
		["without glyf", patched((view, record) => view.setUint8(record("glyf"), 0x78))],
		["with a short head", patched((view, record) => view.setUint32(record("head") + 12, 40))]
	]

	for (const [name, bytes] of unsubsettable) {
		it(`draws with the substitute an embedded TrueType font ${name}`, async () => {
			const { fonts, shaped: result } = await shaped(bytes, "office")

			expect(result.runs[0]?.font.postScriptName).toMatch(/^Arimo/)
			await expect(written(result)).resolves.toBeUndefined()
			fonts.close()
		})
	}
})
