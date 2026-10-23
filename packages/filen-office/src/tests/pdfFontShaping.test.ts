import { readFileSync } from "node:fs"
import { gunzipSync } from "node:zlib"
import { describe, expect, it } from "vitest"
import { readSfnt } from "../core/fonts/sfnt"
// Loading the font set registers the Indic shaper, as every export does.
import "../core/pdfwrite/fonts/fontSet"
import { ShapingFace } from "../core/pdfwrite/fonts/shape"
import { SCRIPT_COMMON, SCRIPT_INHERITED, scriptOf } from "../core/pdfwrite/fonts/unicode"

// Shaping against HarfBuzz: fixtures/pdfFonts/harfbuzz.json holds hb-shape's output (glyph ids and absolute
// origins in font units, visual order) for samples of every script the PDF fonts cover, shaped with the
// faces under assets/fonts.

interface ShapingCase {
	readonly font: string
	readonly text: string
	readonly rtl: boolean
	readonly language?: string
	readonly glyphs: string
}

const FIXTURE = JSON.parse(readFileSync(new URL("./fixtures/pdfFonts/harfbuzz.json", import.meta.url), "utf8")) as {
	readonly cases: readonly ShapingCase[]
}
const faces = new Map<string, ShapingFace>()

function face(file: string): ShapingFace {
	let shaping = faces.get(file)

	if (shaping === undefined) {
		const bytes = new Uint8Array(gunzipSync(readFileSync(new URL(`../../assets/fonts/${file}`, import.meta.url))))

		shaping = new ShapingFace(readSfnt(bytes))
		faces.set(file, shaping)
	}

	return shaping
}

function shape(sample: ShapingCase): string {
	const cps: number[] = []
	const clusters: number[] = []

	for (let i = 0; i < sample.text.length;) {
		const cp = sample.text.codePointAt(i) ?? 0

		cps.push(cp)
		clusters.push(i)
		i += cp > 0xffff ? 2 : 1
	}

	const script = cps.map(scriptOf).find(s => s !== SCRIPT_COMMON && s !== SCRIPT_INHERITED) ?? 0
	const result = face(sample.font).shape({
		cps: Uint32Array.from(cps),
		clusters: Uint32Array.from(clusters),
		count: cps.length,
		before: [],
		after: [],
		script,
		rtl: sample.rtl,
		language: sample.language
	})
	const n = result.buffer.length
	const order = Array.from({ length: n }, (_, k) => (sample.rtl ? n - 1 - k : k))

	return order.map(k => `${result.buffer.glyph[k]}@${Math.round(result.x[k] ?? 0)},${Math.round(result.y[k] ?? 0)}`).join(" ")
}

describe("PDF text shaping", () => {
	for (const sample of FIXTURE.cases) {
		it(`gives HarfBuzz's glyphs and positions for "${sample.text.slice(0, 24)}" in ${sample.font.split("/").pop() ?? ""}`, () => {
			expect(shape(sample)).toBe(sample.glyphs)
		})
	}

	it("ignores a substitution table it cannot read and draws the characters' own glyphs", () => {
		const bytes = new Uint8Array(
			gunzipSync(readFileSync(new URL("../../assets/fonts/noto/noto-naskh-arabic/NotoNaskhArabic-Regular.ttf.gz", import.meta.url)))
		)
		const font = readSfnt(bytes)
		const gsub = font.table("GSUB")

		expect(gsub).toBeDefined()

		// An unknown major version.
		new DataView(bytes.buffer, bytes.byteOffset).setUint16(gsub?.offset ?? 0, 9)

		const shaping = new ShapingFace(readSfnt(bytes))

		expect(shaping.gsub).toBeUndefined()

		const cps = Uint32Array.from([0x0644, 0x0627])
		const result = shaping.shape({
			cps,
			clusters: Uint32Array.from([0, 1]),
			count: 2,
			before: [],
			after: [],
			script: scriptOf(0x0644),
			rtl: true,
			language: undefined
		})

		expect(Array.from(result.buffer.glyph.subarray(0, result.buffer.length))).toEqual([
			shaping.glyphOf(0x0644),
			shaping.glyphOf(0x0627)
		])
	})
})
