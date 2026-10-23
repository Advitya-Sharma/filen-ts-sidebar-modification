import { readFileSync } from "node:fs"
import { gunzipSync } from "node:zlib"
import { describe, expect, it } from "vitest"
import { readSfnt, type SfntFont } from "../core/fonts/sfnt"
import { PDF_FACES } from "../core/pdfwrite/fonts/faces.generated"
import { NO_GLYPH, subsetTrueType } from "../core/pdfwrite/fonts/subset"

function face(file: string): SfntFont {
	return readSfnt(new Uint8Array(gunzipSync(readFileSync(new URL(`../../assets/fonts/${file}`, import.meta.url)))))
}

function table(font: SfntFont, tag: string): Uint8Array {
	const range = font.table(tag)

	if (range === undefined) throw new Error(`no ${tag}`)

	return font.bytes.subarray(range.offset, range.offset + range.length)
}

// [start, end) of each glyph's data in glyf.
function glyphData(font: SfntFont, glyph: number): Uint8Array {
	const head = new DataView(table(font, "head").buffer, table(font, "head").byteOffset)
	const loca = table(font, "loca")
	const view = new DataView(loca.buffer, loca.byteOffset, loca.byteLength)
	const long = head.getInt16(50) === 1
	const start = long ? view.getUint32(glyph * 4) : view.getUint16(glyph * 2) * 2
	const end = long ? view.getUint32(glyph * 4 + 4) : view.getUint16(glyph * 2 + 2) * 2

	return table(font, "glyf").subarray(start, end)
}

function usedFor(font: SfntFont, text: string): Uint8Array {
	const used = new Uint8Array(font.numGlyphs)

	for (const char of text) used[font.glyphId(char.codePointAt(0) ?? 0)] = 1

	return used
}

// A page of running text.
const PAGE = (
	"The quick brown fox jumps over the lazy dog. Typography, office documents and spreadsheets; " +
	"numbers 0123456789 (and punctuation: “quotes”, ‘apostrophes’ — dashes – and hyphens-). "
).repeat(30)

describe("TrueType subsetting for PDF", () => {
	it("numbers the used glyphs densely and keeps their outlines without hinting instructions", () => {
		const font = face("carlito/Carlito-Regular.ttf.gz")
		const used = usedFor(font, "Office")
		const { program, glyphs } = subsetTrueType(font, used)
		const subset = readSfnt(program)
		const kept = [0, ...[...used.keys()].filter(g => used[g] === 1)]

		// The used glyphs and the components of composite ones, nothing else.
		expect(subset.numGlyphs).toBe(glyphs.filter(glyph => glyph !== NO_GLYPH).length)
		expect(subset.numGlyphs).toBeLessThan(16)
		expect(glyphs[0]).toBe(0)
		expect(glyphs[font.glyphId(0x5a)]).toBe(NO_GLYPH)

		for (const glyph of kept) {
			const original = glyphData(font, glyph)

			// Composite glyphs (their components renumbered) are checked below.
			if (new DataView(original.buffer, original.byteOffset).getInt16(0) < 0) continue
			const copied = glyphData(subset, glyphs[glyph] ?? NO_GLYPH)
			const view = new DataView(original.buffer, original.byteOffset, original.byteLength)
			const contours = view.getInt16(0)
			const lengthAt = 10 + contours * 2
			const instructions = view.getUint16(lengthAt)

			expect(subset.advance(glyphs[glyph] ?? NO_GLYPH)).toBe(font.advance(glyph))
			// Header and contour ends as they were, no instructions, then the same flags and coordinates.
			expect(Array.from(copied.subarray(0, lengthAt))).toEqual(Array.from(original.subarray(0, lengthAt)))
			expect(new DataView(copied.buffer, copied.byteOffset).getUint16(lengthAt)).toBe(0)
			expect(Array.from(copied.subarray(lengthAt + 2, lengthAt + 2 + original.length - lengthAt - 2 - instructions))).toEqual(
				Array.from(original.subarray(lengthAt + 2 + instructions))
			)
		}

		for (const tag of ["fpgm", "prep", "cvt "]) expect(subset.table(tag)).toBeUndefined()
	})

	it("pulls in the components of composite glyphs and points them at their new ids", () => {
		const font = face("tinos/Tinos-Regular.ttf.gz")
		const composite = [..."ÅÄÉÖüçñ"]
			.map(char => font.glyphId(char.codePointAt(0) ?? 0))
			.find(glyph => {
				const data = glyphData(font, glyph)

				return data.length >= 10 && new DataView(data.buffer, data.byteOffset).getInt16(0) < 0
			})

		expect(composite).toBeDefined()

		const used = new Uint8Array(font.numGlyphs)

		used[composite ?? 0] = 1

		const { program, glyphs } = subsetTrueType(font, used)
		const subset = readSfnt(program)
		const data = glyphData(font, composite ?? 0)
		const component = new DataView(data.buffer, data.byteOffset).getUint16(12)
		const copied = glyphData(subset, glyphs[composite ?? 0] ?? NO_GLYPH)

		expect(new DataView(copied.buffer, copied.byteOffset).getUint16(12)).toBe(glyphs[component])
		expect(glyphData(subset, glyphs[component] ?? NO_GLYPH).length).toBeGreaterThan(0)
		expect(subset.numGlyphs).toBeLessThan(10)
	})

	it("writes valid table checksums and a whole-file checksum adjustment", () => {
		const font = face("cousine/Cousine-Bold.ttf.gz")
		const bytes = subsetTrueType(font, usedFor(font, "checksum")).program
		const view = new DataView(bytes.buffer)
		let total = 0

		for (let i = 0; i + 4 <= bytes.length; i += 4) total = (total + view.getUint32(i)) >>> 0

		expect(total).toBe(0xb1b0afba)

		const tables = view.getUint16(4)

		for (let t = 0; t < tables; t++) {
			const record = 12 + t * 16
			const tag = String.fromCharCode(...bytes.subarray(record, record + 4))
			const offset = view.getUint32(record + 8)
			const length = view.getUint32(record + 12)
			let sum = 0

			for (let i = 0; i < length; i += 4) {
				let word = 0

				for (let k = 0; k < 4; k++) word = word * 256 + (i + k < length ? (bytes[offset + i + k] ?? 0) : 0)

				sum = (sum + word) >>> 0
			}

			if (tag !== "head") expect(view.getUint32(record + 4)).toBe(sum)
		}
	})

	// Faces of a few hundred glyphs (Caladea, Selawik) spend a fifth of them on a page of Latin text, so for them
	// the bound is what the used glyphs themselves take plus the tables every subset keeps.
	for (const faces of Object.values(PDF_FACES)) {
		const file = faces[0]?.file ?? ""

		it(`keeps a one-page subset of ${file} within 15 % of the full face, or its used glyphs for small faces`, () => {
			const font = face(file)
			const used = usedFor(font, PAGE)
			const subset = subsetTrueType(font, used).program

			if (font.numGlyphs >= 1000) {
				expect(subset.length / font.bytes.length).toBeLessThanOrEqual(0.15)

				return
			}

			let glyphs = 0

			for (let g = 0; g < font.numGlyphs; g++) if (used[g] === 1 || g === 0) glyphs += (glyphData(font, g).length + 3) & ~3

			const kept = ["head", "hhea", "maxp", "hmtx", "loca", "name", "OS/2"].reduce(
				(sum, tag) => sum + (font.table(tag)?.length ?? 0),
				0
			)

			expect(subset.length).toBeLessThanOrEqual(glyphs + kept + 1024)
		})
	}
})
