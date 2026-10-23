import { readFileSync } from "node:fs"
import { gunzipSync } from "node:zlib"
import { describe, expect, it } from "vitest"
import type { PdfFaceSource } from "../core/fonts/manager"
import { readSfnt } from "../core/fonts/sfnt"
import type { PdfFontAssets } from "../core/pdfwrite/fonts/assets"
import { PdfFontSet, type PdfFontFaceRequest } from "../core/pdfwrite/fonts/fontSet"
import { gsubTable, multiply, withTable } from "./pdfFontLimitsFixtures"

const ASSETS = new URL("../../assets/fonts/", import.meta.url)
const assets: PdfFontAssets = { open: () => Promise.reject(new Error("No assets")) }
const CALADEA = new Uint8Array(gunzipSync(readFileSync(new URL("caladea/Caladea-Regular.ttf.gz", ASSETS))))
const SFNT = readSfnt(CALADEA)
const A = SFNT.glyphId(0x61)
const X = SFNT.glyphId(0x78)
const REQUEST: PdfFontFaceRequest = { family: "Embedded", bold: false, italic: false, slot: "hAnsi" }
const DOUBLINGS = 20

function words(...values: number[]): number[] {
	return values.flatMap(value => [(value >>> 8) & 0xff, value & 0xff])
}

// SequenceContext format 3 matching `glyph` and running lookup `target` at it.
function context(glyph: number, target: number): () => number[] {
	return () => words(3, 1, 1, 12, 0, target, 1, 1, glyph)
}

// Contextual lookups that each double "a" through a nested lookup (so a run of "a" passes the limits inside a
// nested lookup), then "x" doubles through a nested lookup that only its context runs.
function nestingFont(): Uint8Array {
	const lookups: (readonly [number, () => number[]])[] = []

	for (let k = 0; k < DOUBLINGS; k++) lookups.push([5, context(A, DOUBLINGS + k)])
	for (let k = 0; k < DOUBLINGS; k++) lookups.push([2, multiply(A, 2)])

	lookups.push([5, context(X, 2 * DOUBLINGS + 1)])
	lookups.push([2, multiply(X, 2)])

	const table = gsubTable(lookups)
	const view = new DataView(table.buffer)
	const feature = view.getUint16(6) + 8

	// The last lookup is reached only through its context, not by the feature.
	view.setUint16(feature + 2, lookups.length - 1)

	return withTable(CALADEA, "GSUB", table)
}

describe("PDF font nested lookups past the shaping limits", () => {
	it("keeps running contextual lookups on a face after its runs went past the limits inside a nested lookup", async () => {
		const bytes = nestingFont()
		const source: PdfFaceSource = {
			kind: "embedded",
			key: "nesting",
			read: () => Promise.resolve(bytes),
			fsType: 0,
			weight: 400,
			italic: false,
			substitute: "arimo"
		}
		const fonts = new PdfFontSet({ assets, resolve: () => source })
		const shape = async (text: string): Promise<readonly number[]> => {
			const pending = fonts.prepare(text, 0, text.length, REQUEST)

			if (pending !== undefined) await pending

			return fonts.shape(text, 0, text.length, REQUEST, { level: 0, flags: 0 }).runs.flatMap(run => [...run.glyphs])
		}

		expect(await shape("x")).toHaveLength(2)

		for (let k = 0; k < 20; k++) await shape("a")

		expect(fonts.unshaped()).toBe(20)
		expect(await shape("x")).toHaveLength(2)
	})
})
