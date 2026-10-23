import type { SfntFont } from "../../fonts/sfnt"
import { pdfName, pdfRaw, type PdfRef } from "../objects"
import type { PdfObjectSink } from "../writer"
import { subsetTrueType } from "./subset"

// One embedded face as a Type0 font (ISO 32000-1 §9.7): Identity-H, a CIDFontType2 with a subset TrueType
// program (FontFile2) or a CIDFontType0 with the whole CFF-flavoured OpenType program (FontFile3 /OpenType),
// widths in /W, and a ToUnicode CMap (§9.10.3).
//
// TrueType: character codes (CIDs) are numbered densely from 1 and mapped to glyphs by a CIDToGIDMap stream.
// Each code is one glyph with one text, so the same glyph can stand for different text (a reordered syllable,
// a shared contextual form) with ToUnicode still exact. A code's /W width is always its glyph's advance:
// PDF/UA-1 §7.21.5 (and PDF/A) require /W to agree with the font program, so kerning and other positioning
// stay in the content stream. CFF programs have no CIDToGIDMap: codes are glyph ids, and a glyph's first text
// stays its text (the run reports the conflict and carries /ActualText).

const MAX_CODE = 0xffff
// OS/2 fsType: embedding must not subset the font.
const NO_SUBSETTING = 0x0100

export interface PdfFontInfo {
	// Without the subset tag.
	readonly postScriptName: string
	readonly flags: number
	readonly bbox: readonly [number, number, number, number]
	readonly italicAngle: number
	readonly ascent: number
	readonly descent: number
	readonly capHeight: number
	readonly stemV: number
}

export class EmbeddedPdfFont {
	readonly unitsPerEm: number
	readonly postScriptName: string
	readonly sfnt: SfntFont
	private readonly scale: number
	private readonly info: PdfFontInfo
	// Per code (index): glyph, width (1/1000 em) and text. Code 0 is unused (TrueType).
	private glyphOf: number[] = [0]
	private widthOf: number[] = [0]
	private textOf: string[] = [""]
	private readonly byKey = new Map<string, number>()
	// The first code of each glyph, for glyphs past the last code.
	private readonly firstCode = new Map<number, number>()
	// CFF: text per glyph id (code).
	private readonly cffText = new Map<number, string>()
	private closed = false

	constructor(sfnt: SfntFont) {
		this.sfnt = sfnt
		this.unitsPerEm = sfnt.unitsPerEm
		this.scale = 1000 / sfnt.unitsPerEm
		this.info = fontInfo(sfnt)
		this.postScriptName = this.info.postScriptName
	}

	get cff(): boolean {
		return this.sfnt.cff
	}

	// The glyph's own advance, 1/1000 em, as /W would give it.
	glyphAdvance(glyph: number): number {
		return Math.round(this.sfnt.advance(glyph) * this.scale)
	}

	// /W width of a code, 1/1000 em.
	advance(code: number): number {
		if (this.sfnt.cff) return this.glyphAdvance(code)

		return this.widthOf[code] ?? 0
	}

	// The code that draws `glyph` for `text`. Not exact when no code can carry that text (CFF, or 65535 codes
	// used): the caller then marks the run with /ActualText.
	code(glyph: number, text: string): { readonly code: number; readonly exact: boolean } {
		if (this.sfnt.cff) {
			const known = this.cffText.get(glyph)

			if (known === undefined) this.cffText.set(glyph, text)

			return { code: glyph, exact: known === undefined || known === text }
		}

		const key = `${glyph}\u0000${text}`
		let code = this.byKey.get(key)

		if (code === undefined) {
			code = this.glyphOf.length

			if (code > MAX_CODE) return { code: this.firstCode.get(glyph) ?? 0, exact: false }

			this.glyphOf.push(glyph)
			this.widthOf.push(this.glyphAdvance(glyph))
			this.textOf.push(text)
			this.byKey.set(key, code)

			if (!this.firstCode.has(glyph)) this.firstCode.set(glyph, code)
		}

		return { code, exact: true }
	}

	async write(sink: PdfObjectSink, ref: PdfRef): Promise<void> {
		if (this.closed) throw new Error("The font was closed before it was written")

		const sfnt = this.sfnt
		const cff = sfnt.cff
		const codes = cff
			? [...this.cffText.keys()].sort((a, b) => a - b)
			: Array.from({ length: this.glyphOf.length - 1 }, (_, k) => k + 1)
		const glyphOfCode = (code: number): number => (cff ? code : (this.glyphOf[code] ?? 0))
		const textOfCode = (code: number): string => (cff ? (this.cffText.get(code) ?? "") : (this.textOf[code] ?? ""))
		const used = new Uint8Array(sfnt.numGlyphs)

		for (const code of codes) used[glyphOfCode(code)] = 1

		const baseFont = `${subsetTag(this.postScriptName, used)}+${this.postScriptName}`
		const info = this.info
		const descriptorRef = sink.allocate()
		const programRef = sink.allocate()
		const cidFontRef = sink.allocate()
		const toUnicodeRef = sink.allocate()

		let subsetGlyphs: Uint16Array | undefined

		if (cff) {
			await sink.stream(programRef, { Subtype: pdfName("OpenType") }, sfnt.bytes, { compress: true })
		} else {
			// A font that forbids subsetting goes in whole (rebuilt as one font, for collections).
			if ((sfnt.fsType & NO_SUBSETTING) !== 0) used.fill(1)

			const subset = subsetTrueType(sfnt, used)

			subsetGlyphs = subset.glyphs
			await sink.stream(programRef, { Length1: subset.program.length }, subset.program, { compress: true })
		}

		sink.object(descriptorRef, {
			Type: pdfName("FontDescriptor"),
			FontName: pdfName(baseFont),
			Flags: info.flags,
			FontBBox: [...info.bbox],
			ItalicAngle: info.italicAngle,
			Ascent: info.ascent,
			Descent: info.descent,
			CapHeight: info.capHeight,
			StemV: info.stemV,
			[cff ? "FontFile3" : "FontFile2"]: programRef
		})

		let cidToGid: PdfRef | undefined

		if (!cff) {
			cidToGid = sink.allocate()

			const map = new Uint8Array(this.glyphOf.length * 2)

			for (let code = 1; code < this.glyphOf.length; code++) {
				const glyph = subsetGlyphs?.[this.glyphOf[code] ?? 0] ?? 0

				map[code * 2] = glyph >>> 8
				map[code * 2 + 1] = glyph & 0xff
			}

			await sink.stream(cidToGid, {}, map, { compress: true })
		}

		sink.object(cidFontRef, {
			Type: pdfName("Font"),
			Subtype: pdfName(cff ? "CIDFontType0" : "CIDFontType2"),
			BaseFont: pdfName(baseFont),
			CIDSystemInfo: { Registry: "Adobe", Ordering: "Identity", Supplement: 0 },
			FontDescriptor: descriptorRef,
			DW: 1000,
			W: pdfRaw(widthArray(codes, code => this.advance(code))),
			CIDToGIDMap: cidToGid
		})
		await sink.stream(toUnicodeRef, {}, toUnicodeCMap(codes, textOfCode), { compress: true })
		sink.object(ref, {
			Type: pdfName("Font"),
			Subtype: pdfName("Type0"),
			BaseFont: pdfName(baseFont),
			Encoding: pdfName("Identity-H"),
			DescendantFonts: [cidFontRef],
			ToUnicode: toUnicodeRef
		})
	}

	close(): void {
		this.closed = true
		this.glyphOf = [0]
		this.widthOf = [0]
		this.textOf = [""]
		this.byKey.clear()
		this.firstCode.clear()
		this.cffText.clear()
	}
}

// /W: runs of consecutive codes as `c [w1 w2 ...]`.
export function widthArray(codes: readonly number[], width: (code: number) => number): string {
	let out = "["

	for (let i = 0; i < codes.length;) {
		const first = codes[i] ?? 0
		let j = i + 1

		while (j < codes.length && codes[j] === (codes[j - 1] ?? 0) + 1) j++

		out += `${first} [`

		for (let k = i; k < j; k++) out += `${k > i ? " " : ""}${width(codes[k] ?? 0)}`

		out += "]"
		i = j
	}

	return `${out}]`
}

// ToUnicode CMap with one bfchar per code, at most 100 per block (§9.10.3, Adobe TN 5411).
export function toUnicodeCMap(codes: readonly number[], text: (code: number) => string): Uint8Array {
	const entries: string[] = []

	for (const code of codes) {
		const value = text(code)
		let hex = ""

		for (let k = 0; k < value.length; k++) hex += hex4(value.charCodeAt(k))

		entries.push(`<${hex4(code)}> <${hex}>`)
	}

	let body =
		"/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n" +
		"/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n"

	for (let i = 0; i < entries.length; i += 100) {
		const block = entries.slice(i, i + 100)

		body += `${block.length} beginbfchar\n${block.join("\n")}\nendbfchar\n`
	}

	body += "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n"

	const out = new Uint8Array(body.length)

	for (let i = 0; i < body.length; i++) out[i] = body.charCodeAt(i)

	return out
}

function hex4(value: number): string {
	return value.toString(16).toUpperCase().padStart(4, "0")
}

// Six capital letters derived from the face and the glyphs used (§9.6.4): the same subset gets the same tag.
export function subsetTag(postScriptName: string, used: Uint8Array): string {
	let hash = 0x811c9dc5

	for (let i = 0; i < postScriptName.length; i++) hash = Math.imul(hash ^ postScriptName.charCodeAt(i), 0x01000193)

	for (let g = 0; g < used.length; g++) {
		if (used[g] === 1) hash = Math.imul(hash ^ g, 0x01000193) ^ (g >>> 16)
	}

	let tag = ""
	let value = hash >>> 0

	for (let k = 0; k < 6; k++) {
		tag += String.fromCharCode(65 + (value % 26))
		value = Math.floor(value / 26) + k * 7919
	}

	return tag
}

function fontInfo(sfnt: SfntFont): PdfFontInfo {
	const view = new DataView(sfnt.bytes.buffer, sfnt.bytes.byteOffset, sfnt.bytes.byteLength)
	const scale = 1000 / sfnt.unitsPerEm
	const head = sfnt.table("head")
	const post = sfnt.table("post")
	const os2 = sfnt.table("OS/2")
	const bbox: [number, number, number, number] =
		head === undefined
			? [0, -200, 1000, 900]
			: [
					Math.round(view.getInt16(head.offset + 36) * scale),
					Math.round(view.getInt16(head.offset + 38) * scale),
					Math.round(view.getInt16(head.offset + 40) * scale),
					Math.round(view.getInt16(head.offset + 42) * scale)
				]
	const italicAngle = post !== undefined && post.length >= 16 ? view.getInt32(post.offset + 4) / 65536 : 0
	const fixedPitch = post !== undefined && post.length >= 16 && view.getUint32(post.offset + 12) !== 0
	const metrics = sfnt.metrics
	const ascender = metrics.useTypoMetrics ? metrics.typoAscender : metrics.hheaAscender
	const descender = metrics.useTypoMetrics ? metrics.typoDescender : metrics.hheaDescender
	// sFamilyClass 1-5 and 7 are serif classes (OS/2 IBM font class parameters).
	const familyClass = os2 !== undefined && os2.length >= 32 ? view.getInt16(os2.offset + 30) >> 8 : 0
	const serif = (familyClass >= 1 && familyClass <= 5) || familyClass === 7
	// Flags (§9.8.2): FixedPitch 1, Serif 2, Symbolic 4 (a CIDFont's glyphs are outside the standard Latin set),
	// Italic 64.
	const flags = (fixedPitch ? 1 : 0) | (serif ? 2 : 0) | 4 | (sfnt.italic || italicAngle !== 0 ? 64 : 0)
	// No stem width in TrueType: estimated from the weight class, as common practice does.
	const stemV = Math.round(10 + 220 * Math.max(0, (sfnt.weight - 50) / 900) ** 2)

	return {
		postScriptName: postScriptNameOf(sfnt),
		flags,
		bbox,
		italicAngle,
		ascent: Math.round(ascender * scale),
		descent: Math.round(descender * scale),
		capHeight: Math.round((sfnt.capHeight ?? ascender) * scale),
		stemV
	}
}

// The name table's PostScript name, else the family and style; reduced to the printable ASCII a PDF name
// and the PostScript name rules allow, at most 63 characters.
function postScriptNameOf(sfnt: SfntFont): string {
	const names = sfnt.names
	const raw = names.postScriptName ?? `${names.family ?? "Font"}-${names.subfamily ?? "Regular"}`
	let out = ""

	for (const char of raw) {
		const code = char.codePointAt(0) ?? 0

		if (code > 0x20 && code < 0x7f && !"[](){}<>/%#".includes(char)) out += char
	}

	return (out === "" ? "Font" : out).slice(0, 63)
}
