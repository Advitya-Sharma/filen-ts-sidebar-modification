import { LINE_METRICS, symbolTable, symbolToUnicode, type SymbolTable } from "../fonts/index"
import { MT_EXTRA } from "./mtextra"

// Fonts and strings of metafile text: LOGFONT fields, the ANSI code pages of 8-bit text, and symbol fonts
// mapped to the Unicode characters they draw (so a Symbol-font α in an equation shows as α, not "a").

export const SYMBOL_CHARSET = 2

export interface MetafileFont {
	// LOGFONT height in logical units: negative is the em height, positive the cell height, 0 the default.
	readonly height: number
	readonly width: number
	// Tenths of a degree, counter-clockwise.
	readonly escapement: number
	readonly weight: number
	readonly italic: boolean
	readonly underline: boolean
	readonly strike: boolean
	readonly charset: number
	readonly face: string
	// The code table of a symbol font, which maps its codes to Unicode; undefined for text fonts.
	readonly symbols: SymbolTable | undefined
	// The face's ascent and descent as fractions of the em.
	readonly ascent: number
	readonly descent: number
}

// Symbol with the pieces equation editors build tall brackets, braces and integrals from (Adobe's table
// leaves them in the private use area): factual mappings from Apple's Mac OS Symbol table, version c02
// (© Apple Computer, Inc.; THIRD_PARTY_NOTICES.md).
const SYMBOL_WITH_PIECES: SymbolTable | undefined = (() => {
	const base = symbolTable("symbol")

	if (base === undefined) return undefined

	const table = base.slice()
	const pieces: readonly (readonly [number, number])[] = [
		[0xbd, 0x23d0],
		[0xbe, 0x23af],
		[0xe6, 0x239b],
		[0xe7, 0x239c],
		[0xe8, 0x239d],
		[0xe9, 0x23a1],
		[0xea, 0x23a2],
		[0xeb, 0x23a3],
		[0xec, 0x23a7],
		[0xed, 0x23a8],
		[0xee, 0x23a9],
		[0xef, 0x23aa],
		[0xf4, 0x23ae],
		[0xf6, 0x239e],
		[0xf7, 0x239f],
		[0xf8, 0x23a0],
		[0xf9, 0x23a4],
		[0xfa, 0x23a5],
		[0xfb, 0x23a6],
		[0xfc, 0x23ab],
		[0xfd, 0x23ac],
		[0xfe, 0x23ad]
	]

	for (const [code, unicode] of pieces) if ((table[code - 0x20] ?? 0) === 0) table[code - 0x20] = unicode

	return table
})()

// Advance widths (em fractions) of those pieces in Symbol, from Adobe's Symbol font metrics (the Core 14
// AFM files): a piece drawn without advances is centred in its Symbol cell, so the wider pieces of a
// substitute face do not run into the characters they enclose.
export function pieceAdvance(code: number): number {
	if ((code >= 0x239b && code <= 0x23a6) || code === 0x23a2) return 0.384
	if (code >= 0x23a7 && code <= 0x23ad) return 0.494
	if (code === 0x2320 || code === 0x2321 || code === 0x23ae) return 0.686
	if (code === 0x23d0) return 0.603
	if (code === 0x23af) return 1

	return 0
}

// The symbol table for a face: the shared symbol fonts, MT Extra, or undefined (a symbol-charset face the
// tables do not know keeps its codes).
export function faceSymbols(face: string): SymbolTable | undefined {
	const name = face.trim().toLowerCase()

	if (name === "symbol") return SYMBOL_WITH_PIECES

	return symbolTable(face) ?? (name === "mt extra" ? MT_EXTRA : undefined)
}

function fraction(units: number, unitsPerEm: number, fallback: number): number {
	const value = units / unitsPerEm

	return Number.isFinite(value) && value > 0 && value < 3 ? value : fallback
}

export function faceMetrics(face: string): { ascent: number; descent: number } {
	const entry = LINE_METRICS.get(face.trim().toLowerCase())

	if (entry !== undefined && entry.basis !== "substitute") {
		const metrics = entry.metrics

		return {
			ascent: fraction(metrics.winAscent, metrics.unitsPerEm, 0.9),
			descent: fraction(metrics.winDescent, metrics.unitsPerEm, 0.22)
		}
	}

	return face.trim().toLowerCase() === "symbol" ? { ascent: 1.005, descent: 0.22 } : { ascent: 0.9, descent: 0.22 }
}

export function metafileFont(
	height: number,
	width: number,
	escapement: number,
	weight: number,
	italic: boolean,
	underline: boolean,
	strike: boolean,
	charset: number,
	face: string
): MetafileFont {
	const metrics = faceMetrics(face)

	return {
		height,
		width,
		escapement,
		weight,
		italic,
		underline,
		strike,
		charset,
		face,
		symbols: faceSymbols(face),
		ascent: metrics.ascent,
		descent: metrics.descent
	}
}

export const DEFAULT_FONT: MetafileFont = metafileFont(0, 0, 0, 400, false, false, false, 0, "")

// The em height in logical units: a positive (cell) height includes the internal leading.
export function emHeight(font: MetafileFont, fallback: number): number {
	const height = font.height

	if (height < 0) return -height
	if (height > 0) return height / (font.ascent + font.descent)

	return fallback
}

// Windows-1252 at 0x80..0x9F (the rest of the page is Latin-1); 0 where it defines nothing.
const CP1252 = [
	0x20ac, 0, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0, 0x17d, 0, 0, 0x2018, 0x2019, 0x201c,
	0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0, 0x17e, 0x178
]

// WHATWG encoding labels of the LOGFONT charsets other than ANSI and symbol.
const CHARSET_ENCODINGS: ReadonlyMap<number, string> = new Map([
	[128, "shift_jis"],
	[129, "euc-kr"],
	[130, "euc-kr"],
	[134, "gbk"],
	[136, "big5"],
	[161, "windows-1253"],
	[162, "windows-1254"],
	[163, "windows-1258"],
	[177, "windows-1255"],
	[178, "windows-1256"],
	[186, "windows-1257"],
	[204, "windows-1251"],
	[222, "windows-874"],
	[238, "windows-1250"]
])

interface Decoder {
	decode(input: Uint8Array): string
}

type DecoderConstructor = new (label: string) => Decoder

// Hermes has no TextDecoder: there, text in those code pages falls back to Windows-1252.
const decoders = new Map<string, Decoder | null>()

function decoderFor(label: string): Decoder | null {
	const cached = decoders.get(label)

	if (cached !== undefined) return cached

	const constructor = (globalThis as { TextDecoder?: DecoderConstructor }).TextDecoder
	let decoder: Decoder | null = null

	if (typeof constructor === "function") {
		try {
			decoder = new constructor(label)
		} catch {
			decoder = null
		}
	}

	decoders.set(label, decoder)

	return decoder
}

// 8-bit metafile text as UTF-16. Symbol-charset text keeps its codes (the symbol table maps them).
export function decodeAnsi(bytes: Uint8Array, charset: number): string {
	const label = CHARSET_ENCODINGS.get(charset)

	if (label !== undefined) {
		const decoder = decoderFor(label)

		if (decoder !== null) return decoder.decode(bytes)
	}

	let out = ""

	for (let start = 0; start < bytes.length; start += 4096) {
		const end = Math.min(bytes.length, start + 4096)
		const units = new Array<number>(end - start)

		for (let i = start; i < end; i++) {
			const byte = bytes[i] ?? 0

			units[i - start] = charset !== SYMBOL_CHARSET && byte >= 0x80 && byte < 0xa0 ? CP1252[byte - 0x80] || byte : byte
		}

		out += String.fromCharCode(...units)
	}

	return out
}

// The character a font draws for a code point: symbol fonts through their table (0 where it draws nothing
// known; MT Extra's unmapped pieces and control codes are dropped), text fonts as is.
export function drawnCodePoint(font: MetafileFont, code: number): number {
	if (code < 0x20) return 0

	const table = font.symbols

	if (table === undefined) return code

	const mapped = symbolToUnicode(table, code)

	if (mapped !== 0) return mapped

	// Codes outside the symbol range (real Unicode text in a symbol-named font) are kept.
	return (code >= 0x20 && code <= 0xff) || (code >= 0xf020 && code <= 0xf0ff) ? 0 : code
}

// The face text is shown in: symbol fonts mapped to Unicode draw from a text face of the same style.
export function displayFace(font: MetafileFont): string {
	if (font.symbols === undefined) return font.face

	const name = font.face.trim().toLowerCase()

	return name === "symbol" || name === "mt extra" ? "Times New Roman" : font.face
}
