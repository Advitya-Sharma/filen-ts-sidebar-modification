// Code pages of RTF text: the document's (\ansi, \mac, \pc, \pca, \ansicpgN) and each font's (\fcharsetN,
// superseded by \cpgN), decoded with the platform's TextDecoder (Microsoft RTF Specification 1.9.1, "Character
// Set", "Font Table"). Where the platform has no decoder for a code page, Windows-1252 stands in and the
// reader reports it; the source bytes are never lost.

// \fcharsetN → code page, from the specification's \fcharsetN table. 0 is "the document's code page", 42 Symbol.
const CHARSET_CODEPAGES: ReadonlyMap<number, number> = new Map([
	[0, 1252],
	[1, 0],
	[2, 42],
	[77, 10000],
	[78, 10001],
	[79, 10003],
	[80, 10008],
	[81, 10002],
	[83, 10005],
	[84, 10004],
	[85, 10006],
	[86, 10081],
	[87, 10021],
	[88, 10029],
	[89, 10007],
	[128, 932],
	[129, 949],
	[130, 1361],
	[134, 936],
	[136, 950],
	[161, 1253],
	[162, 1254],
	[163, 1258],
	[177, 1255],
	[178, 1256],
	[186, 1257],
	[204, 1251],
	[222, 874],
	[238, 1250],
	[254, 437],
	[255, 850]
])

export const SYMBOL_CODEPAGE = 42

// The code page a font's \fcharsetN implies; 0 for "the document's", undefined for charsets without one.
export function charsetCodepage(charset: number): number | undefined {
	return CHARSET_CODEPAGES.get(charset)
}

// Encoding Standard labels (https://encoding.spec.whatwg.org/) for the code pages the specification lists.
const LABELS: ReadonlyMap<number, string> = new Map([
	[708, "iso-8859-6"],
	[819, "windows-1252"],
	[866, "ibm866"],
	[874, "windows-874"],
	[932, "shift_jis"],
	[936, "gbk"],
	[949, "euc-kr"],
	[950, "big5"],
	[1250, "windows-1250"],
	[1251, "windows-1251"],
	[1252, "windows-1252"],
	[1253, "windows-1253"],
	[1254, "windows-1254"],
	[1255, "windows-1255"],
	[1256, "windows-1256"],
	[1257, "windows-1257"],
	[1258, "windows-1258"],
	[10000, "macintosh"],
	[10007, "x-mac-cyrillic"],
	[20866, "koi8-r"],
	[21866, "koi8-u"],
	[20932, "euc-jp"],
	[28591, "iso-8859-1"],
	[28592, "iso-8859-2"],
	[28593, "iso-8859-3"],
	[28594, "iso-8859-4"],
	[28595, "iso-8859-5"],
	[28596, "iso-8859-6"],
	[28597, "iso-8859-7"],
	[28598, "iso-8859-8"],
	[28599, "iso-8859-9"],
	[28603, "iso-8859-13"],
	[28605, "iso-8859-15"],
	[50220, "iso-2022-jp"],
	[51932, "euc-jp"],
	[51949, "euc-kr"],
	[54936, "gb18030"],
	[65001, "utf-8"]
])

// Windows-1252's 0x80..0x9F, for hosts without TextDecoder (Hermes): the rest of its range is Latin-1.
const CP1252_HIGH = [
	0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f, 0x90, 0x2018,
	0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178
]

interface Decoder {
	decode(input: Uint8Array): string
}

type DecoderConstructor = new (label: string) => Decoder

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

// Whether `codepage` decodes as itself here (false: Windows-1252 stands in).
export function codepageSupported(codepage: number): boolean {
	if (codepage === SYMBOL_CODEPAGE || codepage === 1252) return true

	const label = LABELS.get(codepage)

	return label !== undefined && decoderFor(label) !== null
}

function cp1252(bytes: Uint8Array, start: number, end: number): string {
	let out = ""

	for (let at = start; at < end; at += 4096) {
		const stop = Math.min(end, at + 4096)
		const units = new Array<number>(stop - at)

		for (let i = at; i < stop; i++) {
			const byte = bytes[i] ?? 0

			units[i - at] = byte >= 0x80 && byte < 0xa0 ? (CP1252_HIGH[byte - 0x80] ?? byte) : byte
		}

		out += String.fromCharCode(...units)
	}

	return out
}

// Bytes in a code page as UTF-16. Symbol-charset bytes become U+F020..U+F0FF, where Word keeps symbol-font
// characters (spec, "Unicode RTF"); layout maps them through the font's symbol table.
export function decodeCodepage(bytes: Uint8Array, start: number, end: number, codepage: number): string {
	if (end <= start) return ""

	if (codepage === SYMBOL_CODEPAGE) {
		let out = ""

		for (let i = start; i < end; i++) {
			const byte = bytes[i] ?? 0

			out += String.fromCharCode(byte >= 0x20 ? 0xf000 + byte : byte)
		}

		return out
	}

	let ascii = true

	for (let i = start; i < end; i++) {
		if ((bytes[i] ?? 0) >= 0x80) {
			ascii = false

			break
		}
	}

	if (ascii) return cp1252(bytes, start, end)

	if (codepage !== 1252) {
		const label = LABELS.get(codepage)
		const decoder = label === undefined ? null : decoderFor(label)

		if (decoder !== null) return decoder.decode(bytes.subarray(start, end))
	}

	return cp1252(bytes, start, end)
}
