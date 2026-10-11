// Styled math letters as Unicode Mathematical Alphanumeric Symbols (U+1D400–U+1D7FF): MathML Core keeps only
// mathvariant="normal", so bold, script, fraktur and the other OMML styles become distinct characters.

export const MathStyle = {
	NONE: -1,
	BOLD: 0,
	ITALIC: 1,
	BOLD_ITALIC: 2,
	SCRIPT: 3,
	BOLD_SCRIPT: 4,
	FRAKTUR: 5,
	DOUBLE_STRUCK: 6,
	BOLD_FRAKTUR: 7,
	SANS: 8,
	SANS_BOLD: 9,
	SANS_ITALIC: 10,
	SANS_BOLD_ITALIC: 11,
	MONOSPACE: 12
} as const

export type MathStyle = (typeof MathStyle)[keyof typeof MathStyle]

// Each style's Latin block is 52 letters (A–Z, a–z) from U+1D400, in MathStyle order.
const LATIN_BASE = 0x1d400

// Greek blocks (58 characters: capitals, nabla, lowercase, seven variants) by style; 0 = none.
const GREEK_BASE: Readonly<Record<number, number>> = {
	[MathStyle.BOLD]: 0x1d6a8,
	[MathStyle.ITALIC]: 0x1d6e2,
	[MathStyle.BOLD_ITALIC]: 0x1d71c,
	[MathStyle.SANS_BOLD]: 0x1d756,
	[MathStyle.SANS_BOLD_ITALIC]: 0x1d790
}

const DIGIT_BASE: Readonly<Record<number, number>> = {
	[MathStyle.BOLD]: 0x1d7ce,
	[MathStyle.DOUBLE_STRUCK]: 0x1d7d8,
	[MathStyle.SANS]: 0x1d7e2,
	[MathStyle.SANS_BOLD]: 0x1d7ec,
	[MathStyle.MONOSPACE]: 0x1d7f6
}

// Letters encoded earlier in Letterlike Symbols; their slots in the math block are reserved holes.
const HOLES: Readonly<Record<number, Readonly<Record<string, number>>>> = {
	[MathStyle.ITALIC]: { h: 0x210e },
	[MathStyle.SCRIPT]: {
		B: 0x212c,
		E: 0x2130,
		F: 0x2131,
		H: 0x210b,
		I: 0x2110,
		L: 0x2112,
		M: 0x2133,
		R: 0x211b,
		e: 0x212f,
		g: 0x210a,
		o: 0x2134
	},
	[MathStyle.FRAKTUR]: { C: 0x212d, H: 0x210c, I: 0x2111, R: 0x211c, Z: 0x2128 },
	[MathStyle.DOUBLE_STRUCK]: { C: 0x2102, H: 0x210d, N: 0x2115, P: 0x2119, Q: 0x211a, R: 0x211d, Z: 0x2124 }
}

// Position of a Greek letter inside a 58-character style block, or -1. Capitals U+0391–U+03A9 (with ϴ in
// the unassigned U+03A2 slot), then ∇, lowercase U+03B1–U+03C9, then ∂ ϵ ϑ ϰ ϕ ϱ ϖ.
function greekIndex(code: number): number {
	if (code >= 0x391 && code <= 0x3a9 && code !== 0x3a2) return code - 0x391

	switch (code) {
		case 0x3f4:
			return 17
		case 0x2207:
			return 25
		case 0x2202:
			return 51
		case 0x3f5:
			return 52
		case 0x3d1:
			return 53
		case 0x3f0:
			return 54
		case 0x3d5:
			return 55
		case 0x3f1:
			return 56
		case 0x3d6:
			return 57
	}

	return code >= 0x3b1 && code <= 0x3c9 ? 26 + code - 0x3b1 : -1
}

// The styled form of one code point; the code point itself when the style has no such character.
export function styleCodePoint(code: number, style: MathStyle): number {
	if (style === MathStyle.NONE) return code

	const upper = code >= 0x41 && code <= 0x5a
	const lower = code >= 0x61 && code <= 0x7a

	if (upper || lower) {
		const hole = HOLES[style]?.[String.fromCharCode(code)]

		if (hole !== undefined) return hole

		return LATIN_BASE + style * 52 + (upper ? code - 0x41 : 26 + code - 0x61)
	}

	if (code >= 0x30 && code <= 0x39) {
		const base = DIGIT_BASE[style]

		return base === undefined ? code : base + code - 0x30
	}

	const base = GREEK_BASE[style]

	if (base === undefined) return code

	const index = greekIndex(code)

	return index < 0 ? code : base + index
}

export function styleText(text: string, style: MathStyle): string {
	if (style === MathStyle.NONE) return text

	let out = ""

	for (const character of text) out += String.fromCodePoint(styleCodePoint(character.codePointAt(0) ?? 0, style))

	return out
}

// The style of m:rPr's m:scr (script) and m:sty (p, b, i, bi; italic when absent).
export function mathStyle(script: string | undefined, sty: string | undefined): MathStyle {
	const bold = sty === "b" || sty === "bi"
	const italic = sty === undefined || sty === "i" || sty === "bi"

	switch (script) {
		case "script":
			return bold ? MathStyle.BOLD_SCRIPT : MathStyle.SCRIPT
		case "fraktur":
			return bold ? MathStyle.BOLD_FRAKTUR : MathStyle.FRAKTUR
		case "double-struck":
			return MathStyle.DOUBLE_STRUCK
		case "sans-serif":
			return bold ? (italic ? MathStyle.SANS_BOLD_ITALIC : MathStyle.SANS_BOLD) : italic ? MathStyle.SANS_ITALIC : MathStyle.SANS
		case "monospace":
			return MathStyle.MONOSPACE
	}

	if (bold) return italic ? MathStyle.BOLD_ITALIC : MathStyle.BOLD

	// Roman italic and upright are left to MathML: single-letter mi is italic, mathvariant="normal" upright.
	return MathStyle.NONE
}
