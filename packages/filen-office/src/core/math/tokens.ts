// Portions derived from SuperDoc 1.46.3 (packages/layout-engine/painters/dom/src/features/math/converters/math-run.ts),
// copyright Harbour and the SuperDoc contributors, licensed under the GNU Affero General Public License v3.0.
// Modified for @filen/office.

// Splitting math run text into MathML tokens, as Word's own OMML-to-MathML conversion does: digits (with
// one decimal point between digits) form one mn, each operator character is its own mo, every other
// character its own mi. Spaces become mspace widths.

export const TokenKind = { IDENTIFIER: 0, NUMBER: 1, OPERATOR: 2, SPACE: 3 } as const

export type TokenKind = (typeof TokenKind)[keyof typeof TokenKind]

export interface Token {
	readonly kind: TokenKind
	readonly text: string
	// SPACE: width in em.
	readonly width: number
}

const ASCII_OPERATORS = new Set("+-=<>()[]{}|/\\,.;:!?~^_*'\"%&@#$`")

// Characters inside the operator blocks that read as identifiers: ∂ ∅ ∆ ∇ ∞.
const SYMBOL_IDENTIFIERS = new Set([0x2202, 0x2205, 0x2206, 0x2207, 0x221e])

function isOperatorCode(code: number): boolean {
	if (code < 0x80) return ASCII_OPERATORS.has(String.fromCharCode(code))
	if (SYMBOL_IDENTIFIERS.has(code)) return false

	return (
		code === 0xac ||
		code === 0xb1 ||
		code === 0xb7 ||
		code === 0xd7 ||
		code === 0xf7 ||
		code === 0x2016 ||
		(code >= 0x2020 && code <= 0x2022) ||
		code === 0x2026 ||
		(code >= 0x2032 && code <= 0x2037) ||
		code === 0x2044 ||
		(code >= 0x2061 && code <= 0x2064) ||
		(code >= 0x2190 && code <= 0x21ff) ||
		(code >= 0x2200 && code <= 0x22ff) ||
		(code >= 0x2308 && code <= 0x230b) ||
		(code >= 0x2329 && code <= 0x232a) ||
		(code >= 0x23b4 && code <= 0x23b5) ||
		(code >= 0x23dc && code <= 0x23e1) ||
		(code >= 0x25a0 && code <= 0x25ff && code !== 0x25a1) ||
		(code >= 0x27c0 && code <= 0x27ff) ||
		(code >= 0x2900 && code <= 0x2aff) ||
		(code >= 0x3008 && code <= 0x3011)
	)
}

// Widths of the Unicode spaces in em (their nominal widths); U+0020 as Word's math spacing renders it.
const SPACE_WIDTHS: ReadonlyMap<number, number> = new Map([
	[0x20, 0.222],
	[0xa0, 0.222],
	[0x2000, 0.5],
	[0x2001, 1],
	[0x2002, 0.5],
	[0x2003, 1],
	[0x2004, 0.333],
	[0x2005, 0.222],
	[0x2006, 0.167],
	[0x2007, 0.5],
	[0x2008, 0.25],
	[0x2009, 0.167],
	[0x200a, 0.083],
	[0x202f, 0.167],
	[0x205f, 0.222],
	[0x3000, 1]
])

// Zero-width characters that carry no glyph (and would make an empty-looking token).
function isInvisible(code: number): boolean {
	return code === 0x200b || code === 0x200c || code === 0x200d || code === 0xfeff || code === 0x9 || code === 0xa || code === 0xd
}

function isDigit(code: number): boolean {
	return code >= 0x30 && code <= 0x39
}

// Combining marks attach to the token before them.
function isCombining(code: number): boolean {
	return (
		(code >= 0x300 && code <= 0x36f) ||
		(code >= 0x1ab0 && code <= 0x1aff) ||
		(code >= 0x1dc0 && code <= 0x1dff) ||
		(code >= 0x20d0 && code <= 0x20ff) ||
		(code >= 0xfe20 && code <= 0xfe2f)
	)
}

function isLetter(code: number): boolean {
	return (
		(code >= 0x41 && code <= 0x5a) ||
		(code >= 0x61 && code <= 0x7a) ||
		(code >= 0xc0 && !isOperatorCode(code) && !SPACE_WIDTHS.has(code))
	)
}

export interface TokenizeOptions {
	// m:lit: operator characters are ordinary characters.
	readonly literal?: boolean
	// Function names: consecutive letters stay one identifier ("sin", "lim").
	readonly groupLetters?: boolean
}

export function tokenize(text: string, options: TokenizeOptions = {}): Token[] {
	const out: Token[] = []
	const literal = options.literal === true
	const groupLetters = options.groupLetters === true
	let i = 0

	while (i < text.length) {
		const code = text.codePointAt(i) ?? 0
		const size = code > 0xffff ? 2 : 1

		if (isInvisible(code)) {
			i += size

			continue
		}

		const space = SPACE_WIDTHS.get(code)

		if (space !== undefined) {
			const last = out[out.length - 1]

			// Runs of spaces add up into one mspace.
			if (last !== undefined && last.kind === TokenKind.SPACE) {
				out[out.length - 1] = { kind: TokenKind.SPACE, text: "", width: Math.round((last.width + space) * 1000) / 1000 }
			} else {
				out.push({ kind: TokenKind.SPACE, text: "", width: space })
			}

			i += size

			continue
		}

		if (isCombining(code)) {
			const last = out[out.length - 1]

			if (last !== undefined && last.kind !== TokenKind.SPACE) {
				out[out.length - 1] = { kind: last.kind, text: last.text + String.fromCodePoint(code), width: 0 }
			} else {
				out.push({ kind: TokenKind.IDENTIFIER, text: String.fromCodePoint(code), width: 0 })
			}

			i += size

			continue
		}

		if (isDigit(code)) {
			let end = i + 1
			let dot = false

			while (end < text.length) {
				const next = text.charCodeAt(end)

				if (isDigit(next)) {
					end++
				} else if (next === 0x2e && !dot && isDigit(text.charCodeAt(end + 1))) {
					dot = true
					end++
				} else {
					break
				}
			}

			out.push({ kind: TokenKind.NUMBER, text: text.slice(i, end), width: 0 })
			i = end

			continue
		}

		if (!literal && isOperatorCode(code)) {
			out.push({ kind: TokenKind.OPERATOR, text: String.fromCodePoint(code), width: 0 })
			i += size

			continue
		}

		let end = i + size

		if (groupLetters && isLetter(code)) {
			while (end < text.length) {
				const next = text.codePointAt(end) ?? 0

				if (!isLetter(next) && !isCombining(next)) break

				end += next > 0xffff ? 2 : 1
			}
		}

		out.push({ kind: TokenKind.IDENTIFIER, text: text.slice(i, end), width: 0 })
		i = end
	}

	return out
}
