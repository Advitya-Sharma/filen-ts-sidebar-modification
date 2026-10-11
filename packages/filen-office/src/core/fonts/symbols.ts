import { SYMBOL_TABLES } from "./symbols.generated"

export type SymbolTable = readonly number[]

const TABLES: ReadonlyMap<string, SymbolTable> = new Map<string, SymbolTable>([
	["symbol", SYMBOL_TABLES.symbol],
	["wingdings", SYMBOL_TABLES.wingdings],
	["wingdings 2", SYMBOL_TABLES.wingdings2],
	["wingdings 3", SYMBOL_TABLES.wingdings3],
	["webdings", SYMBOL_TABLES.webdings]
])

// The code table of a symbol font (Symbol, Wingdings, Wingdings 2, Wingdings 3, Webdings), by name.
export function symbolTable(family: string): SymbolTable | undefined {
	return TABLES.get(family.trim().toLowerCase())
}

// The Unicode character a symbol font draws for `code`: its code 0x20..0xFF, or that code in the private
// use area (U+F020..U+F0FF), where Word writes symbol-font characters. 0 when the font draws none there.
export function symbolToUnicode(table: SymbolTable, code: number): number {
	const index = (code >= 0xf020 && code <= 0xf0ff ? code - 0xf000 : code) - 0x20

	return index >= 0 && index < 0xe0 ? (table[index] ?? 0) : 0
}

// `text` with each character a symbol font draws replaced by its Unicode character; others are kept.
export function symbolText(table: SymbolTable, text: string): string {
	let out: string | undefined

	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)
		const mapped = symbolToUnicode(table, code)

		if (mapped === 0 || mapped === code) {
			if (out !== undefined) out += text[i] ?? ""
		} else {
			out = (out ?? text.slice(0, i)) + String.fromCodePoint(mapped)
		}
	}

	return out ?? text
}
