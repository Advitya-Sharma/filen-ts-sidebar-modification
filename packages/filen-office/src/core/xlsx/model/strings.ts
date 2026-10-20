import type { XlsxSheetStrings } from "./types"

// A sheet's own strings (inline strings, formula string results, errors, dates as written), each once.
export class SheetStrings implements XlsxSheetStrings {
	private readonly list: string[] = []
	private readonly index = new Map<string, number>()
	private chars = 0

	get length(): number {
		return this.list.length
	}

	text(index: number): string {
		return this.list[index] ?? ""
	}

	intern(text: string): number {
		let index = this.index.get(text)

		if (index === undefined) {
			index = this.list.length
			this.list.push(text)
			this.index.set(text, index)
			this.chars += text.length
		}

		return index
	}

	estimateBytes(): number {
		return this.chars * 2 + this.list.length * 48
	}
}

// Text as Excel shows it: the ST_Xstring escapes _xHHHH_ (ECMA-376 Part 1 §22.9.2.19) decoded; "_x005F_" escapes
// the underscore of a literal "_xHHHH_". Stored text keeps the escapes as written.
export function decodeXstring(text: string): string {
	if (text.indexOf("_x") < 0) return text

	return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)))
}
