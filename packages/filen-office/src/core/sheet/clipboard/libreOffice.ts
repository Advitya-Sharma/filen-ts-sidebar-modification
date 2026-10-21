// LibreOffice's HTML values and number formats, written in the syntax of its locale, read as a workbook holds them.
// LibreOffice writes sdval with the locale's decimal separator and sdnum as "<language>;<format language>;<code>"
// (svtools HTMLOutFuncs::CreateTableDataOptionsValNum), the code in the format language's syntax: its decimal and
// group separators, its name for General, and for some languages its own keyword letters
// (svl ImpSvNumberformatScan::SetDependentKeywords); format language 0 is the system locale's. Its own HTML import
// reads them the same way (sfx2 SfxHTMLParser::GetTableDataOptionsValNum). Separators and General's name come from
// its locale data (i18npool/source/localedata/data, 7.5), by LCID ([MS-LCID]). The separator, General and keyword
// tables below are taken from LibreOffice (MPL-2.0, incorporating Apache-2.0 work); see THIRD_PARTY_NOTICES.md.

import { MAX_NUMFMT_LENGTH } from "../numfmt/edit"

interface LocaleSyntax {
	readonly decimal: string
	readonly group: string
	readonly general: string
	// What NNNN writes after the day of the week's name.
	readonly weekday: string
}

const EN_US: LocaleSyntax = { decimal: ".", group: ",", general: "General", weekday: ", " }

// [decimal, group, General's name, the day of the week's separator, LCIDs (hex)]; every other LCID writes en-US's.
const PROFILES: readonly (readonly [string, string, string, string, string])[] = [
	[
		",",
		"\u00a0",
		"Standard",
		" ",
		"14 40c 40e 414 415 41d 423 428 42c 430 431 432 433 434 435 436 440 443 444 46c 611 62b 62c 635 699 814 c0c 140c 180c 1c09 280c 300c 340c 800c 840c 880c 8c0c"
	],
	[
		",",
		".",
		"General",
		", ",
		"41a 48d 631 63a 64a 64f 659 65b 65c 65d 65e 65f 660 661 662 664 665 678 679 67a 67c 67d 680 681 682 683 68d 68e 85f 141a 8277 827b 827e 827f 8284"
	],
	[
		".",
		",",
		"General",
		" ",
		"43c 451 452 45e 491 63b 809 83c 851 c51 1809 3009 3c09 4009 4409 8051 8809 9009 9809 9c09 a009 a409 a809 ac09"
	],
	[",", ".", "Standard", ", ", "403 407 408 424 42e 456 46e 612 623 624 630 63c 82e c07 c1a 1007 281a 301a 8c1a cc1a"],
	[",", ".", "Standard", " ", "406 40f 410 418 41f 421 438 46f 476 487 610 621 62a 633 695 6a7 6a8 80c 8076"],
	["\u066b", "\u066c", "General", "\u060c ", "401 801 c01 1001 1801 1c01 2401 2801 2c01 3001 3401 3801 3c01 4001"],
	[".", ",", "Standard", " ", "439 445 446 447 449 44a 44b 44c 44e 454 461 845 8009 8036"],
	[".", ",", "Estandar", " ", "80a 100a 180a 1c0a 280a 300a 440a 480a 4c0a 500a"],
	[",", "\u00a0", "General", ", ", "402 425 442 64b 675 694"],
	[",", ".", "Estandar", ", ", "c0a 200a 2c0a 380a 3c0a"],
	[".", "'", "Standard", ", ", "807 810 100c 1407"],
	[",", "\u00a0", "Standard", ", ", "405 419 41b"],
	[",", ".", "General", " ", "42a 453 b009"],
	[",", ".", "Estandar", " ", "240a 340a 400a"],
	[".", ",", "\u5e38\u89c4", ", ", "4 804"],
	[",", "\u00a0", "Yleinen", " ", "40b 81d"],
	[",", ".", "Standaard", " ", "413 813"],
	[",", ".", "Geral", ", ", "416 8016"],
	[",", "\u00a0", "Standard", "", "422 426"],
	[",", ".", "Bendras", " ", "427 827"],
	[".", "\u00a0", "General", ", ", "42b 437"],
	[",", "\u00a0", "General", " ", "43f 62d"],
	[",", "\u00a0", "Skoueriek", " ", "47e 629"],
	[".", ",", "General", ",\u00a0", "637 63e"],
	[",", "\u00a0", "Geral", ", ", "816 8416"],
	[".", ",", "Standard", ", ", "411"],
	[".", ",", "Standard", "", "41e"],
	["\u066b", "\u066c", "\u0639\u0627\u062f\u06cc", " ", "429"],
	[",", ".", "Estandarra", ", ", "42d"],
	[",", ".", "\u041e\u043f\u0448\u0442\u043e", ", ", "42f"],
	[".", ",", "General", "\u104a ", "455"],
	[",", ".", "General", ",\u00a0", "462"],
	[".", "\u066c", "General", " \u060c", "465"],
	[",", "\u00a0", "General", ",\u00a0", "639"],
	[".", ",", "Kadaywan", ", ", "6a3"],
	[",", ".", "Standard", " a-i ", "6a6"],
	["\u066b", "\u00a0", "General", "\u060c ", "1401"],
	[",", "\u00a0", "Estandar", ", ", "140a"],
	[".", ",", "General", "\u060c ", "2001"]
]

let syntaxes: Map<number, LocaleSyntax> | null = null

function syntaxOf(lcid: number): LocaleSyntax {
	if (syntaxes === null) {
		const built = new Map<number, LocaleSyntax>()

		for (const [decimal, group, general, weekday, lcids] of PROFILES) {
			const syntax = { decimal, group, general, weekday }

			for (const id of lcids.split(" ")) built.set(Number.parseInt(id, 16), syntax)
		}

		syntaxes = built
	}

	return syntaxes.get(lcid) ?? EN_US
}

// The languages whose format codes use their own keyword letters, each mapped to the English one.
const GERMAN = [0x0407, 0x0807, 0x0c07, 0x1007, 0x1407]
const SPANISH = [0x040a, 0x080a, 0x0c0a, 0x100a, 0x140a, 0x180a, 0x1c0a, 0x200a, 0x240a, 0x280a, 0x2c0a, 0x300a, 0x340a, 0x380a]
const KEYWORD_LETTERS: readonly (readonly [readonly number[], Readonly<Record<string, string>>])[] = [
	[GERMAN, { T: "D", J: "Y" }],
	[[0x0413, 0x0813], { J: "Y", U: "H" }],
	[[0x040c, 0x080c, 0x0c0c, 0x100c, 0x140c, 0x180c], { J: "D", A: "Y", O: "A" }],
	[[0x0410, 0x0810], { G: "D", A: "Y", X: "G", O: "A" }],
	[[...SPANISH, 0x3c0a, 0x400a, 0x440a, 0x480a, 0x4c0a, 0x500a, 0x0816, 0x0416], { A: "Y", O: "A" }],
	[[0x040b], { P: "D", K: "M", V: "Y", T: "H" }],
	[[0x041d, 0x081d, 0x0406, 0x0014, 0x0414, 0x0814], { T: "H" }]
]

function lettersOf(lcid: number): Readonly<Record<string, string>> {
	for (const [lcids, letters] of KEYWORD_LETTERS) if (lcids.includes(lcid)) return letters

	return {}
}

// The colours Excel names, and German's names for them.
const COLORS = new Set(["BLACK", "BLUE", "CYAN", "GREEN", "MAGENTA", "RED", "WHITE", "YELLOW"])
const GERMAN_COLORS: Readonly<Record<string, string>> = {
	SCHWARZ: "Black",
	BLAU: "Blue",
	CYAN: "Cyan",
	GRÜN: "Green",
	MAGENTA: "Magenta",
	ROT: "Red",
	GELB: "Yellow",
	WEISS: "White"
}

export interface LibreOfficeNumber {
	// The cell's value, null when sdval does not read.
	readonly value: number | null
	readonly boolean: boolean
	// The format as an Excel code; undefined when there is none, or none a workbook can take.
	readonly format: string | undefined
}

function isSpaceLike(char: string): boolean {
	return char === " " || char === " " || char === " "
}

// sdval with the locale's decimal separator (never grouped), or "." where the system's locale writes that.
function valueOf(text: string, syntax: LocaleSyntax): { value: number | null; decimal: string | null } {
	const trimmed = text.trim()

	if (trimmed === "") return { value: null, decimal: null }

	const decimal = trimmed.includes(syntax.decimal) ? syntax.decimal : trimmed.includes(".") ? "." : trimmed.includes(",") ? "," : null
	const value = Number(decimal === null ? trimmed : trimmed.replace(decimal, "."))

	return { value: Number.isFinite(value) ? value : null, decimal }
}

// A cell's sdval and sdnum read; sdnum undefined for a cell without one (its sdval read as en-US's or with a comma).
export function libreOfficeNumber(sdval: string | undefined, sdnum: string | undefined): LibreOfficeNumber {
	const fields = sdnum?.split(";") ?? []
	const language = Number.parseInt(fields[0] ?? "", 10)
	const syntax = syntaxOf(Number.isFinite(language) ? language : 0)
	const read = sdval === undefined ? { value: null, decimal: null } : valueOf(sdval, syntax)
	const formatLanguage = Number.parseInt(fields[1] ?? "", 10)
	const codeLcid = Number.isFinite(formatLanguage) && formatLanguage !== 0 ? formatLanguage : Number.isFinite(language) ? language : 0
	const code = fields.length > 2 ? fields.slice(2).join(";") : ""
	const codeSyntax = syntaxOf(codeLcid)
	const boolean = code !== "" && code.toUpperCase() === (GERMAN.includes(codeLcid) ? "LOGISCH" : "BOOLEAN")

	// The system's locale (format language 0) writes the value too: one with another decimal separator is no locale
	// known here, and its code is read as no format.
	const known = formatLanguage !== 0 || read.decimal === null || read.decimal === codeSyntax.decimal

	// A code past Excel's 255 characters is none a workbook holds, whatever it converts to.
	return {
		value: read.value,
		boolean,
		format: boolean || code === "" || code.length > MAX_NUMFMT_LENGTH || !known ? undefined : excelCode(code, codeLcid, codeSyntax)
	}
}

type Piece =
	| { readonly kind: "literal"; readonly text: string }
	| { readonly kind: "decimal" }
	| { readonly kind: "group"; readonly char: string }
	// A "." or "," that is neither separator in the language.
	| { readonly kind: "mark"; readonly char: string }
	| { readonly kind: "placeholder"; readonly char: string }
	| { readonly kind: "letters"; readonly letter: string; readonly count: number }
	| { readonly kind: "date"; readonly text: string }

const DATE_LETTERS = new Set(["D", "M", "Y", "H", "S", "A", "N", "G", "E"])

function letterRun(code: string, at: number): number {
	const upper = code.charAt(at).toUpperCase()
	let end = at + 1

	while (end < code.length && code.charAt(end).toUpperCase() === upper) end++

	return end
}

function isLetter(char: string): boolean {
	return char.toUpperCase() !== char.toLowerCase()
}

// Compared in place, a character at a time: nothing sliced or case-mapped past the word's length.
function startsWithIgnoringCase(text: string, at: number, word: string): boolean {
	if (word === "" || at + word.length > text.length) return false

	for (let k = 0; k < word.length; k++) {
		const char = text.charAt(at + k)
		const expected = word.charAt(k)

		if (char !== expected && char.toUpperCase() !== expected.toUpperCase()) return false
	}

	return true
}

// A bracket's content in Excel's syntax, or null for one a workbook has no meaning for.
function bracket(content: string, syntax: LocaleSyntax, letters: Readonly<Record<string, string>>, german: boolean): string | null {
	const first = content.charAt(0)

	if (first === "<" || first === ">" || first === "=") return syntax.decimal === "." ? content : content.replace(syntax.decimal, ".")
	if (first === "$") return content

	const upper = content.toUpperCase()

	if (COLORS.has(upper) || /^COLOR\d{1,2}$/.test(upper)) return content
	if (german) {
		const color = GERMAN_COLORS[upper]

		if (color !== undefined && Object.hasOwn(GERMAN_COLORS, upper)) return color

		const indexed = /^FARBE(\d{1,2})$/.exec(upper)?.[1]

		if (indexed !== undefined) return `Color${indexed}`
	}

	// Elapsed time: [HH], [MM], [SS] in the language's letters.
	if (content !== "" && letterRun(content, 0) === content.length) {
		const letter = letters[upper.charAt(0)] ?? upper.charAt(0)

		if (letter === "H" || letter === "M" || letter === "S") return letter.toLowerCase().repeat(content.length)
	}

	return null
}

// One section of a code in pieces; null when it holds what a workbook cannot.
function pieces(section: string, lcid: number, syntax: LocaleSyntax): Piece[] | null {
	const letters = lettersOf(lcid)
	const german = GERMAN.includes(lcid)
	const out: Piece[] = []
	let i = 0

	while (i < section.length) {
		const char = section.charAt(i)

		if (char === '"') {
			const close = section.indexOf('"', i + 1)
			const end = close < 0 ? section.length : close + 1

			out.push({ kind: "literal", text: section.slice(i, end) })
			i = end
		} else if (char === "\\" || char === "_" || char === "*") {
			out.push({ kind: "literal", text: section.slice(i, i + 2) })
			i += 2
		} else if (char === "[") {
			const close = section.indexOf("]", i + 1)

			if (close < 0) return null

			const content = bracket(section.slice(i + 1, close), syntax, letters, german)

			if (content === null) return null

			const elapsed = /^[hms]+$/.test(content)

			out.push(elapsed ? { kind: "date", text: `[${content}]` } : { kind: "literal", text: `[${content}]` })
			i = close + 1
		} else if (startsWithIgnoringCase(section, i, "AM/PM")) {
			out.push({ kind: "date", text: section.slice(i, i + 5) })
			i += 5
		} else if (startsWithIgnoringCase(section, i, "A/P")) {
			out.push({ kind: "date", text: section.slice(i, i + 3) })
			i += 3
		} else if (startsWithIgnoringCase(section, i, syntax.general) || startsWithIgnoringCase(section, i, "General")) {
			const length = startsWithIgnoringCase(section, i, syntax.general) ? syntax.general.length : 7

			out.push({ kind: "literal", text: "General" })
			i += length
		} else if ((char === "E" || char === "e") && (section.charAt(i + 1) === "+" || section.charAt(i + 1) === "-")) {
			out.push({ kind: "literal", text: `E${section.charAt(i + 1)}` })
			i += 2
		} else if (char === syntax.decimal) {
			out.push({ kind: "decimal" })
			i++
		} else if (char === syntax.group || (isSpaceLike(syntax.group) && isSpaceLike(char))) {
			out.push({ kind: "group", char })
			i++
		} else if (char === "0" || char === "#" || char === "?") {
			out.push({ kind: "placeholder", char })
			i++
		} else if (isLetter(char)) {
			const end = letterRun(section, i)
			const upper = char.toUpperCase()

			out.push({ kind: "letters", letter: letters[upper] ?? upper, count: end - i })
			i = end
		} else {
			out.push(char === "." || char === "," ? { kind: "mark", char } : { kind: "literal", text: char })
			i++
		}
	}

	return out
}

function dateLetters(letter: string, count: number, syntax: LocaleSyntax): string | null {
	switch (letter) {
		case "D":
		case "M":
		case "Y":
		case "H":
		case "S":
		case "G":
		case "E":
			return letter.toLowerCase().repeat(count)
		// The day of the week: AAA/AAAA, NN/NNN.
		case "A":
			return count === 3 ? "ddd" : count === 4 ? "dddd" : null
		case "N":
			return count === 2 ? "ddd" : count === 3 ? "dddd" : count === 4 ? `dddd${quoted(syntax.weekday)}` : null
		default:
			return null
	}
}

function quoted(text: string): string {
	return text === "" || text.includes('"') ? "" : `"${text}"`
}

function isPlaceholder(piece: Piece | undefined): boolean {
	return piece?.kind === "placeholder"
}

// A separator or mark in Excel's syntax: in a date, written as it is (the separator of a fraction of a second "." and
// a "." that is not one escaped); in a number, the decimal point ".", grouping or scaling ",", and the others escaped.
function separator(parts: readonly Piece[], at: number, date: boolean, syntax: LocaleSyntax): string {
	const part = parts[at]
	const previous = parts[at - 1]
	const next = parts[at + 1]

	if (part === undefined || part.kind === "literal" || part.kind === "date" || part.kind === "placeholder" || part.kind === "letters")
		return ""

	const char = part.kind === "decimal" ? syntax.decimal : part.char

	if (date) {
		const afterSeconds = previous?.kind === "letters" && previous.letter === "S"

		if (part.kind === "decimal") return afterSeconds ? "." : char

		return char === "." && afterSeconds ? "\\." : char
	}

	if (part.kind === "decimal") return "."

	if (part.kind === "group") {
		// Grouping between placeholders, or scaling after one; a space anywhere else stays a space.
		const grouping =
			(isPlaceholder(previous) || previous?.kind === "group") && (isPlaceholder(next) || next?.kind === "group" || !isSpaceLike(char))

		if (grouping) return ","
		if (isSpaceLike(char)) return char
	}

	return `\\${char}`
}

// A code in LibreOffice's syntax for `lcid` as Excel writes it; undefined when it holds what a workbook cannot.
export function excelCode(code: string, lcid: number, syntax: LocaleSyntax = syntaxOf(lcid)): string | undefined {
	const sections: string[] = []
	let start = 0
	let quoted = false

	// Sections split at their unquoted, unescaped semicolons.
	for (let i = 0; i <= code.length; i++) {
		const char = code.charAt(i)

		if (i === code.length || (char === ";" && !quoted)) {
			sections.push(code.slice(start, i))
			start = i + 1
		} else if (char === '"') quoted = !quoted
		else if (char === "\\" && !quoted) i++
	}

	const out: string[] = []

	for (const section of sections) {
		const parts = pieces(section, lcid, syntax)

		if (parts === null) return undefined

		const date = parts.some(part => part.kind === "date" || (part.kind === "letters" && DATE_LETTERS.has(part.letter)))
		let text = ""

		for (let k = 0; k < parts.length; k++) {
			const part = parts[k]

			if (part === undefined) continue

			switch (part.kind) {
				case "literal":
				case "date":
					text += part.text
					break
				case "placeholder":
					text += part.char
					break
				case "letters": {
					const written = date ? dateLetters(part.letter, part.count, syntax) : null

					if (written === null) return undefined

					text += written
					break
				}
				case "decimal":
				case "group":
				case "mark":
					text += separator(parts, k, date, syntax)
					break
			}
		}

		out.push(text)
	}

	return out.join(";")
}
