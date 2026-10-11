import { execFileSync } from "node:child_process"
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { format, resolveConfig } from "prettier"

// Generates src/core/fonts/symbols.generated.ts: the Unicode characters the symbol fonts draw at their
// codes 0x20..0xFF. Wingdings, Wingdings 2, Wingdings 3 and Webdings come from the detailed tables of
// WG2 N4384 (L2/12-368), the mapping Unicode encoded them by; Symbol from Adobe's mapping on unicode.org.
// The PDF is read with Ghostscript's text device, which must be installed (`gs`). Both sources and their
// terms are listed in THIRD_PARTY_NOTICES.md.
//
//   pnpm run generate:symbols

const SOURCES = {
	"n4384.pdf": {
		url: "https://www.unicode.org/L2/L2012/12368-n4384.pdf",
		sha256: "15839e133b3ab10b84d68f41f25138d9dedbc94a8a6c2dbb155439d597084a0b"
	},
	"symbol.txt": {
		url: "https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt",
		sha256: "deb78ca840a429311939b9d165890873f71fb23ef223ceeb144a6c6d641a7e52"
	}
} as const

type Source = keyof typeof SOURCES

const CACHE = fileURLToPath(new URL("../../node_modules/.cache/filen-office-symbols/", import.meta.url))
const OUTPUT = fileURLToPath(new URL("../core/fonts/symbols.generated.ts", import.meta.url))

// N4384 numbers its sets 0 Webdings, 1 Wingdings, 2 Wingdings 2, 3 Wingdings 3.
const N4384_SETS = ["webdings", "wingdings", "wingdings2", "wingdings3"] as const
const FIRST = 0x20
const COUNT = 0xe0

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

async function fetchSource(name: Source): Promise<string> {
	const path = `${CACHE}${name}`
	const { url, sha256: pinned } = SOURCES[name]

	if (!existsSync(path) || sha256(new Uint8Array(readFileSync(path))) !== pinned) {
		const response = await fetch(url)

		if (!response.ok) throw new Error(`${name}: HTTP ${response.status}`)

		const bytes = new Uint8Array(await response.arrayBuffer())

		if (sha256(bytes) !== pinned) throw new Error(`${name} does not match its pinned SHA-256`)

		mkdirSync(CACHE, { recursive: true })
		writeFileSync(path, bytes)
	}

	return path
}

// Section 6 ("Detailed tables") lists every character once as "<set><decimal code> [glyph] <UCS> ...".
function n4384Tables(pdf: string): Record<(typeof N4384_SETS)[number], number[]> {
	const text = execFileSync("gs", ["-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=txtwrite", "-sOutputFile=-", pdf], {
		encoding: "utf8",
		maxBuffer: 64 * 1024 * 1024
	})
	const lines = text.split("\n")
	const start = lines.findIndex(line => /^\s*6\. Detailed tables/.test(line))
	const end = lines.findIndex(line => /^\s*7\. Source reference/.test(line))

	if (start < 0 || end < start) throw new Error("N4384: the detailed tables were not found")

	const tables = N4384_SETS.map(() => new Array<number>(COUNT).fill(0))

	for (const line of lines.slice(start, end)) {
		const match = /^\s+(\d)(\d{3})\s+(?:\S\s+)?([0-9A-F]{4,5})\s/.exec(line)

		if (match === null) continue

		const table = tables[Number(match[1])]
		const code = Number(match[2])

		if (table === undefined || code < FIRST || code >= FIRST + COUNT) throw new Error(`N4384: unexpected id ${match[1]}${match[2]}`)

		const value = parseInt(match[3] ?? "", 16)
		const previous = table[code - FIRST] ?? 0

		if (previous !== 0 && previous !== value) throw new Error(`N4384: ${match[1]}${match[2]} is mapped twice`)

		table[code - FIRST] = value
	}

	for (const table of tables) table[0] = 0x20

	const counts = tables.map(table => table.filter(value => value !== 0).length)

	// 867 characters, less the one not encoded (w-1255), plus each set's space.
	if (counts.reduce((sum, count) => sum + count, 0) !== 866 + 4) throw new Error(`N4384: unexpected counts ${counts.join(", ")}`)

	return { webdings: tables[0] ?? [], wingdings: tables[1] ?? [], wingdings2: tables[2] ?? [], wingdings3: tables[3] ?? [] }
}

// Adobe maps a few codes twice (Delta as U+0394 and U+2206, mu as U+00B5 and U+03BC ...): the Greek
// letter wins, else the first line. Private-use targets have no character, except the serif and sans
// designs of the registered, copyright and trademark signs, which Unicode encodes once.
const ADOBE_DESIGNS: Readonly<Record<string, number>> = {
	registerserif: 0xae,
	registersans: 0xae,
	copyrightserif: 0xa9,
	copyrightsans: 0xa9,
	trademarkserif: 0x2122,
	trademarksans: 0x2122
}

function adobeSymbol(path: string): number[] {
	const table = new Array<number>(COUNT).fill(0)

	for (const line of readFileSync(path, "utf8").split("\n")) {
		const match = /^([0-9A-F]{4})\t([0-9A-F]{2})\t#[^#]*#\s*(\S+)/.exec(line)

		if (match === null) continue

		const mapped = parseInt(match[1] ?? "", 16)
		const code = parseInt(match[2] ?? "", 16)
		const value = mapped >= 0xe000 && mapped <= 0xf8ff ? (ADOBE_DESIGNS[match[3] ?? ""] ?? 0) : mapped

		if (code < FIRST || value === 0) continue

		const previous = table[code - FIRST] ?? 0
		const greek = (c: number): boolean => c >= 0x370 && c <= 0x3ff

		if (previous === 0 || (greek(value) && !greek(previous))) table[code - FIRST] = value
	}

	return table
}

async function main(): Promise<void> {
	const tables = { symbol: adobeSymbol(await fetchSource("symbol.txt")), ...n4384Tables(await fetchSource("n4384.pdf")) }
	const rows = Object.entries(tables).map(([name, table]) => `\t${name}: [${table.map(value => `0x${value.toString(16)}`).join(", ")}]`)
	const source = [
		"// Generated by `pnpm run generate:symbols` from WG2 N4384 and Adobe's Symbol mapping. Do not edit.",
		"// Sources, licences and the Unicode, Inc. notice: THIRD_PARTY_NOTICES.md (Unicode symbol-font mappings).",
		"",
		`// The character each symbol font draws at codes 0x${FIRST.toString(16)}..0xff, 0 where it draws none.`,
		"export const SYMBOL_TABLES = {",
		rows.join(",\n"),
		"} as const",
		""
	].join("\n")

	writeFileSync(OUTPUT, await format(source, { ...(await resolveConfig(OUTPUT)), filepath: OUTPUT }))
}

await main()
