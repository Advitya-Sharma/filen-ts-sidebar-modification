import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { format, resolveConfig } from "prettier"
import { BIDI_CLASS_NAMES, LINE_BREAK_CLASS_NAMES, WORD_BREAK_CLASS_NAMES, packRuns, packNumbers } from "../core/text/packed"

// Generates src/core/text/unicode.generated.ts and the sampled conformance fixtures from the Unicode
// Character Database files pinned below (version and SHA-256). Downloads go to a cache under node_modules;
// the full conformance files stay there for the opt-in full test run (FILEN_OFFICE_UNICODE_FULL=1).
//
//   pnpm exec tsx src/scripts/unicodeTables.ts [--offline]

export const UNICODE_VERSION = "18.0.0"

const BASE = `https://www.unicode.org/Public/${UNICODE_VERSION}/ucd/`

export const UCD_FILES = {
	"LineBreak.txt": "ae8cf1970c73f3f1a12d77852df96c4b2b1723ca8b388e36bb5739250c6849de",
	"EastAsianWidth.txt": "a0cf29eacd00cfcaec4381c6b7c281685f18dbb4e7ff82b4076ccb342ca839aa",
	"extracted/DerivedGeneralCategory.txt": "d6b151d2d40ee9b1876d26f417980f45ffae47b6055ccf7203cb31f07a030f94",
	"emoji/emoji-data.txt": "80d00f8e616a0ef27fd6b8de3b758c06383b5d917e2977709578e68baf733bf1",
	"extracted/DerivedBidiClass.txt": "d9e23222522551348ea1ccfbb4f62efbf98982afb95840f8959c08ed992c5607",
	"BidiBrackets.txt": "4b3b62e4a14b84ee752808c810c602534921c09a4a1bf78cfbee566d66c125b3",
	"UnicodeData.txt": "0736451de439ae7baf1425136617da495e09ee5afbe6e394374db7009ea08950",
	"auxiliary/LineBreakTest.txt": "fd9ff9eca411080085b41471bba79ae19d9db2e16924f9770b8509f1affd3f95",
	"BidiCharacterTest.txt": "045b24d2c8ab066951bd32fe8c6b4de34647f72b5b1c7df0265f24ab53573e01",
	"BidiTest.txt": "9af2f882a4ab50912e388f069a673b94eacd82fa6d07d20a3ff7f3c759e905aa",
	"auxiliary/WordBreakProperty.txt": "8dbfa17063e11084201f33c3e76d485d3b9166930c71db8e39ed1c9234171aec",
	"auxiliary/WordBreakTest.txt": "3dd70c071781276067c680d87303f60434adce7b067bd063194af374edad86a5"
} as const

export type UcdFile = keyof typeof UCD_FILES

export const CACHE = fileURLToPath(new URL(`../../node_modules/.cache/filen-office-unicode/${UNICODE_VERSION}/`, import.meta.url))
const OUTPUT = fileURLToPath(new URL("../core/text/unicode.generated.ts", import.meta.url))
export const FIXTURES = fileURLToPath(new URL("../tests/fixtures/unicode/", import.meta.url))

const MAX_CODE_POINT = 0x10ffff

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

function cachePath(file: UcdFile): string {
	return `${CACHE}${file.replaceAll("/", "_")}`
}

// A pinned UCD file from the cache, verified. Throws when it is missing or does not match its pin.
export function readCached(file: UcdFile): string {
	const bytes = new Uint8Array(readFileSync(cachePath(file)))

	if (sha256(bytes) !== UCD_FILES[file]) throw new Error(`${file} does not match its pinned SHA-256`)

	return new TextDecoder().decode(bytes)
}

async function download(file: UcdFile): Promise<void> {
	const path = cachePath(file)

	if (existsSync(path) && sha256(new Uint8Array(readFileSync(path))) === UCD_FILES[file]) return

	const response = await fetch(`${BASE}${file}`)

	if (!response.ok) throw new Error(`${file}: HTTP ${response.status}`)

	const bytes = new Uint8Array(await response.arrayBuffer())

	if (sha256(bytes) !== UCD_FILES[file]) throw new Error(`${file} does not match its pinned SHA-256`)

	mkdirSync(CACHE, { recursive: true })
	writeFileSync(path, bytes)
}

// Calls `entry` for every data line of a UCD property file (ranges inclusive), after its @missing defaults.
function forEachEntry(source: string, entry: (first: number, last: number, value: string) => void): void {
	const missing: [number, number, string][] = []
	const explicit: [number, number, string][] = []

	for (const raw of source.split("\n")) {
		const isMissing = raw.startsWith("# @missing:")
		const line = isMissing ? raw.slice("# @missing:".length) : raw.replace(/#.*/, "")

		if (line.trim() === "") continue

		const [range = "", value = ""] = line.split(";").map(part => part.trim())
		const [first = "", last = first] = range.split("..")

		;(isMissing ? missing : explicit).push([Number.parseInt(first, 16), Number.parseInt(last, 16), value])
	}

	for (const [first, last, value] of [...missing, ...explicit]) entry(first, last, value)
}

function propertyArray(source: string, map: (value: string) => number, initial: number): Uint8Array {
	const values = new Uint8Array(MAX_CODE_POINT + 1).fill(initial)

	forEachEntry(source, (first, last, value) => values.fill(map(value), first, last + 1))

	return values
}

const BIDI_LONG_NAMES: Record<string, string> = {
	Left_To_Right: "L",
	Right_To_Left: "R",
	Arabic_Letter: "AL",
	European_Number: "EN",
	European_Separator: "ES",
	European_Terminator: "ET",
	Arabic_Number: "AN",
	Common_Separator: "CS",
	Nonspacing_Mark: "NSM",
	Boundary_Neutral: "BN",
	Paragraph_Separator: "B",
	Segment_Separator: "S",
	White_Space: "WS",
	Other_Neutral: "ON",
	Left_To_Right_Embedding: "LRE",
	Left_To_Right_Override: "LRO",
	Right_To_Left_Embedding: "RLE",
	Right_To_Left_Override: "RLO",
	Pop_Directional_Format: "PDF",
	Left_To_Right_Isolate: "LRI",
	Right_To_Left_Isolate: "RLI",
	First_Strong_Isolate: "FSI",
	Pop_Directional_Isolate: "PDI"
}

function indexOf(names: readonly string[], name: string): number {
	const index = names.indexOf(name)

	if (index < 0) throw new Error(`Unknown property value ${name}`)

	return index
}

const LB = (name: string): number => indexOf(LINE_BREAK_CLASS_NAMES, name)

// Line_Break with LB1's default resolution (AI, SG, XX as AL; SA as CM for Mn/Mc, else AL), QU split by
// General_Category Pi/Pf, plus East_Asian_Width F/W/H (bit 6) and Extended_Pictographic unassigned
// (bit 7). CJ stays: strict or normal breaking resolves it at run time.
function lineBreakValues(): Uint8Array {
	const gc = propertyArray(readCached("extracted/DerivedGeneralCategory.txt"), value => indexOf(GC_NAMES, value), indexOf(GC_NAMES, "Cn"))
	const ea = propertyArray(readCached("EastAsianWidth.txt"), value => (value === "F" || value === "W" || value === "H" ? 1 : 0), 0)
	const pictographic = new Uint8Array(MAX_CODE_POINT + 1)

	forEachEntry(readCached("emoji/emoji-data.txt"), (first, last, value) => {
		if (value === "Extended_Pictographic") pictographic.fill(1, first, last + 1)
	})

	const raw = propertyArray(readCached("LineBreak.txt"), value => indexOf(RAW_LINE_BREAK, value), indexOf(RAW_LINE_BREAK, "XX"))
	const out = new Uint8Array(MAX_CODE_POINT + 1)
	const mn = indexOf(GC_NAMES, "Mn")
	const mc = indexOf(GC_NAMES, "Mc")

	for (let cp = 0; cp <= MAX_CODE_POINT; cp++) {
		const name = RAW_LINE_BREAK[raw[cp] ?? 0] ?? "XX"
		const category = gc[cp] ?? 0
		let resolved: number

		switch (name) {
			case "AI":
			case "SG":
			case "XX":
				resolved = LB("AL")
				break
			case "SA":
				resolved = category === mn || category === mc ? LB("CM") : LB("AL")
				break
			case "QU":
				resolved = GC_NAMES[category] === "Pi" ? LB("QU_PI") : GC_NAMES[category] === "Pf" ? LB("QU_PF") : LB("QU")
				break
			default:
				resolved = LB(name)
		}

		const unassignedPictographic = pictographic[cp] === 1 && GC_NAMES[category] === "Cn"

		out[cp] = resolved | ((ea[cp] ?? 0) << 6) | (unassignedPictographic ? 0x80 : 0)
	}

	return out
}

const GC_NAMES = [
	"Cn",
	"Lu",
	"Ll",
	"Lt",
	"Lm",
	"Lo",
	"Mn",
	"Mc",
	"Me",
	"Nd",
	"Nl",
	"No",
	"Pc",
	"Pd",
	"Ps",
	"Pe",
	"Pi",
	"Pf",
	"Po",
	"Sm",
	"Sc",
	"Sk",
	"So",
	"Zs",
	"Zl",
	"Zp",
	"Cc",
	"Cf",
	"Cs",
	"Co"
] as const

const RAW_LINE_BREAK = [...LINE_BREAK_CLASS_NAMES.filter(name => name !== "QU_PI" && name !== "QU_PF"), "AI", "SG", "XX", "SA"] as const

function bidiValues(): Uint8Array {
	return propertyArray(
		readCached("extracted/DerivedBidiClass.txt"),
		value => indexOf(BIDI_CLASS_NAMES, BIDI_LONG_NAMES[value] ?? value),
		indexOf(BIDI_CLASS_NAMES, "L")
	)
}

// [open, close] pairs, ascending by the opening bracket, and canonical singleton decompositions of bracket
// characters (U+2329 -> U+3008, ...), which N0 treats as equal.
function brackets(): { pairs: number[]; equivalents: number[] } {
	const pairs: number[] = []
	const bracketSet = new Set<number>()

	for (const raw of readCached("BidiBrackets.txt").split("\n")) {
		const line = raw.replace(/#.*/, "").trim()

		if (line === "") continue

		const [code = "", paired = "", kind = ""] = line.split(";").map(part => part.trim())
		const cp = Number.parseInt(code, 16)

		bracketSet.add(cp)

		if (kind === "o") pairs.push(cp, Number.parseInt(paired, 16))
	}

	const equivalents: number[] = []

	for (const line of readCached("UnicodeData.txt").split("\n")) {
		const fields = line.split(";")
		const cp = Number.parseInt(fields[0] ?? "", 16)
		const decomposition = fields[5] ?? ""

		if (!bracketSet.has(cp) || decomposition === "" || decomposition.startsWith("<") || decomposition.includes(" ")) continue

		equivalents.push(cp, Number.parseInt(decomposition, 16))
	}

	return { pairs, equivalents }
}

// Word_Break | General_Category Lu or Lt << 5 | Ll << 6 | Extended_Pictographic << 7.
function wordBreakValues(): Uint8Array {
	const out = propertyArray(
		readCached("auxiliary/WordBreakProperty.txt"),
		value => indexOf(WORD_BREAK_CLASS_NAMES, value),
		indexOf(WORD_BREAK_CLASS_NAMES, "Other")
	)
	const gc = propertyArray(readCached("extracted/DerivedGeneralCategory.txt"), value => indexOf(GC_NAMES, value), indexOf(GC_NAMES, "Cn"))

	for (let cp = 0; cp <= MAX_CODE_POINT; cp++) {
		const category = GC_NAMES[gc[cp] ?? 0]

		if (category === "Lu" || category === "Lt") out[cp] = (out[cp] ?? 0) | 0x20
		else if (category === "Ll") out[cp] = (out[cp] ?? 0) | 0x40
	}

	forEachEntry(readCached("emoji/emoji-data.txt"), (first, last, value) => {
		if (value !== "Extended_Pictographic") return

		for (let cp = first; cp <= last; cp++) out[cp] = (out[cp] ?? 0) | 0x80
	})

	return out
}

function runsOf(values: Uint8Array): [number, number][] {
	const runs: [number, number][] = []

	for (let cp = 0; cp <= MAX_CODE_POINT;) {
		const value = values[cp] ?? 0
		let end = cp + 1

		while (end <= MAX_CODE_POINT && values[end] === value) end++

		runs.push([end - cp, value])
		cp = end
	}

	return runs
}

// Hand-written cases at the head of BidiCharacterTest.txt are kept whole; the generated permutations and
// LineBreakTest.txt and WordBreakTest.txt are sampled every `step` cases.
const LINE_BREAK_SAMPLE_STEP = 16
const WORD_BREAK_SAMPLE_STEP = 2
const BIDI_SAMPLE_STEP = 96
const BIDI_GENERATED_MARKER = "# Permutations of sequences containing paired brackets"

function sample(file: UcdFile, step: number, keepUntil: string | null): string {
	const lines = readCached(file).split("\n")
	const out = [`# Sampled from ${file} ${UNICODE_VERSION} (SHA-256 ${UCD_FILES[file]}) by src/scripts/unicodeTables.ts.`]
	let keepAll = keepUntil !== null
	let index = 0

	for (const line of lines) {
		if (keepUntil !== null && line.startsWith(keepUntil)) keepAll = false
		if (line.startsWith("#") || line.trim() === "") continue

		const data = line.replace(/\s*#.*$/, "")

		if (keepAll || index++ % step === 0) out.push(data)
	}

	return `${out.join("\n")}\n`
}

async function main(): Promise<void> {
	if (!process.argv.includes("--offline")) {
		for (const file of Object.keys(UCD_FILES) as UcdFile[]) await download(file)
	}

	const lineBreak = runsOf(lineBreakValues())
	const bidi = runsOf(bidiValues())
	const { pairs, equivalents } = brackets()
	const wordBreak = runsOf(wordBreakValues())
	const source = `// Generated by src/scripts/unicodeTables.ts from the Unicode Character Database ${UNICODE_VERSION}. Do not edit.

export const UNICODE_VERSION = "${UNICODE_VERSION}"

// Line_Break (resolved, see packed.ts) | East_Asian_Width F/W/H << 6 | unassigned Extended_Pictographic << 7.
export const LINE_BREAK_RUNS = "${packRuns(lineBreak)}"

export const BIDI_CLASS_RUNS = "${packRuns(bidi)}"

// Opening and closing code points of each bracket pair, as deltas.
export const BIDI_BRACKET_PAIRS = "${packNumbers(pairs, true)}"

// Bracket code point and its canonical equivalent, as deltas.
export const BIDI_BRACKET_EQUIVALENTS = "${packNumbers(equivalents, true)}"

// Word_Break | General_Category Lu or Lt << 5 | Ll << 6 | Extended_Pictographic << 7.
export const WORD_BREAK_RUNS = "${packRuns(wordBreak)}"
`

	writeFileSync(OUTPUT, await format(source, { ...(await resolveConfig(OUTPUT)), filepath: OUTPUT }))
	mkdirSync(FIXTURES, { recursive: true })
	writeFileSync(`${FIXTURES}LineBreakTest.sample.txt`, sample("auxiliary/LineBreakTest.txt", LINE_BREAK_SAMPLE_STEP, null))
	writeFileSync(`${FIXTURES}BidiCharacterTest.sample.txt`, sample("BidiCharacterTest.txt", BIDI_SAMPLE_STEP, BIDI_GENERATED_MARKER))
	writeFileSync(`${FIXTURES}WordBreakTest.sample.txt`, sample("auxiliary/WordBreakTest.txt", WORD_BREAK_SAMPLE_STEP, null))
	console.log(
		`line break runs ${lineBreak.length}, bidi runs ${bidi.length}, bracket pairs ${pairs.length / 2}, word break runs ${wordBreak.length}`
	)
}

if (process.argv[1] === fileURLToPath(import.meta.url)) await main()
