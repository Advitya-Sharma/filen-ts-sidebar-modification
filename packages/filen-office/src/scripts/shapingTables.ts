import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { format, resolveConfig } from "prettier"
import { packNumbers, packRuns } from "../core/text/packed"
import { CACHE, UNICODE_VERSION, forEachEntry, propertyArray, readCached, runsOf } from "./unicodeTables"

// Generates src/core/pdfwrite/fonts/shaping.generated.ts, the Unicode data the PDF text shaper needs: Script,
// Joining_Type, marks and Default_Ignorable_Code_Point, Indic syllabic and positional categories,
// Canonical_Combining_Class and Bidi_Mirroring_Glyph. Run `pnpm run generate:unicode` first (it caches
// DerivedGeneralCategory.txt).
//
//   pnpm exec tsx src/scripts/shapingTables.ts [--offline]

const BASE = `https://www.unicode.org/Public/${UNICODE_VERSION}/ucd/`

const FILES = {
	"Scripts.txt": "0071fd81b6aeae25f6e8bce8efec3066a6476a91b49bdb2f52dc76e817862a6a",
	"ArabicShaping.txt": "8ccde4ebd070500a68e8bcb9d5514522fb0bebe57b9ebad3ff925b8e1e669f83",
	"IndicSyllabicCategory.txt": "a2b3aacf6b3e7bad4ca351ef985d9543825e20280ff280c25f646d9bc4ce304c",
	"IndicPositionalCategory.txt": "c9ad44d267c317ee0e1ec66cbd4d15033c2964f905d11fc8905915f9a5a9ae52",
	"BidiMirroring.txt": "cd54810ebf52f0e61a730c8b9cb25975de6c85f6d788a559b416afd548923fd6",
	"DerivedCoreProperties.txt": "09c928886a178fcafd93c29e4bd59073a058e5a100b716d425cb563ab50f68c9",
	"extracted/DerivedCombiningClass.txt": "ef6b2611cfb660dba3f6b458b9eb4b05f44ed2417302ee7749d7f0f348793121"
} as const

type ShapingFile = keyof typeof FILES

const OUTPUT = fileURLToPath(new URL("../core/pdfwrite/fonts/shaping.generated.ts", import.meta.url))
const MAX_CODE_POINT = 0x10ffff

// Scripts the shaper tells apart, by ISO 15924 code and Unicode name. Index 0 is every other script.
const SCRIPTS: readonly (readonly [string, string])[] = [
	["Zzzz", "Unknown"],
	["Zyyy", "Common"],
	["Zinh", "Inherited"],
	["Latn", "Latin"],
	["Grek", "Greek"],
	["Cyrl", "Cyrillic"],
	["Armn", "Armenian"],
	["Hebr", "Hebrew"],
	["Arab", "Arabic"],
	["Syrc", "Syriac"],
	["Thaa", "Thaana"],
	["Nkoo", "Nko"],
	["Deva", "Devanagari"],
	["Beng", "Bengali"],
	["Guru", "Gurmukhi"],
	["Gujr", "Gujarati"],
	["Orya", "Oriya"],
	["Taml", "Tamil"],
	["Telu", "Telugu"],
	["Knda", "Kannada"],
	["Mlym", "Malayalam"],
	["Sinh", "Sinhala"],
	["Thai", "Thai"],
	["Laoo", "Lao"],
	["Tibt", "Tibetan"],
	["Mymr", "Myanmar"],
	["Geor", "Georgian"],
	["Hang", "Hangul"],
	["Ethi", "Ethiopic"],
	["Cher", "Cherokee"],
	["Khmr", "Khmer"],
	["Mong", "Mongolian"],
	["Hira", "Hiragana"],
	["Kana", "Katakana"],
	["Bopo", "Bopomofo"],
	["Hani", "Han"]
]

const JOINING_TYPES = ["U", "R", "D", "C", "L", "T"] as const

// Indic syllabic categories as the shaper groups them (OpenType "Creating and supporting OpenType fonts for
// Indic scripts", character classes).
const INDIC_CATEGORIES = [
	"X",
	"C",
	"V",
	"N",
	"H",
	"ZWNJ",
	"ZWJ",
	"M",
	"SM",
	"A",
	"PLACEHOLDER",
	"DOTTEDCIRCLE",
	"RS",
	"Repha",
	"Ra",
	"CM",
	"Symbol",
	"CS"
] as const

const SYLLABIC: Record<string, (typeof INDIC_CATEGORIES)[number]> = {
	Avagraha: "Symbol",
	Bindu: "SM",
	Brahmi_Joining_Number: "PLACEHOLDER",
	Cantillation_Mark: "A",
	Consonant: "C",
	Consonant_Dead: "C",
	Consonant_Final: "CM",
	Consonant_Head_Letter: "C",
	Consonant_Initial_Postfixed: "C",
	Consonant_Killer: "M",
	Consonant_Medial: "CM",
	Consonant_Placeholder: "PLACEHOLDER",
	Consonant_Preceding_Repha: "Repha",
	Consonant_Prefixed: "X",
	Consonant_Subjoined: "CM",
	Consonant_Succeeding_Repha: "CM",
	Consonant_With_Stacker: "CS",
	Gemination_Mark: "SM",
	Invisible_Stacker: "H",
	Joiner: "ZWJ",
	Modifying_Letter: "X",
	Non_Joiner: "ZWNJ",
	Nukta: "N",
	Number: "PLACEHOLDER",
	Number_Joiner: "PLACEHOLDER",
	Pure_Killer: "M",
	Register_Shifter: "RS",
	Reordering_Killer: "M",
	Syllable_Modifier: "SM",
	Tone_Letter: "X",
	Tone_Mark: "N",
	Virama: "H",
	Visarga: "SM",
	Vowel: "V",
	Vowel_Dependent: "M",
	Vowel_Independent: "V",
	Other: "X"
}

// Indic_Positional_Category folded to where the character (or the last part of a split one) sits: 0 none,
// 1 left, 2 top, 3 bottom, 4 right, 5 left and right (split, ends right), 6 visual-order left, 7 overstruck.
const POSITIONS: Record<string, number> = {
	Left: 1,
	Top: 2,
	Bottom: 3,
	Right: 4,
	Left_And_Right: 5,
	Visual_Order_Left: 6,
	Overstruck: 7,
	Top_And_Right: 4,
	Top_And_Left: 2,
	Top_And_Left_And_Right: 4,
	Top_And_Bottom: 3,
	Top_And_Bottom_And_Right: 4,
	Top_And_Bottom_And_Left: 3,
	Bottom_And_Right: 4,
	Bottom_And_Left: 3
}

// The letter RA of each script whose RA + virama forms a reph.
const REPH_RA = [0x0930, 0x09b0, 0x09f0, 0x0a30, 0x0ab0, 0x0b30, 0x0bb0, 0x0c30, 0x0cb0, 0x0d30, 0x0dbb]

const INDIC_RANGE: readonly [number, number] = [0x0900, 0x0dff]

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

function cachePath(file: ShapingFile): string {
	return `${CACHE}${file.replaceAll("/", "_")}`
}

function read(file: ShapingFile): string {
	const bytes = new Uint8Array(readFileSync(cachePath(file)))

	if (sha256(bytes) !== FILES[file]) throw new Error(`${file} does not match its pinned SHA-256`)

	return new TextDecoder().decode(bytes)
}

async function download(file: ShapingFile): Promise<void> {
	const path = cachePath(file)

	if (existsSync(path) && sha256(new Uint8Array(readFileSync(path))) === FILES[file]) return

	const response = await fetch(`${BASE}${file}`)

	if (!response.ok) throw new Error(`${file}: HTTP ${response.status}`)

	const bytes = new Uint8Array(await response.arrayBuffer())

	if (sha256(bytes) !== FILES[file]) throw new Error(`${file} does not match its pinned SHA-256`)

	mkdirSync(CACHE, { recursive: true })
	writeFileSync(path, bytes)
}

function scriptValues(): Uint8Array {
	const byName = new Map(SCRIPTS.map(([, name], index) => [name, index]))

	return propertyArray(read("Scripts.txt"), value => byName.get(value) ?? 0, 0)
}

// Joining_Type (bits 0-2), General_Category M* (bit 3), Default_Ignorable_Code_Point (bit 4).
function propsValues(): Uint8Array {
	const marks = new Set(["Mn", "Mc", "Me"])
	const transparent = new Set(["Mn", "Me", "Cf"])
	const out = new Uint8Array(MAX_CODE_POINT + 1)

	forEachEntry(readCached("extracted/DerivedGeneralCategory.txt"), (first, last, value) => {
		// Characters ArabicShaping.txt does not list are T when Mn, Me or Cf, else U.
		const bits = (marks.has(value) ? 8 : 0) | (transparent.has(value) ? JOINING_TYPES.indexOf("T") : 0)

		out.fill(bits, first, last + 1)
	})

	for (const raw of read("ArabicShaping.txt").split("\n")) {
		const line = raw.replace(/#.*/, "").trim()

		if (line === "") continue

		const [code = "", , type = ""] = line.split(";").map(part => part.trim())
		const cp = Number.parseInt(code, 16)
		const index = JOINING_TYPES.indexOf(type as (typeof JOINING_TYPES)[number])

		if (index < 0) throw new Error(`Unknown joining type ${type}`)

		out[cp] = ((out[cp] ?? 0) & ~7) | index
	}

	forEachEntry(read("DerivedCoreProperties.txt"), (first, last, value) => {
		if (value !== "Default_Ignorable_Code_Point") return

		for (let cp = first; cp <= last; cp++) out[cp] = (out[cp] ?? 0) | 16
	})

	return out
}

// Indic category (bits 0-4) and matra position (bits 5-7), for the Indic block range only.
function indicValues(): Uint8Array {
	const out = new Uint8Array(MAX_CODE_POINT + 1)
	const category = new Uint8Array(MAX_CODE_POINT + 1)
	const position = new Uint8Array(MAX_CODE_POINT + 1)

	forEachEntry(read("IndicSyllabicCategory.txt"), (first, last, value) => {
		const mapped = SYLLABIC[value]

		if (mapped === undefined) throw new Error(`Unknown Indic syllabic category ${value}`)

		category.fill(INDIC_CATEGORIES.indexOf(mapped), first, last + 1)
	})
	forEachEntry(read("IndicPositionalCategory.txt"), (first, last, value) => {
		if (value === "NA" || value === "Not_Applicable") return

		const mapped = POSITIONS[value]

		if (mapped === undefined) throw new Error(`Unknown Indic positional category ${value}`)

		position.fill(mapped, first, last + 1)
	})

	for (const cp of REPH_RA) category[cp] = INDIC_CATEGORIES.indexOf("Ra")

	const [from, to] = INDIC_RANGE

	for (let cp = from; cp <= to; cp++) out[cp] = (category[cp] ?? 0) | ((position[cp] ?? 0) << 5)

	// Characters that stand in for a base consonant in any Indic run.
	for (const cp of [0x00a0, 0x00d7, 0x2010, 0x2011, 0x2012, 0x2013, 0x2014, 0x2015, 0x2022, 0x25fb, 0x25fc, 0x25fd, 0x25fe]) {
		out[cp] = INDIC_CATEGORIES.indexOf("PLACEHOLDER")
	}

	out[0x25cc] = INDIC_CATEGORIES.indexOf("DOTTEDCIRCLE")
	out[0x200c] = INDIC_CATEGORIES.indexOf("ZWNJ")
	out[0x200d] = INDIC_CATEGORIES.indexOf("ZWJ")

	return out
}

function combiningValues(): Uint8Array {
	return propertyArray(read("extracted/DerivedCombiningClass.txt"), value => Number.parseInt(value, 10), 0)
}

function mirrorPairs(): number[] {
	const pairs: number[] = []

	for (const raw of read("BidiMirroring.txt").split("\n")) {
		const line = raw.replace(/#.*/, "").trim()

		if (line === "") continue

		const [from = "", to = ""] = line.split(";").map(part => part.trim())

		pairs.push(Number.parseInt(from, 16), Number.parseInt(to, 16))
	}

	return pairs
}

async function main(): Promise<void> {
	if (!process.argv.includes("--offline")) {
		for (const file of Object.keys(FILES) as ShapingFile[]) await download(file)
	}

	const scripts = runsOf(scriptValues())
	const props = runsOf(propsValues())
	const indic = runsOf(indicValues())
	const combining = runsOf(combiningValues())
	const mirrors = mirrorPairs()
	const source = `// Generated by src/scripts/shapingTables.ts from the Unicode Character Database ${UNICODE_VERSION}. Do not edit.

// ISO 15924 codes; SCRIPT_RUNS values index this list (0: any other script).
export const SCRIPT_CODES = ${JSON.stringify(SCRIPTS.map(([code]) => code))} as const

export const SCRIPT_RUNS = "${packRuns(scripts)}"

// Joining_Type (bits 0-2: ${JOINING_TYPES.join(", ")}) | General_Category M* << 3 | Default_Ignorable_Code_Point << 4.
export const SHAPING_PROPS_RUNS = "${packRuns(props)}"

// Indic category (bits 0-4: ${INDIC_CATEGORIES.join(", ")}) | position << 5 (1 left, 2 top, 3 bottom,
// 4 right, 5 left and right, 6 visual-order left, 7 overstruck).
export const INDIC_RUNS = "${packRuns(indic)}"

export const COMBINING_CLASS_RUNS = "${packRuns(combining)}"

// Bidi_Mirroring_Glyph pairs, as deltas.
export const MIRROR_PAIRS = "${packNumbers(mirrors, true)}"
`

	writeFileSync(OUTPUT, await format(source, { ...(await resolveConfig(OUTPUT)), filepath: OUTPUT }))
	console.log(
		`script runs ${scripts.length}, props runs ${props.length}, indic runs ${indic.length}, ccc runs ${combining.length}, mirrors ${mirrors.length / 2}`
	)
}

if (process.argv[1] === fileURLToPath(import.meta.url)) await main()
