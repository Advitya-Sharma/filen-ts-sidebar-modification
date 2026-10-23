import { CodePointTable, unpackNumbers } from "../../text/packed"
import { COMBINING_CLASS_RUNS, INDIC_RUNS, MIRROR_PAIRS, SCRIPT_CODES, SCRIPT_RUNS, SHAPING_PROPS_RUNS } from "./shaping.generated"

// Lazily expanded Unicode data for shaping. Nothing is decoded until a PDF is first written.

export type ScriptCode = (typeof SCRIPT_CODES)[number]

export const JoiningType = { U: 0, R: 1, D: 2, C: 3, L: 4, T: 5 } as const

export const IndicCategory = {
	X: 0,
	C: 1,
	V: 2,
	N: 3,
	H: 4,
	ZWNJ: 5,
	ZWJ: 6,
	M: 7,
	SM: 8,
	A: 9,
	PLACEHOLDER: 10,
	DOTTEDCIRCLE: 11,
	RS: 12,
	Repha: 13,
	Ra: 14,
	CM: 15,
	Symbol: 16,
	CS: 17
} as const

export const MatraPosition = { NONE: 0, LEFT: 1, TOP: 2, BOTTOM: 3, RIGHT: 4, SPLIT: 5, VISUAL_LEFT: 6, OVERSTRUCK: 7 } as const

let scripts: CodePointTable | null = null
let props: CodePointTable | null = null
let indic: CodePointTable | null = null
let combining: CodePointTable | null = null
let mirrors: Map<number, number> | null = null

export const SCRIPT_COMMON = SCRIPT_CODES.indexOf("Zyyy")
export const SCRIPT_INHERITED = SCRIPT_CODES.indexOf("Zinh")

export function scriptOf(cp: number): number {
	return (scripts ??= new CodePointTable(SCRIPT_RUNS)).get(cp)
}

export function scriptCode(index: number): ScriptCode {
	return SCRIPT_CODES[index] ?? "Zzzz"
}

export function joiningType(cp: number): number {
	return (props ??= new CodePointTable(SHAPING_PROPS_RUNS)).get(cp) & 7
}

export function isMark(cp: number): boolean {
	return ((props ??= new CodePointTable(SHAPING_PROPS_RUNS)).get(cp) & 8) !== 0
}

export function isDefaultIgnorable(cp: number): boolean {
	return ((props ??= new CodePointTable(SHAPING_PROPS_RUNS)).get(cp) & 16) !== 0
}

export function indicCategory(cp: number): number {
	return (indic ??= new CodePointTable(INDIC_RUNS)).get(cp) & 31
}

export function matraPosition(cp: number): number {
	return (indic ??= new CodePointTable(INDIC_RUNS)).get(cp) >>> 5
}

// Canonical_Combining_Class with HarfBuzz's modified classes (hb-unicode.hh, HarfBuzz 14.6.0, MIT): Hebrew
// points and Arabic harakat in the order fonts are built for (shadda before the vowel marks), Telugu length
// marks and Thai, Lao and Tibetan vowel signs.
export function combiningClass(cp: number): number {
	const ccc = (combining ??= new CodePointTable(COMBINING_CLASS_RUNS)).get(cp)

	return MODIFIED_CLASSES.get(ccc) ?? ccc
}

const MODIFIED_CLASSES = new Map([
	[10, 22],
	[11, 15],
	[12, 16],
	[13, 17],
	[14, 23],
	[15, 18],
	[16, 19],
	[17, 20],
	[18, 21],
	[19, 14],
	[20, 24],
	[21, 12],
	[22, 25],
	[23, 13],
	[24, 10],
	[25, 11],
	[27, 28],
	[28, 29],
	[29, 30],
	[30, 31],
	[31, 32],
	[32, 33],
	[33, 27],
	[84, 4],
	[91, 5],
	[103, 3],
	[130, 132],
	[132, 131]
])

export function mirrorOf(cp: number): number {
	if (mirrors === null) {
		const list = unpackNumbers(MIRROR_PAIRS, true)

		mirrors = new Map()

		for (let i = 0; i + 1 < list.length; i += 2) mirrors.set(list[i] ?? 0, list[i + 1] ?? 0)
	}

	return mirrors.get(cp) ?? 0
}

// OpenType script tags for an ISO 15924 code, preferred first (new Indic tags before the old ones).
export function openTypeScripts(code: ScriptCode): readonly string[] {
	return OT_SCRIPTS[code] ?? [code.toLowerCase()]
}

const OT_SCRIPTS: Partial<Record<ScriptCode, readonly string[]>> = {
	Zzzz: ["DFLT"],
	Zyyy: ["DFLT"],
	Zinh: ["DFLT"],
	Deva: ["dev2", "deva"],
	Beng: ["bng2", "beng"],
	Guru: ["gur2", "guru"],
	Gujr: ["gjr2", "gujr"],
	Orya: ["ory2", "orya"],
	Taml: ["tml2", "taml"],
	Telu: ["tel2", "telu"],
	Knda: ["knd2", "knda"],
	Mlym: ["mlm2", "mlym"],
	Mymr: ["mym2", "mymr"],
	Laoo: ["lao "],
	Nkoo: ["nko "],
	Hira: ["kana"],
	Kana: ["kana"],
	Hang: ["hang"],
	Hani: ["hani"],
	Bopo: ["bopo"]
}

export function codePointAt(text: string, index: number, end: number): number {
	const high = text.charCodeAt(index)

	if (high < 0xd800 || high > 0xdbff || index + 1 >= end) return high

	const low = text.charCodeAt(index + 1)

	return low >= 0xdc00 && low <= 0xdfff ? ((high - 0xd800) << 10) + (low - 0xdc00) + 0x10000 : high
}

// OpenType language system tag for a BCP 47 language tag (OpenType "Language system tags" registry), for
// the languages documents commonly mark; undefined uses the script's default language system.
export function openTypeLanguage(language: string | undefined): string | undefined {
	if (language === undefined || language === "") return undefined

	const lower = language.toLowerCase()
	const [primary = ""] = lower.split(/[-_]/)

	if (primary === "zh") {
		if (/-(hk|mo)\b/.test(lower)) return "ZHH "
		if (/-(hant|tw)\b/.test(lower)) return "ZHT "

		return "ZHS "
	}

	return OT_LANGUAGES[primary]
}

const OT_LANGUAGES: Record<string, string> = {
	ja: "JAN ",
	ko: "KOR ",
	ar: "ARA ",
	fa: "FAR ",
	ur: "URD ",
	ps: "PAS ",
	sd: "SND ",
	ug: "UYG ",
	ku: "KUR ",
	he: "IWR ",
	iw: "IWR ",
	yi: "JII ",
	hi: "HIN ",
	mr: "MAR ",
	ne: "NEP ",
	sa: "SAN ",
	bn: "BEN ",
	as: "ASM ",
	pa: "PAN ",
	gu: "GUJ ",
	or: "ORI ",
	ta: "TAM ",
	te: "TEL ",
	kn: "KAN ",
	ml: "MAL ",
	si: "SNH ",
	th: "THA ",
	lo: "LAO ",
	hy: "HYE ",
	ka: "KAT ",
	am: "AMH ",
	ru: "RUS ",
	uk: "UKR ",
	be: "BEL ",
	bg: "BGR ",
	sr: "SRB ",
	mk: "MKD ",
	kk: "KAZ ",
	ky: "KIR ",
	tt: "TAT ",
	ba: "BSH ",
	cv: "CHU ",
	el: "ELL ",
	tr: "TRK ",
	az: "AZE ",
	ro: "ROM ",
	mo: "MOL ",
	pl: "PLK ",
	nl: "NLD ",
	de: "DEU ",
	ca: "CAT ",
	vi: "VIT ",
	ga: "IRI ",
	is: "ISL ",
	ms: "MLY ",
	sk: "SKY ",
	cs: "CSY ",
	hu: "HUN ",
	sv: "SVE ",
	fi: "FIN ",
	da: "DAN ",
	nb: "NOR ",
	no: "NOR ",
	fr: "FRA ",
	es: "ESP ",
	pt: "PTG ",
	it: "ITA ",
	hr: "HRV ",
	sl: "SLV ",
	lt: "LTH ",
	lv: "LVI ",
	et: "ETI ",
	sq: "SQI ",
	eu: "EUQ ",
	gl: "GAL ",
	mn: "MNG ",
	kmr: "KUR "
}
