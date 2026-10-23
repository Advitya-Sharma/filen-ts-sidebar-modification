import type { DocLevel, DocLists } from "../lists"
import { applyChar, applyPara, defaultChar, defaultPara } from "../props"
import { type PropsContext, paragraphProperties, runProperties } from "./props"
import { STORY_NAMESPACES, attr, val } from "./xml"

// PlfLst / PlfLfo → numbering.xml (ECMA-376 Part 1 §17.9): one abstractNum per LSTF, one num per LFO (numId =
// iLfo, as paragraphs refer to it), level overrides as lvlOverride.

// MSONFC ([MS-OSHARED] §2.2.1.3) → ST_NumberFormat; unknown values fall back to decimal.
const FORMATS: readonly string[] = [
	"decimal",
	"upperRoman",
	"lowerRoman",
	"upperLetter",
	"lowerLetter",
	"ordinal",
	"cardinalText",
	"ordinalText",
	"hex",
	"chicago",
	"ideographDigital",
	"japaneseCounting",
	"aiueo",
	"iroha",
	"decimalFullWidth",
	"decimalHalfWidth",
	"japaneseLegal",
	"japaneseDigitalTenThousand",
	"decimalEnclosedCircle",
	"decimalFullWidth2",
	"aiueoFullWidth",
	"irohaFullWidth",
	"decimalZero",
	"bullet",
	"ganada",
	"chosung",
	"decimalEnclosedFullstop",
	"decimalEnclosedParen",
	"decimalEnclosedCircleChinese",
	"ideographEnclosedCircle",
	"ideographTraditional",
	"ideographZodiac",
	"ideographZodiacTraditional",
	"taiwaneseCounting",
	"ideographLegalTraditional",
	"taiwaneseCountingThousand",
	"taiwaneseDigital",
	"chineseCounting",
	"chineseLegalSimplified",
	"chineseCountingThousand",
	"decimal",
	"koreanDigital",
	"koreanCounting",
	"koreanLegal",
	"koreanDigital2",
	"hebrew1",
	"arabicAlpha",
	"hebrew2",
	"arabicAbjad",
	"hindiVowels",
	"hindiConsonants",
	"hindiNumbers",
	"hindiCounting",
	"thaiLetters",
	"thaiNumbers",
	"thaiCounting",
	"vietnameseCounting",
	"numberInDash",
	"russianLower",
	"russianUpper"
]

export function numberFormat(nfc: number): string {
	if (nfc === 0xff) return "none"

	return FORMATS[nfc] ?? "decimal"
}

const JC = ["left", "center", "right"]
const SUFFIX = ["tab", "space", "nothing"]

function hex8(value: number): string {
	return (value >>> 0).toString(16).toUpperCase().padStart(8, "0")
}

// The level text with its placeholders (code units 0–8 at the rgbxchNums offsets) as %1 … %9.
function levelText(level: DocLevel): string {
	const placeholders = new Set(level.placeholders)
	let out = ""

	for (let k = 0; k < level.text.length; k++) {
		const code = level.text.charCodeAt(k)

		out += placeholders.has(k) && code <= 8 ? `%${String(code + 1)}` : level.text.charAt(k)
	}

	return out
}

function level(entry: DocLevel, ilvl: number, context: PropsContext): string {
	const para = defaultPara()
	const char = defaultChar()

	applyPara(para, entry.papx, 0, entry.papx.length)
	applyChar(char, entry.chpx, 0, entry.chpx.length, { style: defaultChar() })

	const ppr = paragraphProperties(para, defaultPara(), context)
	const rpr = runProperties(char, defaultChar(), context)

	return (
		`<w:lvl${attr("w:ilvl", ilvl)}${entry.tentative ? ' w:tentative="1"' : ""}>` +
		val("start", Math.max(0, entry.start)) +
		val("numFmt", numberFormat(entry.nfc)) +
		(entry.noRestart ? val("lvlRestart", entry.restartLimit) : "") +
		(entry.legal ? "<w:isLgl/>" : "") +
		val("suff", SUFFIX[entry.follow] ?? "tab") +
		`<w:lvlText${attr("w:val", levelText(entry))}/>` +
		val("lvlJc", JC[entry.jc] ?? "left") +
		(ppr === "" ? "" : `<w:pPr>${ppr}</w:pPr>`) +
		(rpr === "" ? "" : `<w:rPr>${rpr}</w:rPr>`) +
		"</w:lvl>"
	)
}

export interface NumberingPart {
	readonly xml: string
	// numIds written (iLfo values with a list behind them).
	readonly numIds: ReadonlySet<number>
}

export function numberingPart(lists: DocLists, context: PropsContext): NumberingPart | undefined {
	if (lists.lists.length === 0 || lists.instances.length === 0) return undefined

	const out: string[] = [`<w:numbering ${STORY_NAMESPACES}>`]
	const byLsid = new Map<number, number>()

	lists.lists.forEach((list, k) => {
		if (!byLsid.has(list.lsid)) byLsid.set(list.lsid, k)

		out.push(
			`<w:abstractNum${attr("w:abstractNumId", k)}>` +
				val("nsid", hex8(list.lsid)) +
				val("multiLevelType", list.simple ? "singleLevel" : list.hybrid ? "hybridMultilevel" : "multilevel") +
				val("tmpl", hex8(list.tplc))
		)

		list.levels.forEach((entry, ilvl) => out.push(level(entry, ilvl, context)))
		out.push("</w:abstractNum>")
	})

	const numIds = new Set<number>()

	lists.instances.forEach((instance, k) => {
		const abstract = byLsid.get(instance.lsid)

		if (abstract === undefined) return

		numIds.add(k + 1)
		out.push(`<w:num${attr("w:numId", k + 1)}>${val("abstractNumId", abstract)}`)

		for (const override of instance.overrides) {
			if (override.level > 8 || (override.startAt === undefined && override.formatting === undefined)) continue

			out.push(
				`<w:lvlOverride${attr("w:ilvl", override.level)}>` +
					(override.startAt === undefined ? "" : val("startOverride", Math.max(0, override.startAt))) +
					(override.formatting === undefined ? "" : level(override.formatting, override.level, context)) +
					"</w:lvlOverride>"
			)
		}

		out.push("</w:num>")
	})

	out.push("</w:numbering>")

	return numIds.size === 0 ? undefined : { xml: out.join(""), numIds }
}
