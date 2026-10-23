// Values RTF encodes as numbers: list number formats (\levelnfcN) and dates (DTTM).

// ST_NumberFormat (ECMA-376 Part 1 §17.18.59) for \levelnfcN, by the meaning the RTF specification gives each
// code; codes without a WordprocessingML format read as decimal.
const NUMBER_FORMATS: readonly string[] = [
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
	if (nfc === 255) return "none"

	return NUMBER_FORMATS[nfc] ?? "decimal"
}

// A DTTM (\revdttmN, \atndate): minutes in bits 0-5, hours 6-10, day 11-15, month 16-19, years since 1900
// 20-28. Undefined for 0 or fields out of range.
export function dttm(value: number): string | undefined {
	if (!Number.isFinite(value) || value === 0) return undefined

	const minute = value & 0x3f
	const hour = (value >>> 6) & 0x1f
	const day = (value >>> 11) & 0x1f
	const month = (value >>> 16) & 0xf
	const year = ((value >>> 20) & 0x1ff) + 1900

	if (month < 1 || month > 12 || day < 1 || day > 31 || hour > 23 || minute > 59) return undefined

	return isoDate([year, month, day, hour, minute, 0])
}

// [year, month, day, hour, minute, second] as an ISO 8601 UTC time.
export function isoDate(date: readonly number[]): string {
	const [year = 0, month = 0, day = 0, hour = 0, minute = 0, second = 0] = date
	const pad = (value: number, width = 2): string => String(Math.max(0, value)).padStart(width, "0")

	return `${pad(year, 4)}-${pad(Math.max(1, month))}-${pad(Math.max(1, day))}T${pad(hour)}:${pad(minute)}:${pad(second)}Z`
}
