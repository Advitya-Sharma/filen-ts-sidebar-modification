import { franc } from "franc-min"
import type { Language } from "./cases/types"

const FRANC_LATIN: Record<string, Language> = {
	eng: "en",
	deu: "de",
	spa: "es",
	fra: "fr",
	por: "pt",
	tur: "tr",
	ind: "id"
}

const MIN_LETTERS = 12

// Short Latin-script replies ("Merci !") are too short for trigrams and are left ungraded.
const MIN_LATIN_LETTERS = 24

const NOT_CJK = "[^\\s\\p{Script=Han}\\p{Script=Hiragana}\\p{Script=Katakana}]"

// Quotes, markup, file names, handles, numbers and paths are copied from tool results and say
// nothing about the reply's language. CJK runs carry no spaces, so a "word" stops at the script
// boundary.
function stripVerbatim(text: string): string {
	return text
		.replace(/`[^`]*`|<[^>]*>|\{[^}]*\}|\([^)]*\)/g, " ")
		.replace(/"[^"\n]*"|“[^”\n]*”|„[^“”\n]*[“”]|«[^»\n]*»|「[^」\n]*」|『[^』\n]*』/g, " ")
		.replace(new RegExp(`${NOT_CJK}*(?:[_/\\\\\\d]|\\.\\p{L})${NOT_CJK}*`, "gu"), " ")
}

function count(text: string, pattern: RegExp): number {
	return text.match(pattern)?.length ?? 0
}

// Function words: on short replies they beat trigrams, which confuse Spanish with Portuguese and
// English with almost anything.
const FUNCTION_WORDS: Record<string, readonly string[]> = {
	en: [
		"the",
		"and",
		"you",
		"your",
		"to",
		"of",
		"is",
		"are",
		"in",
		"it",
		"can",
		"please",
		"with",
		"for",
		"this",
		"that",
		"have",
		"not",
		"would",
		"i"
	],
	de: [
		"und",
		"der",
		"die",
		"das",
		"ich",
		"du",
		"nicht",
		"ist",
		"mit",
		"für",
		"auf",
		"ein",
		"eine",
		"zu",
		"den",
		"dem",
		"sie",
		"wir",
		"bitte",
		"kann"
	],
	es: [
		"el",
		"la",
		"los",
		"las",
		"y",
		"que",
		"en",
		"un",
		"una",
		"por",
		"para",
		"con",
		"no",
		"es",
		"tu",
		"puedo",
		"del",
		"lo",
		"hay",
		"está"
	],
	fr: [
		"le",
		"la",
		"les",
		"et",
		"des",
		"un",
		"une",
		"est",
		"pour",
		"que",
		"dans",
		"je",
		"vous",
		"tu",
		"ne",
		"pas",
		"avec",
		"sur",
		"du",
		"ou"
	],
	pt: ["o", "os", "e", "que", "em", "um", "uma", "para", "com", "não", "é", "você", "do", "da", "no", "na", "dos", "das", "está", "ou"],
	tr: ["ve", "bir", "bu", "için", "ile", "ne", "mi", "değil", "var", "yok", "çok", "gibi", "olarak", "şu", "ben", "sen", "veya", "ama"],
	id: [
		"dan",
		"yang",
		"di",
		"ke",
		"dari",
		"ini",
		"itu",
		"untuk",
		"dengan",
		"saya",
		"anda",
		"tidak",
		"ada",
		"akan",
		"bisa",
		"sudah",
		"kamu",
		"atau"
	]
}

function functionWordVote(prose: string): Language | null {
	const words = prose.toLowerCase().split(/[^\p{L}]+/u)
	const scores = Object.entries(FUNCTION_WORDS)
		.map(([lang, list]) => [lang, words.filter(word => list.includes(word)).length] as const)
		.sort((a, b) => b[1] - a[1])
	const [first, second] = scores

	return first !== undefined && first[1] >= 3 && first[1] >= 2 * (second?.[1] ?? 0) ? (first[0] as Language) : null
}

// Script decides the non-Latin languages outright; franc-min (trigram model, restricted to the
// evaluated Latin-script languages) decides the rest, after spelling checks for the pairs it
// confuses on short text.
export function detectLanguage(text: string): Language | "und" {
	const cleaned = stripVerbatim(text)
	const kana = count(cleaned, /[\p{Script=Hiragana}\p{Script=Katakana}]/gu)
	const han = count(cleaned, /\p{Script=Han}/gu)
	const scripts: [Language, number][] = [
		["ar", count(cleaned, /\p{Script=Arabic}/gu)],
		["hi", count(cleaned, /\p{Script=Devanagari}/gu)],
		["ru", count(cleaned, /\p{Script=Cyrillic}/gu)],
		[kana > 0 ? "ja" : "zh", kana + han]
	]
	const latin = count(cleaned, /\p{Script=Latin}/gu)
	const [topScript, topCount] = scripts.reduce((best, entry) => (entry[1] > best[1] ? entry : best))

	// CJK packs a word into one or two characters.
	if ((topScript === "ja" || topScript === "zh" ? topCount * 3 : topCount) > latin) {
		return topCount >= (topScript === "ja" || topScript === "zh" ? 4 : MIN_LETTERS) ? topScript : "und"
	}

	// Names (chat partners, file titles) carry another language's letters, so a capitalized word is
	// evidence only at the start of a sentence; German keeps enough lowercase function words.
	const prose = cleaned
		.split(/\s+/u)
		.filter((word, index, all) => word.length > 0 && (!/^\p{Lu}/u.test(word) || index === 0 || /[.!?:]$/u.test(all[index - 1] ?? "")))
		.join(" ")

	if (count(prose, /\p{Script=Latin}/gu) < MIN_LATIN_LETTERS) {
		return "und"
	}

	if (prose.split(" ").filter(word => /[ığş]/u.test(word)).length >= 2) {
		return "tr"
	}

	const voted = functionWordVote(prose)

	if (voted !== null) {
		return voted
	}

	return (
		FRANC_LATIN[
			franc(prose, {
				only: Object.keys(FRANC_LATIN),
				minLength: MIN_LETTERS
			})
		] ?? "und"
	)
}
