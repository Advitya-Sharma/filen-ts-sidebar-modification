// Which shipped dictionary checks text in a language, from its tag as documents carry it (OOXML w:lang,
// BCP 47 with "-" or "_"). Regional variants without a dictionary of their own use the nearest one, the
// way Word and LibreOffice fall back; tags that mean "no linguistic content" get none.

export const DICTIONARY_IDS = [
	"en_US",
	"en_GB",
	"en_CA",
	"en_AU",
	"de_DE",
	"de_AT",
	"de_CH",
	"fr",
	"es_ES",
	"it_IT",
	"pt_BR",
	"pt_PT",
	"nl_NL",
	"pl_PL",
	"ru_RU",
	"sv_SE",
	"da_DK",
	"cs_CZ"
] as const

export type DictionaryId = (typeof DICTIONARY_IDS)[number]

// English regions that write American spelling; every other region writes British (en_CA and en_AU have
// their own dictionaries, New Zealand is closest to Australia).
const AMERICAN_ENGLISH = new Set(["us", "as", "gu", "mp", "pr", "um", "vi", "ph", "lr"])

const BY_LANGUAGE: Readonly<Record<string, DictionaryId>> = {
	fr: "fr",
	es: "es_ES",
	it: "it_IT",
	nl: "nl_NL",
	pl: "pl_PL",
	ru: "ru_RU",
	sv: "sv_SE",
	da: "da_DK",
	cs: "cs_CZ"
}

const NO_LANGUAGE = new Set(["", "x-none", "zxx", "und", "mis"])

// The dictionary for a language tag, or undefined when none fits.
export function dictionaryFor(tag: string): DictionaryId | undefined {
	const normalized = tag.trim().toLowerCase().replaceAll("_", "-")

	if (NO_LANGUAGE.has(normalized)) return undefined

	const [language = "", ...subtags] = normalized.split("-")
	// The region is the first two-letter or three-digit subtag after the language (a script may precede it).
	const region = subtags.find(subtag => /^(?:[a-z]{2}|\d{3})$/.test(subtag))

	switch (language) {
		case "en":
			if (region === undefined || AMERICAN_ENGLISH.has(region)) return "en_US"
			if (region === "ca") return "en_CA"
			if (region === "au" || region === "nz") return "en_AU"

			return "en_GB"
		case "de":
			if (region === "at") return "de_AT"
			if (region === "ch" || region === "li") return "de_CH"

			return "de_DE"
		case "pt":
			return region === "br" ? "pt_BR" : "pt_PT"
		default:
			return BY_LANGUAGE[language]
	}
}
