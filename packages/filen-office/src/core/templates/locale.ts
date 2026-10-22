// What a blank document takes from the user's locale: its language tag (BCP 47, as w:lang carries it) and
// its paper. US Letter is the paper of the United States, Canada, Mexico and the Philippines; every other
// region uses ISO 216 A4.

const LETTER_REGIONS: ReadonlySet<string> = new Set(["US", "CA", "MX", "PH"])
const FALLBACK_TAG = "en-US"

export type TemplatePaper = "letter" | "a4"

export interface TemplateLocale {
	// language-region where the region is known, else the language alone.
	readonly tag: string
	readonly paper: TemplatePaper
}

function canonical(tag: string): Intl.Locale | undefined {
	try {
		return new Intl.Locale(tag)
	} catch {
		return undefined
	}
}

// The likely region of a tag without one ("de" -> DE), where the platform knows it.
function likelyRegion(locale: Intl.Locale): string | undefined {
	try {
		return locale.maximize().region
	} catch {
		return undefined
	}
}

export function templateLocale(tag: string): TemplateLocale {
	const locale = canonical(tag) ?? canonical(FALLBACK_TAG)
	const language = locale?.language ?? "en"
	const region = locale === undefined ? "US" : (locale.region ?? likelyRegion(locale))
	const shown = region === undefined ? language : `${language}-${region}`

	return {
		tag: /^[A-Za-z0-9-]+$/.test(shown) ? shown : FALLBACK_TAG,
		paper: region !== undefined && LETTER_REGIONS.has(region) ? "letter" : "a4"
	}
}
