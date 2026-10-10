import type { XmlDocument } from "../../xml/index"
import { readSettings, type DocxSettings } from "../model/settings"
import type { LanguageProps } from "../props/types"
import { DEFAULT_COLOR_SCHEME_MAPPING, type ColorSchemeMapping } from "../theme/theme"

// The parts of word/settings.xml (§17.15) that property resolution depends on, derived from the model's
// settings reader.
export interface StyleSettings {
	readonly colorMapping: ColorSchemeMapping
	// Languages theme fonts are picked for (w:themeFontLang, §17.15.1.88).
	readonly themeFontLanguage: LanguageProps
	// Twips.
	readonly defaultTabStop: number
	// 11 = Word 2003, 12 = 2007, 14 = 2010, 15 = 2013 and later.
	readonly compatibilityMode: number | undefined
	// MS-DOCX §2.3.1: when off, the default paragraph style's 11/12pt size and left justification do not
	// override a table style.
	readonly overrideTableStyleFontSizeAndJustification: boolean
}

export const DEFAULT_STYLE_SETTINGS: StyleSettings = {
	colorMapping: DEFAULT_COLOR_SCHEME_MAPPING,
	themeFontLanguage: {},
	defaultTabStop: 720,
	compatibilityMode: undefined,
	overrideTableStyleFontSizeAndJustification: false
}

export function styleSettings(settings: DocxSettings): StyleSettings {
	if (!settings.present) return DEFAULT_STYLE_SETTINGS

	const languages = settings.themeFontLanguages
	const themeFontLanguage: { latin?: string; eastAsia?: string; bidi?: string } = {}

	if (languages?.latin !== undefined) themeFontLanguage.latin = languages.latin
	if (languages?.eastAsia !== undefined) themeFontLanguage.eastAsia = languages.eastAsia
	if (languages?.bidi !== undefined) themeFontLanguage.bidi = languages.bidi

	return {
		colorMapping:
			settings.colorSchemeMapping === undefined
				? DEFAULT_COLOR_SCHEME_MAPPING
				: { ...DEFAULT_COLOR_SCHEME_MAPPING, ...settings.colorSchemeMapping },
		themeFontLanguage,
		defaultTabStop: settings.defaultTabStop ?? 720,
		compatibilityMode: settings.compatibilityMode,
		overrideTableStyleFontSizeAndJustification: settings.overrideTableStyleFontSizeAndJustification
	}
}

export function readStyleSettings(document: XmlDocument): StyleSettings {
	return styleSettings(readSettings(document))
}
