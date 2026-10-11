import { describe, expect, it } from "vitest"
import { dictionaryFor } from "../languages"

describe("dictionaryFor", () => {
	it.each([
		["en-US", "en_US"],
		["en", "en_US"],
		["EN_us", "en_US"],
		["en-PH", "en_US"],
		["en-GB", "en_GB"],
		["en-IE", "en_GB"],
		["en-IN", "en_GB"],
		["en-ZA", "en_GB"],
		["en-CA", "en_CA"],
		["en-AU", "en_AU"],
		["en-NZ", "en_AU"],
		["en-Latn-US", "en_US"],
		["de-DE", "de_DE"],
		["de", "de_DE"],
		["de-LU", "de_DE"],
		["de-AT", "de_AT"],
		["de-CH", "de_CH"],
		["de-LI", "de_CH"],
		["fr-FR", "fr"],
		["fr-CA", "fr"],
		["es-MX", "es_ES"],
		["it-CH", "it_IT"],
		["pt-BR", "pt_BR"],
		["pt-PT", "pt_PT"],
		["pt-AO", "pt_PT"],
		["nl-BE", "nl_NL"],
		["pl-PL", "pl_PL"],
		["ru-RU", "ru_RU"],
		["sv-FI", "sv_SE"],
		["da-DK", "da_DK"],
		["cs-CZ", "cs_CZ"],
		[" en-us ", "en_US"]
	])("%s -> %s", (tag, id) => {
		expect(dictionaryFor(tag)).toBe(id)
	})

	it.each(["", "x-none", "zxx", "und", "uk-UA", "nb-NO", "nn-NO", "tr-TR", "ja-JP", "zh-CN", "ar-SA", "enx"])("%s has none", tag => {
		expect(dictionaryFor(tag)).toBeUndefined()
	})
})
