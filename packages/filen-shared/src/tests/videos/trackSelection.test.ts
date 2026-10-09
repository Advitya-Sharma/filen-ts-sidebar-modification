import { describe, expect, it } from "vitest"
import {
	rememberSubtitle,
	selectAudioTrack,
	selectSubtitleTrack,
	type AudioTrackInfo,
	type SubtitleMode,
	type SubtitleTrackInfo
} from "@filen/shared/videos"

const audio = (lang: string | null, isDefault = false): AudioTrackInfo => ({ lang, default: isDefault })

function sub(lang: string | null, extra: Partial<SubtitleTrackInfo> = {}): SubtitleTrackInfo {
	return {
		lang,
		forced: false,
		default: false,
		hi: false,
		external: false,
		...extra
	}
}

describe("selectAudioTrack", () => {
	const tracks = [audio("jpn", true), audio("eng"), audio("ger")]

	it("returns -1 without tracks", () => {
		expect(selectAudioTrack([], { userLanguages: ["en"] })).toBe(-1)
	})

	it("honours an explicit pick over everything", () => {
		expect(selectAudioTrack(tracks, { userLanguages: ["de"], rememberedLang: "en", explicitIndex: 0 })).toBe(0)
		expect(selectAudioTrack(tracks, { userLanguages: ["de"], explicitIndex: 9 })).toBe(2)
	})

	it("prefers the remembered series language over user languages", () => {
		expect(selectAudioTrack(tracks, { userLanguages: ["de-DE"], rememberedLang: "en" })).toBe(1)
		expect(selectAudioTrack(tracks, { userLanguages: ["de-DE"], rememberedLang: "eng" })).toBe(1)
	})

	it("ignores a remembered language the file lacks", () => {
		expect(selectAudioTrack(tracks, { userLanguages: ["de"], rememberedLang: "fr" })).toBe(2)
	})

	it("ranks by user language order, then default, then position", () => {
		expect(selectAudioTrack(tracks, { userLanguages: ["en-US", "de"] })).toBe(1)
		expect(selectAudioTrack(tracks, { userLanguages: ["fr"] })).toBe(0)
		expect(selectAudioTrack([audio("fra"), audio("ita")], { userLanguages: ["xx"] })).toBe(0)
		expect(selectAudioTrack([audio(null), audio(null, true)], { userLanguages: [] })).toBe(1)
	})

	it("prefers a playable preferred-or-default track over transcoding", () => {
		const playable = (track: AudioTrackInfo) => track.lang !== "eng"

		// preferred English needs transcode; the playable default Japanese wins
		expect(selectAudioTrack(tracks, { userLanguages: ["en"], playable })).toBe(0)
		// a playable second preferred language beats it too
		expect(selectAudioTrack(tracks, { userLanguages: ["en", "de"], playable })).toBe(2)
		// nothing playable: keep the best match
		expect(selectAudioTrack(tracks, { userLanguages: ["en"], playable: () => false })).toBe(1)
		// non-default, non-preferred playable tracks are not widened to
		expect(selectAudioTrack([audio("eng"), audio("fra")], { userLanguages: ["en"], playable: t => t.lang === "fra" })).toBe(0)
	})

	it("prefers a playable track within the remembered language", () => {
		const two = [audio("eng"), audio("eng")]

		expect(selectAudioTrack(two, { userLanguages: [], rememberedLang: "en", playable: (_t, i) => i === 1 })).toBe(1)
		expect(selectAudioTrack(two, { userLanguages: [], rememberedLang: "en", playable: () => false })).toBe(0)
	})
})

describe("selectSubtitleTrack", () => {
	const select = (tracks: SubtitleTrackInfo[], mode: SubtitleMode, audioLang: string | null, extra: object = {}) =>
		selectSubtitleTrack(tracks, { mode, audioLang, userLanguages: ["en-US"], ...extra })

	const tracks = [sub("fra"), sub("eng", { forced: true }), sub("eng"), sub(null, { forced: true }), sub("eng", { hi: true })]

	it("none disables subtitles", () => {
		expect(select(tracks, "none", "jpn")).toBe(-1)
	})

	it("smart: non-preferred audio picks the first preferred non-forced track", () => {
		expect(select(tracks, "smart", "jpn")).toBe(2)
		expect(select(tracks, "smart", null)).toBe(2)
		expect(select([sub("fra")], "smart", "jpn")).toBe(-1)
	})

	it("smart: preferred audio picks forced only (audio language first, then undefined language)", () => {
		expect(select(tracks, "smart", "eng")).toBe(1)
		expect(select([sub("eng"), sub(null, { forced: true })], "smart", "en")).toBe(1)
		expect(select([sub("eng")], "smart", "eng")).toBe(-1)
		expect(select([sub("fra", { forced: true })], "smart", "eng")).toBe(-1)
	})

	it("sorts external > default > preferred > preferred forced > forced undefined > forced", () => {
		expect(select([sub("eng"), sub("eng", { external: true })], "smart", "jpn")).toBe(1)
		expect(select([sub("eng"), sub("eng", { default: true })], "smart", "jpn")).toBe(1)
		expect(select([sub("eng", { forced: true }), sub("eng")], "smart", "jpn")).toBe(1)
		expect(select([sub("eng", { forced: true })], "smart", "jpn")).toBe(0)
		expect(select([sub("fra", { forced: true }), sub(null, { forced: true })], "onlyForced", "jpn")).toBe(1)
		expect(select([sub("fra", { forced: true }), sub("deu", { forced: true })], "onlyForced", "jpn")).toBe(0)
	})

	it("orders preferred languages by the user's ranking", () => {
		const list = [sub("deu"), sub("eng")]

		expect(selectSubtitleTrack(list, { mode: "smart", audioLang: "jpn", userLanguages: ["en", "de"] })).toBe(1)
		expect(selectSubtitleTrack(list, { mode: "smart", audioLang: "jpn", userLanguages: ["de", "en"] })).toBe(0)
	})

	it("onlyForced ignores everything else", () => {
		expect(select(tracks, "onlyForced", "jpn")).toBe(1)
		expect(select([sub("eng"), sub("eng", { default: true })], "onlyForced", "jpn")).toBe(-1)
	})

	it("default picks default or forced flagged tracks", () => {
		expect(select([sub("eng"), sub("fra", { default: true })], "default", "jpn")).toBe(1)
		expect(select([sub("eng"), sub("fra")], "default", "jpn")).toBe(-1)
		expect(select([sub("eng"), sub("fra", { forced: true })], "default", "jpn")).toBe(1)
	})

	it("never selects unsupported (image) tracks", () => {
		expect(select([sub("eng", { unsupported: true, default: true }), sub("eng")], "smart", "jpn")).toBe(1)
		expect(select([sub("eng", { unsupported: true })], "smart", "jpn")).toBe(-1)
		expect(
			selectSubtitleTrack([sub("eng", { unsupported: true })], {
				mode: "smart",
				audioLang: null,
				userLanguages: ["en"],
				explicitIndex: 0
			})
		).toBe(-1)
	})

	it("remembered choices are language based and beat the mode", () => {
		const list = [sub("eng", { hi: true }), sub("eng"), sub("fra"), sub("fra", { forced: true })]
		const remembered = rememberSubtitle(list[2] as SubtitleTrackInfo)

		expect(remembered).toEqual({ lang: "fr", forced: false, hi: false })
		expect(select(list, "none", "jpn", { remembered })).toBe(2)
		expect(select(list, "smart", "jpn", { remembered: { lang: "en", forced: false, hi: true } })).toBe(0)
		expect(select(list, "smart", "jpn", { remembered: { lang: "en", forced: false, hi: false } })).toBe(1)
		// hi flag differs, language and forced still match
		expect(select([sub("eng", { hi: true })], "none", "jpn", { remembered: { lang: "en", forced: false, hi: false } })).toBe(0)
		// remembered forced does not match a non-forced track of the same language only when forced exists
		expect(select(list, "none", "jpn", { remembered: { lang: "fr", forced: true, hi: false } })).toBe(3)
		// layout change: different order, same language
		expect(select([sub("fra"), sub("deu")], "none", "jpn", { remembered: { lang: "fr", forced: false, hi: false } })).toBe(0)
	})

	it("remembered off and explicit choices win", () => {
		expect(select(tracks, "smart", "jpn", { remembered: "none" })).toBe(-1)
		expect(select(tracks, "smart", "jpn", { remembered: "none", explicitIndex: 0 })).toBe(0)
		expect(select(tracks, "smart", "jpn", { explicitIndex: -1 })).toBe(-1)
		expect(select(tracks, "smart", "jpn", { explicitIndex: 99 })).toBe(-1)
	})

	it("falls through when the remembered language is absent", () => {
		expect(select(tracks, "smart", "jpn", { remembered: { lang: "ja", forced: false, hi: false } })).toBe(2)
	})
})

describe("selectSubtitleTrack default mode", () => {
	it("picks an external track like a default or forced one", () => {
		const list = [sub("fra"), sub("eng", { external: true })]

		expect(selectSubtitleTrack(list, { mode: "default", audioLang: "jpn", userLanguages: ["en"] })).toBe(1)
		expect(selectSubtitleTrack([sub("fra"), sub("eng")], { mode: "default", audioLang: "jpn", userLanguages: ["en"] })).toBe(-1)
	})
})
