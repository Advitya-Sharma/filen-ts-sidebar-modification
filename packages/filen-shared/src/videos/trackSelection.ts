import { languagePreferenceRank, languagesMatch, normalizeLanguage } from "./languages"

export interface AudioTrackInfo {
	lang: string | null
	default: boolean
}

export interface SelectAudioOptions {
	// Normalised or raw codes, most preferred first (e.g. `navigator.languages`).
	userLanguages: readonly string[]
	// Language remembered for this series (never an index).
	rememberedLang?: string | null
	// A track the user picked in this session; wins over everything.
	explicitIndex?: number | null
	// Whether the browser can play a track without transcoding.
	playable?: (track: AudioTrackInfo, index: number) => boolean
}

function normalizedPreferences(list: readonly string[]): string[] {
	const out: string[] = []

	for (const entry of list) {
		const normalized = normalizeLanguage(entry)

		if (normalized !== null && !out.includes(normalized)) {
			out.push(normalized)
		}
	}

	return out
}

// Index into `tracks`, -1 when there are none.
// Order: explicit pick, remembered series language, then preferred language / default; with no explicit choice a natively
// playable preferred-or-default track beats one that would need transcoding.
export function selectAudioTrack(tracks: readonly AudioTrackInfo[], options: SelectAudioOptions): number {
	const count = tracks.length

	if (count === 0) {
		return -1
	}

	const { explicitIndex, rememberedLang, playable } = options

	if (explicitIndex !== undefined && explicitIndex !== null && explicitIndex >= 0 && explicitIndex < count) {
		return explicitIndex
	}

	const isPlayable = (i: number): boolean => playable === undefined || playable(tracks[i] as AudioTrackInfo, i)

	if (rememberedLang !== undefined && rememberedLang !== null) {
		let first = -1

		for (let i = 0; i < count; i++) {
			if (languagesMatch((tracks[i] as AudioTrackInfo).lang, rememberedLang)) {
				if (first === -1) {
					first = i
				}

				if (isPlayable(i)) {
					return i
				}
			}
		}

		if (first !== -1) {
			return first
		}
	}

	const preferences = normalizedPreferences(options.userLanguages)
	const rank = (i: number): number => {
		const r = languagePreferenceRank((tracks[i] as AudioTrackInfo).lang, preferences)

		return r === -1 ? preferences.length : r
	}

	const order: number[] = []

	for (let i = 0; i < count; i++) {
		order.push(i)
	}

	order.sort(
		(a, b) =>
			rank(a) - rank(b) || Number((tracks[b] as AudioTrackInfo).default) - Number((tracks[a] as AudioTrackInfo).default) || a - b
	)

	const best = order[0] as number

	if (playable === undefined || isPlayable(best)) {
		return best
	}

	for (const i of order) {
		if ((rank(i) < preferences.length || (tracks[i] as AudioTrackInfo).default) && isPlayable(i)) {
			return i
		}
	}

	return best
}

export interface SubtitleTrackInfo {
	lang: string | null
	forced: boolean
	default: boolean
	// Hearing impaired / SDH / CC.
	hi: boolean
	external: boolean
	// Image formats (PGS, VobSub) and anything else that cannot be rendered.
	unsupported?: boolean
}

export type SubtitleMode = "smart" | "default" | "onlyForced" | "none"

// A remembered choice, language-based so it survives different track layouts.
export interface RememberedSubtitle {
	lang: string | null
	forced: boolean
	hi: boolean
}

export interface SelectSubtitleOptions {
	mode: SubtitleMode
	// Language of the selected audio track.
	audioLang: string | null
	userLanguages: readonly string[]
	// Remembered per-series choice; "none" = the user turned subtitles off.
	remembered?: RememberedSubtitle | "none" | null
	// A track the user picked in this session, or -1 for off; wins over everything.
	explicitIndex?: number | null
}

export function rememberSubtitle(track: SubtitleTrackInfo): RememberedSubtitle {
	return {
		lang: normalizeLanguage(track.lang),
		forced: track.forced,
		hi: track.hi
	}
}

// Sort tiers: external, default, preferred non-forced, preferred forced, forced with undefined language, forced, rest.
function tierOf(track: SubtitleTrackInfo, preferences: readonly string[]): number {
	if (track.external) {
		return 0
	}

	if (track.default) {
		return 1
	}

	const preferred = languagePreferenceRank(track.lang, preferences) !== -1

	if (preferred && !track.forced) {
		return 2
	}

	if (preferred && track.forced) {
		return 3
	}

	if (track.forced && normalizeLanguage(track.lang) === null) {
		return 4
	}

	return track.forced ? 5 : 6
}

function findRemembered(tracks: readonly SubtitleTrackInfo[], remembered: RememberedSubtitle): number {
	let langForced = -1
	let langOnly = -1

	for (let i = 0; i < tracks.length; i++) {
		const track = tracks[i] as SubtitleTrackInfo

		if (track.unsupported === true || normalizeLanguage(track.lang) !== remembered.lang) {
			continue
		}

		if (track.forced === remembered.forced) {
			if (track.hi === remembered.hi) {
				return i
			}

			if (langForced === -1) {
				langForced = i
			}
		} else if (langOnly === -1 && !track.forced) {
			langOnly = i
		}
	}

	return langForced !== -1 ? langForced : langOnly
}

// Index into `tracks`, -1 for off.
export function selectSubtitleTrack(tracks: readonly SubtitleTrackInfo[], options: SelectSubtitleOptions): number {
	const { explicitIndex, remembered, mode, audioLang } = options

	if (explicitIndex !== undefined && explicitIndex !== null) {
		return explicitIndex >= 0 && explicitIndex < tracks.length && (tracks[explicitIndex] as SubtitleTrackInfo).unsupported !== true
			? explicitIndex
			: -1
	}

	if (remembered === "none") {
		return -1
	}

	if (remembered !== undefined && remembered !== null) {
		const found = findRemembered(tracks, remembered)

		if (found !== -1) {
			return found
		}
	}

	if (mode === "none") {
		return -1
	}

	const preferences = normalizedPreferences(options.userLanguages)
	const sorted: number[] = []

	for (let i = 0; i < tracks.length; i++) {
		if ((tracks[i] as SubtitleTrackInfo).unsupported !== true) {
			sorted.push(i)
		}
	}

	const tier = (i: number): number => tierOf(tracks[i] as SubtitleTrackInfo, preferences)
	const rank = (i: number): number => {
		const r = languagePreferenceRank((tracks[i] as SubtitleTrackInfo).lang, preferences)

		return r === -1 ? preferences.length : r
	}

	sorted.sort((a, b) => tier(a) - tier(b) || rank(a) - rank(b) || a - b)

	const get = (i: number): SubtitleTrackInfo => tracks[i] as SubtitleTrackInfo

	if (mode === "onlyForced") {
		return sorted.find(i => get(i).forced) ?? -1
	}

	if (mode === "default") {
		return sorted.find(i => get(i).default || get(i).forced || get(i).external) ?? -1
	}

	// Smart: audio already in a preferred language needs only forced subtitles (foreign parts); otherwise the first preferred track.
	const audioNormalized = normalizeLanguage(audioLang)
	const audioPreferred = audioNormalized !== null && preferences.includes(audioNormalized)

	if (audioPreferred) {
		return (
			sorted.find(i => get(i).forced && languagesMatch(get(i).lang, audioLang)) ??
			sorted.find(i => get(i).forced && normalizeLanguage(get(i).lang) === null) ??
			-1
		)
	}

	return sorted.find(i => languagePreferenceRank(get(i).lang, preferences) !== -1) ?? -1
}
