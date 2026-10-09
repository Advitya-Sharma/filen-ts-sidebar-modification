// Files that describe a video rather than being one: subtitles, artwork, .nfo and .edl, matched by name only.

import { normalizeLanguage } from "./languages"

export type SidecarSubtitleFormat = "srt" | "vtt" | "ass" | "ssa" | "microdvd" | "smi" | "vobsub" | "pgs"

// `.sub` is VobSub when an `.idx` of the same stem sits beside it (the `.idx` then stands for the pair); otherwise it
// is presumed MicroDVD until its content says so.
export function subtitleFormat(extension: string, hasIdxPartner: boolean): SidecarSubtitleFormat | null {
	switch (extension) {
		case "srt":
		case "vtt":
		case "ass":
		case "ssa": {
			return extension
		}

		case "smi":
		case "sami": {
			return "smi"
		}

		case "idx": {
			return "vobsub"
		}

		case "sub": {
			return hasIdxPartner ? null : "microdvd"
		}

		case "sup": {
			return "pgs"
		}

		default: {
			return null
		}
	}
}

const TRACK_NUMBER = /^\d+_(?=\p{L})/u

export interface SubtitleNameInfo {
	language: string | null
	forced: boolean
	// Hearing-impaired / SDH / closed captions.
	sdh: boolean
	isDefault: boolean
	// Tokens that are neither flags nor the language, in name order ("Commentary").
	title: string
}

// Reads the dot-separated tokens after the video stem ("Movie.en.forced.srt" → ["en", "forced"]) from the right:
// flags, then the first language found, everything else is title. "hi" is ambiguous (Hindi, hearing impaired):
// it is the hearing-impaired flag when another token names the language.
export function parseSubtitleTokens(tokens: readonly string[]): SubtitleNameInfo {
	let language: string | null = null
	let forced = false
	let sdh = false
	let isDefault = false
	let sawHi = false
	const title: string[] = []

	for (let i = tokens.length - 1; i >= 0; i--) {
		// "2_English.srt" inside Subs/<video>/ numbers the track before naming it.
		const token = (tokens[i] ?? "").trim().replace(TRACK_NUMBER, "")
		const lower = token.toLowerCase()

		if (lower.length === 0) {
			continue
		}

		if (lower === "default") {
			isDefault = true
		} else if (lower === "forced" || lower === "foreign") {
			forced = true
		} else if (lower === "sdh" || lower === "cc") {
			sdh = true
		} else if (lower === "hi" && !sawHi) {
			sawHi = true
		} else {
			const resolved: string | null = language === null ? normalizeLanguage(lower) : null

			if (resolved !== null) {
				language = resolved
			} else {
				title.push(token)
			}
		}
	}

	if (sawHi) {
		if (language === null) {
			language = normalizeLanguage("hi")
		} else {
			sdh = true
		}
	}

	return {
		language,
		forced,
		sdh,
		isDefault,
		title: title.reverse().join(" ")
	}
}

const FLAG_TOKENS = new Set(["default", "forced", "foreign", "sdh", "cc", "hi"])

function isSubtitleTagToken(token: string): boolean {
	const lower = token.trim().replace(TRACK_NUMBER, "").toLowerCase()

	return FLAG_TOKENS.has(lower) || normalizeLanguage(lower) !== null
}

// Subtitle whose stem names no video ("It.2017.1080p.x264-GROUP.srt", "English.srt", "2_English.srt"): only the
// run of flag and language tokens at its end counts, and the first token is the foreign title unless it stands
// alone, so title words are never read as languages.
export function parseUnmatchedSubtitleName(stem: string): SubtitleNameInfo {
	const tokens = stem.split(".")
	let from = tokens.length

	while (from > (tokens.length > 1 ? 1 : 0) && isSubtitleTagToken(tokens[from - 1] ?? "")) {
		from--
	}

	return parseSubtitleTokens(tokens.slice(from))
}

export type ArtSlot = "poster" | "backdrop" | "logo" | "thumb"

export interface ArtName {
	slot: ArtSlot
	// Lower wins when several files fill one slot.
	rank: number
}

function art(slot: ArtSlot, rank: number): ArtName {
	return {
		slot,
		rank
	}
}

// Whole-name artwork in a title's own directory ("poster.jpg", "fanart.png").
const DIRECTORY_ART = new Map<string, ArtName>([
	["poster", art("poster", 0)],
	["folder", art("poster", 1)],
	["cover", art("poster", 2)],
	["default", art("poster", 3)],
	["movie", art("poster", 4)],
	["show", art("poster", 4)],
	["fanart", art("backdrop", 0)],
	["backdrop", art("backdrop", 1)],
	["background", art("backdrop", 2)],
	["art", art("backdrop", 3)],
	["backdrop1", art("backdrop", 4)],
	["landscape", art("backdrop", 5)],
	["thumb", art("backdrop", 6)],
	["logo", art("logo", 0)],
	["clearlogo", art("logo", 1)]
])

export function directoryArtName(lowerStem: string): ArtName | null {
	return DIRECTORY_ART.get(lowerStem) ?? null
}

// "<video stem>-<suffix>" artwork.
const STEM_ART = new Map<string, ArtName>([
	["poster", art("poster", 0)],
	["folder", art("poster", 1)],
	["cover", art("poster", 2)],
	["default", art("poster", 3)],
	["movie", art("poster", 4)],
	["fanart", art("backdrop", 0)],
	["backdrop", art("backdrop", 1)],
	["background", art("backdrop", 2)],
	["art", art("backdrop", 3)],
	["landscape", art("backdrop", 4)],
	["thumb", art("thumb", 0)],
	["logo", art("logo", 0)],
	["clearlogo", art("logo", 1)]
])

// A bare "<video stem>.jpg" is that video's poster (movies) or still (episodes).
export const STEM_IMAGE_ART: ArtName = art("poster", 5)

export interface StemArtName {
	// Lowercase stem of the video the image belongs to.
	owner: string
	art: ArtName
}

export function stemArtName(lowerStem: string): StemArtName | null {
	const dash = lowerStem.lastIndexOf("-")

	if (dash <= 0) {
		return null
	}

	const found = STEM_ART.get(lowerStem.slice(dash + 1))

	return found === undefined
		? null
		: {
				owner: lowerStem.slice(0, dash).trimEnd(),
				art: found
			}
}

const SEASON_ART = /^season[ _-]?(\d{1,4}|specials)-(poster|fanart|landscape)$/u

export interface SeasonArtName {
	season: number
	art: ArtName
}

// Season artwork kept in the series directory ("season02-poster.jpg", "season-specials-poster.jpg").
export function seasonArtName(lowerStem: string): SeasonArtName | null {
	const match = SEASON_ART.exec(lowerStem)

	if (match === null) {
		return null
	}

	return {
		season: match[1] === "specials" ? 0 : Number(match[1]),
		art: match[2] === "poster" ? art("poster", 1) : art("backdrop", 1)
	}
}

export type SidecarDirKind = "subtitles" | "metadata"

// Directories that only hold sidecars for the videos next to them.
export function sidecarDirKind(name: string): SidecarDirKind | null {
	const lower = name.toLowerCase()

	if (lower === "subs" || lower === "subtitles" || lower === "subtitle" || lower === "sub") {
		return "subtitles"
	}

	return lower === "metadata" ? "metadata" : null
}
