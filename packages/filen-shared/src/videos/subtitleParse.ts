import { MAX_SUBTITLE_CHARS, applyCueOffset, type CueStore } from "./subtitles/cues"
import { parseAss } from "./subtitles/ass"
import { DEFAULT_MICRODVD_FPS, isMicroDvd, parseMicroDvd } from "./subtitles/microdvd"
import { parseSmi } from "./smi"
import { parseSrt } from "./subtitles/srt"
import { parseVtt } from "./subtitles/vtt"

export type SubtitleFormat = "srt" | "vtt" | "ass" | "microdvd" | "smi"

// By file extension; `.sub` needs the content check (`isMicroDvd`) to tell MicroDVD from VobSub.
export function subtitleFormatFromExtension(extension: string): SubtitleFormat | "sub" | null {
	switch (extension.replace(/^\./, "").toLowerCase()) {
		case "srt":
			return "srt"
		case "vtt":
			return "vtt"
		case "ass":
		case "ssa":
			return "ass"
		case "smi":
		case "sami":
			return "smi"
		case "sub":
			return "sub"
		default:
			return null
	}
}

export interface ParseSubtitleOptions {
	fps?: number
	// Seconds, positive = later.
	offset?: number
	// SAMI: pick the track with this class or language; defaults to the first.
	smiClass?: string
	smiLang?: string
}

// Parses text of a known format into one cue store; null when refused (over the size cap, VobSub `.sub`, no SAMI track).
export function parseSubtitle(text: string, format: SubtitleFormat | "sub", options: ParseSubtitleOptions = {}): CueStore | null {
	if (text.length > MAX_SUBTITLE_CHARS) {
		return null
	}

	let store: CueStore | null

	switch (format) {
		case "srt":
			store = parseSrt(text)

			break
		case "vtt":
			store = parseVtt(text)

			break
		case "ass":
			store = parseAss(text)

			break
		case "microdvd":
			store = parseMicroDvd(text, options.fps ?? DEFAULT_MICRODVD_FPS)

			break
		case "sub":
			store = isMicroDvd(text) ? parseMicroDvd(text, options.fps ?? DEFAULT_MICRODVD_FPS) : null

			break
		case "smi": {
			const tracks = parseSmi(text)

			if (tracks === null || tracks.length === 0) {
				return null
			}

			const picked =
				(options.smiClass !== undefined ? tracks.find(track => track.className === options.smiClass) : undefined) ??
				(options.smiLang !== undefined ? tracks.find(track => track.lang === options.smiLang) : undefined) ??
				tracks[0]

			store = picked === undefined ? null : picked.store

			break
		}
	}

	return store === null || options.offset === undefined ? store : applyCueOffset(store, options.offset)
}
