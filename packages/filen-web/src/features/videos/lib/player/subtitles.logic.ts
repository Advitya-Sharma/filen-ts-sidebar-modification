import {
	normalizeLanguage,
	selectSubtitleTrack,
	type RememberedSubtitle,
	type SidecarSubtitleFormat,
	type SubtitleFormat,
	type SubtitleTrackInfo
} from "@filen/shared/videos"
import type { VideoProbe } from "@/lib/media/probe/types"
import { isTapTextCodec } from "@/workers/remux/textTap"
import type { VideoFileRef, VideoPartView } from "@/workers/videoIndex/view.logic"

// The subtitles a part offers: sidecars beside it (read when chosen, one read each) and the tracks inside the
// file. Text tracks inside a Matroska file come from the remux engine's tap over what playback reads; image
// formats (PGS, VobSub) and text inside an MP4 can't be shown.

export type SubtitleUnavailable = "image" | "embedded"

export interface PlayerSubtitleTrack {
	// "s:<stable uuid>" for a sidecar, "e:<track id>" for a track in the file.
	id: string
	lang: string | null
	// Title from the name or the track ("Commentary"); may be empty.
	title: string
	forced: boolean
	hi: boolean
	isDefault: boolean
	external: boolean
	file: VideoFileRef | null
	format: SidecarSubtitleFormat | null
	unavailable: SubtitleUnavailable | null
}

// What the subtitle menu's choice is: Smart (resolved per item), off, or one track.
export type SubtitleChoice = { kind: "auto" } | { kind: "off" } | { kind: "track"; id: string }

const IMAGE_CODECS = /PGS|VOBSUB|DVB|HDMV|S_IMAGE/i

export function subtitleTracksOf(part: VideoPartView, probe: VideoProbe | null): PlayerSubtitleTrack[] {
	const tracks: PlayerSubtitleTrack[] = part.subtitles.map(subtitle => ({
		id: `s:${subtitle.file.stableUuid}`,
		lang: subtitle.language,
		title: subtitle.title,
		forced: subtitle.forced,
		hi: subtitle.sdh,
		isDefault: subtitle.isDefault,
		external: true,
		file: subtitle.file,
		format: subtitle.format,
		unavailable: subtitle.format === "pgs" || subtitle.format === "vobsub" ? "image" : null
	}))

	for (const sub of probe?.subs ?? []) {
		tracks.push({
			id: `e:${String(sub.id)}`,
			lang: sub.lang,
			title: sub.name ?? "",
			forced: sub.forced,
			hi: sub.hi,
			isDefault: sub.def,
			external: false,
			file: null,
			format: null,
			unavailable: IMAGE_CODECS.test(sub.codecId)
				? "image"
				: (probe?.container === "mkv" || probe?.container === "webm") && isTapTextCodec(sub.codecId)
					? null
					: "embedded"
		})
	}

	return tracks
}

function selectionInfo(track: PlayerSubtitleTrack): SubtitleTrackInfo {
	return {
		lang: track.lang,
		forced: track.forced,
		default: track.isDefault,
		hi: track.hi,
		external: track.external,
		unsupported: track.unavailable !== null
	}
}

// The track a choice shows; null for none. Smart takes the series' remembered choice into account.
export function resolveSubtitleChoice(params: {
	choice: SubtitleChoice
	tracks: readonly PlayerSubtitleTrack[]
	audioLang: string | null
	userLanguages: readonly string[]
	remembered: RememberedSubtitle | "none" | null
}): PlayerSubtitleTrack | null {
	const { choice, tracks } = params

	if (choice.kind === "off") {
		return null
	}

	if (choice.kind === "track") {
		const track = tracks.find(entry => entry.id === choice.id)

		return track?.unavailable === null ? track : null
	}

	const index = selectSubtitleTrack(tracks.map(selectionInfo), {
		mode: "smart",
		audioLang: params.audioLang,
		userLanguages: params.userLanguages,
		remembered: params.remembered
	})

	return tracks[index] ?? null
}

// The series' remembered subtitle is stored as one register value: the language, plus "|f" for forced and "|h"
// for hearing impaired; null = off.
export function encodeRememberedSubtitle(track: PlayerSubtitleTrack | null): string | null {
	if (track === null) {
		return null
	}

	const lang = normalizeLanguage(track.lang) ?? "und"

	return `${lang}${track.forced ? "|f" : ""}${track.hi ? "|h" : ""}`
}

export function decodeRememberedSubtitle(value: string | null | undefined): RememberedSubtitle | "none" | null {
	if (value === undefined) {
		return null
	}

	if (value === null) {
		return "none"
	}

	const [lang = "", ...flags] = value.split("|")

	return {
		lang: lang === "und" || lang.length === 0 ? null : normalizeLanguage(lang),
		forced: flags.includes("f"),
		hi: flags.includes("h")
	}
}

// The parser for a sidecar's format; `.sub` is told from VobSub by its content.
export function parseFormatOf(format: SidecarSubtitleFormat): SubtitleFormat | "sub" | null {
	switch (format) {
		case "srt":
		case "vtt":
		case "smi":
			return format
		case "ass":
		case "ssa":
			return "ass"
		case "microdvd":
			return "sub"
		default:
			return null
	}
}

// The track C brings back after subtitles were turned off: the same track when the item still has it, else one
// in the same language with the same forced and hearing-impaired flags (the next episode's sidecar). Null when
// nothing matches.
export function restoreSubtitleTrack(tracks: readonly PlayerSubtitleTrack[], last: PlayerSubtitleTrack | null): PlayerSubtitleTrack | null {
	if (last === null) {
		return null
	}

	const usable = tracks.filter(track => track.unavailable === null)
	const lang = normalizeLanguage(last.lang)

	return (
		usable.find(track => track.id === last.id) ??
		usable.find(track => normalizeLanguage(track.lang) === lang && track.forced === last.forced && track.hi === last.hi) ??
		usable.find(track => normalizeLanguage(track.lang) === lang && !track.forced) ??
		null
	)
}

// The embedded track number a choice shows, null for a sidecar or none.
export function embeddedTrackNumber(track: PlayerSubtitleTrack | null): number | null {
	if (track === null || track.external || !track.id.startsWith("e:")) {
		return null
	}

	const number = Number(track.id.slice(2))

	return Number.isSafeInteger(number) ? number : null
}

// A subtitle choice holds for the item it was made on; any other item starts from Smart.
export interface ItemSubtitleChoice {
	key: string | null
	choice: SubtitleChoice
}

export function subtitleChoiceFor(state: ItemSubtitleChoice, itemKey: string | null): SubtitleChoice {
	return state.key !== null && state.key === itemKey ? state.choice : { kind: "auto" }
}
