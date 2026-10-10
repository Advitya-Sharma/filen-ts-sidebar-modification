import {
	EXTRA_TYPE_ORDER,
	type HomeEntry,
	type HomeIndex,
	type LadderPath,
	type SeriesPrimaryAction,
	type VideoProbe
} from "@filen/shared/videos"
import type { TFunction } from "i18next"
import { qualityLabel } from "@/features/videos/lib/videoGrid.logic"
import type {
	VideoExtraView,
	VideoHomeItem,
	VideoPartView,
	VideoSeasonView,
	VideoTitleDetail,
	VideoTitleView,
	VideoVersionView
} from "@/workers/videoIndex/view.logic"

// The movie and series pages' decisions that need no DOM: which header shows, what the probe-dependent slots hold,
// the primary button, the version picker, the extras shelf and the season selector.

// The cinematic header only for a real backdrop: a sidecar image the index found (zero bytes to know), or a landscape
// cover inside the file that its stored probe already located (a local read); a generated frame is never stretched to
// fill it. Everything else gets the poster header over a dominant-colour band.
export type DetailHeaderMode = "backdrop" | "band"

export function detailHeaderMode(title: Pick<VideoTitleView, "backdrop">, landscapeCover = false): DetailHeaderMode {
	return title.backdrop !== null || landscapeCover ? "backdrop" : "band"
}

// A slot whose content comes from the probe: reserved while the probe is out, then filled or empty for good.
export type Slot<T> = { state: "pending" } | { state: "empty" } | { state: "ready"; value: T }

export interface DetailSlots {
	runtimeSec: Slot<number>
	badges: Slot<string[]>
	endsAtMs: Slot<number>
	// Audio tracks and subtitle choices to pick from.
	tracks: Slot<true>
}

function slot<T>(value: T | null): Slot<T> {
	return value === null ? { state: "empty" } : { state: "ready", value }
}

const PENDING: Slot<never> = { state: "pending" }

function hasSubtitles(part: VideoPartView | null, probe: VideoProbe): boolean {
	return (part?.subtitles.length ?? 0) > 0 || probe.subs.length > 0
}

// The probe-dependent fields of a playable item. `probe` undefined: not read yet (every slot pending); null: no probe
// can be had (only what the index and the NFO know). The runtime falls back to the NFO's.
export function detailSlots({
	probe,
	part,
	height,
	nfoRuntimeMinutes,
	resumeSec,
	now
}: {
	probe: VideoProbe | null | undefined
	part: VideoPartView | null
	// Picture height the name advertises, 0 when it says nothing.
	height: number
	nfoRuntimeMinutes: number | null
	resumeSec: number
	now: number
}): DetailSlots {
	if (probe === undefined) {
		return {
			runtimeSec: nfoRuntimeMinutes === null ? PENDING : { state: "ready", value: nfoRuntimeMinutes * 60 },
			badges: PENDING,
			endsAtMs: PENDING,
			tracks: PENDING
		}
	}

	const duration = probe?.durationSec ?? (nfoRuntimeMinutes === null ? null : nfoRuntimeMinutes * 60)
	const badges: string[] = []

	if (probe !== null) {
		const quality = qualityLabel(probe.video?.h ?? height)

		if (quality.length > 0) {
			badges.push(quality)
		}

		if (probe.video?.dv !== undefined) {
			badges.push("DV")
		} else if (probe.video?.transfer === "pq" || probe.video?.transfer === "hlg") {
			badges.push("HDR")
		}

		const channels = probe.audio.find(track => track.def)?.ch ?? probe.audio[0]?.ch ?? null

		if (channels !== null && channels >= 6) {
			badges.push(`${String(channels - 1)}.1`)
		}

		if (hasSubtitles(part, probe)) {
			badges.push("CC")
		}
	} else if (qualityLabel(height).length > 0) {
		badges.push(qualityLabel(height))
	}

	return {
		runtimeSec: slot(duration !== null && duration > 0 ? duration : null),
		badges: slot(badges.length > 0 ? badges : null),
		endsAtMs: slot(duration !== null && duration > resumeSec ? now + (duration - resumeSec) * 1000 : null),
		tracks: slot(probe !== null && (probe.audio.length > 0 || hasSubtitles(part, probe)) ? true : null)
	}
}

// ── Movie ────────────────────────────────────────────────────────────────────

export type MovieAction =
	| { kind: "play"; stableUuid: string }
	| { kind: "resume"; stableUuid: string; positionSec: number; durationSec: number; progress: number; remainingSec: number }

// The stable uuid of each version's first part: what the watch state is keyed by and Play opens.
export function versionKeys(versions: readonly VideoVersionView[]): string[] {
	return versions.flatMap(version => {
		const first = version.parts[0]

		return first === undefined ? [] : [first.file.stableUuid]
	})
}

// The version the page starts on: the most recently played one, else the primary.
export function defaultVersion(versions: readonly VideoVersionView[], entries: ReadonlyMap<string, HomeEntry>): number {
	let best = 0
	let bestPlayed = 0

	versions.forEach((version, index) => {
		const key = version.parts[0]?.file.stableUuid
		const played = key === undefined ? 0 : (entries.get(key)?.state.lastPlayedMs ?? 0)

		if (played > bestPlayed) {
			best = index
			bestPlayed = played
		}
	})

	return best
}

// Play, or "Resume 42:10" with its bar when the chosen version has a resume point.
export function moviePrimaryAction(stableUuid: string, entry: HomeEntry | undefined): MovieAction {
	const state = entry?.state

	if (state === undefined || !state.inProgress || !(state.durationSec > 0)) {
		return { kind: "play", stableUuid }
	}

	return {
		kind: "resume",
		stableUuid,
		positionSec: state.positionSec,
		durationSec: state.durationSec,
		progress: Math.min(1, Math.max(0, state.positionSec / state.durationSec)),
		remainingSec: Math.max(0, state.durationSec - state.positionSec)
	}
}

// The watch state of every video of the index by stable uuid (one pass).
export function entriesByStableUuid(index: HomeIndex): Map<string, HomeEntry> {
	return new Map(index.entries.map(entry => [entry.video.id, entry]))
}

// A picker only when there is something to pick.
export function showsVersionPicker(versions: readonly VideoVersionView[]): boolean {
	return versions.length >= 2
}

export type LadderBadge = "direct" | "remux" | "unsupported"

// What the version picker says of a version once its probe is known: played as it is, or repackaged.
export function ladderBadge(path: LadderPath): LadderBadge {
	switch (path) {
		case "direct":
			return "direct"
		case "remux":
		case "remux+transcode":
		case "video-only":
			return "remux"
		case "unsupported":
			return "unsupported"
	}
}

// Trailers first, then featurettes and so on (samples never reach the shelf); stable within a type.
export function extrasShelf(extras: readonly VideoExtraView[]): VideoExtraView[] {
	const rank = (extra: VideoExtraView): number => {
		const at = EXTRA_TYPE_ORDER.indexOf(extra.type)

		return at === -1 ? EXTRA_TYPE_ORDER.length : at
	}

	return extras
		.map((extra, index) => ({ extra, index }))
		.sort((a, b) => rank(a.extra) - rank(b.extra) || a.index - b.index)
		.map(entry => entry.extra)
}

// ── Series ───────────────────────────────────────────────────────────────────

// Tabs up to this many regular seasons, a select past it.
export const SEASON_TABS_MAX = 6

export interface SeasonLayout {
	// Numbered seasons, ascending.
	regular: VideoSeasonView[]
	// Season 0: last and collapsed; shown as the only list when there is nothing else.
	specials: VideoSeasonView | null
	selector: "tabs" | "select" | "none"
}

export function seasonLayout(seasons: readonly VideoSeasonView[]): SeasonLayout {
	const regular = seasons.filter(season => season.number > 0 && season.episodes.length > 0).sort((a, b) => a.number - b.number)
	const specials = seasons.find(season => season.number === 0 && season.episodes.length > 0) ?? null

	return {
		regular,
		specials,
		selector: regular.length <= 1 ? "none" : regular.length <= SEASON_TABS_MAX ? "tabs" : "select"
	}
}

// The season the page opens on: the Next Up episode's, else the one last viewed here, else the first.
export function defaultSeason({
	regular,
	action,
	lastViewed
}: {
	regular: readonly number[]
	action: SeriesPrimaryAction | null
	lastViewed: number | null
}): number | null {
	const next = action === null || action.type === "playFirst" ? null : (action.entry.video.season ?? null)

	if (next !== null && regular.includes(next)) {
		return next
	}

	if (lastViewed !== null && regular.includes(lastViewed)) {
		return lastViewed
	}

	return regular[0] ?? null
}

export type SeriesButton =
	| { kind: "resume"; stableUuid: string; season: number; episode: number; remainingSec: number }
	| { kind: "next"; stableUuid: string; season: number; episode: number }
	| { kind: "play"; stableUuid: string }

// "Resume S2E5 · 21 min left", "Play S2E6", or "Play" (nothing watched yet, or everything watched).
export function seriesButton(action: SeriesPrimaryAction | null): SeriesButton | null {
	if (action === null) {
		return null
	}

	const { video, state } = action.entry
	const season = video.season ?? 0
	const episode = video.episode ?? 0

	if (action.type === "resume" && episode > 0) {
		return { kind: "resume", stableUuid: video.id, season, episode, remainingSec: Math.max(0, state.durationSec - state.positionSec) }
	}

	if (action.type === "play" && episode > 0) {
		return { kind: "next", stableUuid: video.id, season, episode }
	}

	return { kind: "play", stableUuid: video.id }
}

// Every version's first file of a season's episodes: what Mark season watched writes, as one edit.
export function seasonKeys(season: VideoSeasonView): string[] {
	return season.episodes.flatMap(episode => versionKeys(episode.versions))
}

export interface SeasonProgress {
	unwatched: number
	watched: boolean
}

// Unplayed episodes of a season (one with a resume point counts as unplayed) and whether all are played.
export function seasonProgress(season: VideoSeasonView, entries: ReadonlyMap<string, HomeEntry>): SeasonProgress {
	let unwatched = 0

	for (const episode of season.episodes) {
		const played = versionKeys(episode.versions).some(key => entries.get(key)?.state.played === true)

		if (!played) {
			unwatched++
		}
	}

	return { unwatched, watched: season.episodes.length > 0 && unwatched === 0 }
}

export interface EpisodeWatch {
	played: boolean
	// 0-1 with a resume point
	progress: number | null
	remainingSec: number | null
	// The version Play opens: the one last played, else the first.
	stableUuid: string | null
}

export function episodeWatch(versions: readonly VideoVersionView[], entries: ReadonlyMap<string, HomeEntry>): EpisodeWatch {
	const keys = versionKeys(versions)
	let latest: HomeEntry | undefined

	for (const key of keys) {
		const entry = entries.get(key)

		if (entry !== undefined && entry.state.lastPlayedMs > (latest?.state.lastPlayedMs ?? 0)) {
			latest = entry
		}
	}

	const state = latest?.state
	const resumable = state !== undefined && state.inProgress && state.durationSec > 0

	return {
		played: keys.some(key => entries.get(key)?.state.played === true),
		progress: resumable ? Math.min(1, Math.max(0, state.positionSec / state.durationSec)) : null,
		remainingSec: resumable ? Math.max(0, state.durationSec - state.positionSec) : null,
		stableUuid: latest?.video.id ?? keys[0] ?? null
	}
}

// "2h 44m", "51m": a runtime or what is left of one.
export function durationText(seconds: number, t: TFunction<"videos">): string {
	const minutes = Math.max(1, Math.round(seconds / 60))

	return minutes >= 60
		? t("videosHomeRuntimeHours", { hours: Math.floor(minutes / 60), minutes: minutes % 60 })
		: t("videosHomeRuntimeMinutes", { minutes })
}

// "4m left", "1h 31m left": what is left of a resume point, the same on every title page.
export function timeLeftText(seconds: number, t: TFunction<"videos">): string {
	return t("detailTimeLeft", { time: durationText(seconds, t) })
}

// The seasons a series' counts and Mark series watched cover: the numbered ones, or the specials when that is all
// there is.
export function countedSeasons(layout: SeasonLayout): VideoSeasonView[] {
	return layout.regular.length > 0 ? layout.regular : layout.specials === null ? [] : [layout.specials]
}

function homeItem(
	detail: VideoTitleDetail,
	version: VideoVersionView,
	fields: Pick<VideoHomeItem, "itemClass" | "season" | "episode" | "episodeEnd" | "episodeTitle" | "group" | "dateAdded">
): VideoHomeItem[] {
	const first = version.parts[0]

	return first === undefined
		? []
		: [
				{
					...fields,
					titleId: detail.title.id,
					stableUuid: first.file.stableUuid,
					uuid: first.file.uuid,
					size: first.file.size,
					label: version.label,
					unsupported: first.unsupported,
					folder: ""
				}
			]
}

// One title's videos as the watch selectors take them (one per version, its first part): what a title page derives
// its buttons and progress from, without the whole library's list.
export function titleHomeItems(detail: VideoTitleDetail): VideoHomeItem[] {
	const { title } = detail
	const plain = { season: 0, episode: 0, episodeEnd: 0, episodeTitle: "", group: title.id, dateAdded: title.dateAdded }

	switch (detail.kind) {
		case "movie":
			return detail.versions.flatMap(version => homeItem(detail, version, { ...plain, itemClass: "movie" }))
		case "other":
			return detail.part === null ? [] : homeItem(detail, { label: "", parts: [detail.part] }, { ...plain, itemClass: "other" })
		case "series":
			return detail.seasons.flatMap(season =>
				season.episodes.flatMap((episode, index) =>
					episode.versions.flatMap(version =>
						homeItem(detail, version, {
							itemClass: "episode",
							season: episode.season,
							episode: episode.episode,
							episodeEnd: episode.episodeEnd,
							episodeTitle: episode.title,
							group: `${title.id}/${String(season.number)}/${String(index)}`,
							dateAdded: episode.dateAdded
						})
					)
				)
			)
	}
}
