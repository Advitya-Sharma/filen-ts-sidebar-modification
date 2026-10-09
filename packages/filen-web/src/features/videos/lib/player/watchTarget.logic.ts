import type { WatchItemClass } from "@filen/shared/videos"
import type { VideoEpisodeView, VideoPartView, VideoTitleDetail, VideoVersionView } from "@/workers/videoIndex/view.logic"

// What the player plays, resolved from a title's detail (the local index; no request): the item a route
// names, the episodes around it, and the version to keep across episodes.

export interface WatchEpisodeInfo {
	season: number
	episode: number
	episodeEnd: number
	title: string
	// Position in the series' play order.
	order: number
}

export interface WatchItemTarget {
	// The watch-state key: the first part's stable uuid.
	key: string
	titleId: string
	titleName: string
	kind: "movie" | "episode" | "other" | "extra"
	itemClass: WatchItemClass
	versionLabel: string
	// More than one for a stack (cd1, cd2, ...).
	parts: VideoPartView[]
	episode: WatchEpisodeInfo | null
	// The series directory, for the remembered languages.
	seriesDirUuid: string | null
}

// The series' episodes in play order: seasons ascending with specials (season 0) after the rest, episodes
// ascending within a season (unnumbered ones last, as listed).
export function episodeOrder(detail: Extract<VideoTitleDetail, { kind: "series" }>): VideoEpisodeView[] {
	const seasons = [...detail.seasons].sort((a, b) => (a.number === 0 ? 1 : 0) - (b.number === 0 ? 1 : 0) || a.number - b.number)
	const out: VideoEpisodeView[] = []

	for (const season of seasons) {
		const numbered = season.episodes.filter(episode => episode.episode > 0).sort((a, b) => a.episode - b.episode)
		const rest = season.episodes.filter(episode => episode.episode <= 0)

		out.push(...numbered, ...rest)
	}

	return out
}

function firstStable(version: VideoVersionView): string | null {
	return version.parts[0]?.file.stableUuid ?? null
}

function hasFile(version: VideoVersionView, stableUuid: string): boolean {
	return version.parts.some(part => part.file.stableUuid === stableUuid)
}

// The version of `versions` holding `stableUuid`, else the one labelled like `preferLabel`, else the first.
export function pickVersion(
	versions: readonly VideoVersionView[],
	stableUuid: string | null,
	preferLabel: string | null
): VideoVersionView | null {
	if (stableUuid !== null) {
		const holding = versions.find(version => hasFile(version, stableUuid))

		if (holding !== undefined) {
			return holding
		}
	}

	if (preferLabel !== null) {
		const labelled = versions.find(version => version.label === preferLabel)

		if (labelled !== undefined) {
			return labelled
		}
	}

	return versions[0] ?? null
}

function episodeTarget(
	detail: Extract<VideoTitleDetail, { kind: "series" }>,
	order: readonly VideoEpisodeView[],
	index: number,
	version: VideoVersionView
): WatchItemTarget | null {
	const episode = order[index]
	const key = firstStable(version)

	if (episode === undefined || key === null) {
		return null
	}

	return {
		key,
		titleId: detail.title.id,
		titleName: detail.title.title,
		kind: "episode",
		itemClass: "episode",
		versionLabel: version.label,
		parts: version.parts,
		episode: {
			season: episode.season,
			episode: episode.episode,
			episodeEnd: episode.episodeEnd,
			title: episode.title,
			order: index
		},
		seriesDirUuid: detail.title.id
	}
}

// The item `stableUuid` names within the title (any part of any version), else what the title plays first.
export function resolveWatchTarget(detail: VideoTitleDetail, stableUuid: string | null): WatchItemTarget | null {
	const title = detail.title

	if (detail.kind === "series") {
		const order = episodeOrder(detail)
		let index = stableUuid === null ? -1 : order.findIndex(episode => episode.versions.some(version => hasFile(version, stableUuid)))

		if (index === -1) {
			const primary = title.primary?.stableUuid ?? null

			index =
				primary === null
					? 0
					: Math.max(
							0,
							order.findIndex(episode => episode.versions.some(version => hasFile(version, primary)))
						)
		}

		const episode = order[index]
		const version = episode === undefined ? null : pickVersion(episode.versions, stableUuid, null)

		return version === null ? null : episodeTarget(detail, order, index, version)
	}

	const extras = detail.kind === "other" ? [] : detail.extras
	const extra = stableUuid === null ? undefined : extras.find(entry => entry.part.file.stableUuid === stableUuid)

	if (extra !== undefined) {
		return {
			key: extra.part.file.stableUuid,
			titleId: title.id,
			titleName: extra.title.length > 0 ? extra.title : title.title,
			kind: "extra",
			itemClass: "extra",
			versionLabel: "",
			parts: [extra.part],
			episode: null,
			seriesDirUuid: null
		}
	}

	if (detail.kind === "other") {
		return detail.part === null
			? null
			: {
					key: detail.part.file.stableUuid,
					titleId: title.id,
					titleName: title.title,
					kind: "other",
					itemClass: "other",
					versionLabel: "",
					parts: [detail.part],
					episode: null,
					seriesDirUuid: null
				}
	}

	const version = pickVersion(detail.versions, stableUuid ?? title.primary?.stableUuid ?? null, null)
	const key = version === null ? null : firstStable(version)

	return version === null || key === null
		? null
		: {
				key,
				titleId: title.id,
				titleName: title.title,
				kind: "movie",
				itemClass: "movie",
				versionLabel: version.label,
				parts: version.parts,
				episode: null,
				seriesDirUuid: null
			}
}

// The episode after (1) or before (-1) `current` in play order, in the same version where the next has one.
// Specials step only among specials; the regular seasons never run into them.
export function adjacentEpisode(detail: VideoTitleDetail, current: WatchItemTarget, direction: 1 | -1): WatchItemTarget | null {
	if (detail.kind !== "series" || current.episode === null) {
		return null
	}

	const order = episodeOrder(detail)
	const from = order.findIndex(episode => episode.versions.some(version => hasFile(version, current.key)))

	if (from === -1) {
		return null
	}

	const special = current.episode.season === 0
	const index = from + direction
	const next = order[index]

	if (next === undefined || (next.season === 0) !== special) {
		return null
	}

	const version = pickVersion(next.versions, null, current.versionLabel)

	return version === null ? null : episodeTarget(detail, order, index, version)
}

// Every episode as a target, for the episodes panel.
export function allEpisodes(detail: VideoTitleDetail, versionLabel: string | null): WatchItemTarget[] {
	if (detail.kind !== "series") {
		return []
	}

	const order = episodeOrder(detail)
	const out: WatchItemTarget[] = []

	order.forEach((episode, index) => {
		const version = pickVersion(episode.versions, null, versionLabel)
		const target = version === null ? null : episodeTarget(detail, order, index, version)

		if (target !== null) {
			out.push(target)
		}
	})

	return out
}

// "S1:E3", "S1:E3-E4", "E12" (absolute numbering), "Special 2".
export function episodeCode(episode: WatchEpisodeInfo): string {
	if (episode.episode <= 0) {
		return ""
	}

	const range =
		episode.episodeEnd > episode.episode ? `E${String(episode.episode)}–E${String(episode.episodeEnd)}` : `E${String(episode.episode)}`

	return episode.season > 0 ? `S${String(episode.season)}:${range}` : range
}

// ── Stacks ──────────────────────────────────────────────────────────────────
// A stacked item plays part after part; its progress is kept as (part, position). With every part's
// duration known it is reported over the whole item, so completion and Continue Watching see one video.
// Without them it is reported per part, and a part before the last reports no duration: it can't complete
// the item. The recorded duration says which form a record has.

export interface StackProgress {
	pos: number
	dur: number
	part: number
}

// A cumulative record's duration matches the whole item's within this.
const STACK_TOTAL_TOLERANCE_SECONDS = 2

function knownDurations(durations: readonly (number | null)[]): number[] | null {
	const out: number[] = []

	for (const duration of durations) {
		if (duration === null || !(duration > 0)) {
			return null
		}

		out.push(duration)
	}

	return out
}

// Position within `part` to the item's progress.
export function stackProgress(durations: readonly (number | null)[], part: number, pos: number, partDuration: number): StackProgress {
	if (durations.length <= 1) {
		return { pos, dur: partDuration, part }
	}

	const known = knownDurations(durations)

	if (known === null) {
		return { pos, dur: part === durations.length - 1 ? partDuration : 0, part }
	}

	let before = 0

	for (let i = 0; i < part && i < known.length; i++) {
		before += known[i] ?? 0
	}

	return { pos: before + pos, dur: known.reduce((sum, value) => sum + value, 0), part }
}

// Whether a stored record is over the whole item (else it is per part).
export function isCumulativeStackProgress(durations: readonly (number | null)[], dur: number): boolean {
	const known = knownDurations(durations)

	return (
		known !== null && known.length > 1 && Math.abs(known.reduce((sum, value) => sum + value, 0) - dur) <= STACK_TOTAL_TOLERANCE_SECONDS
	)
}

// A stored progress back to (part, position within it).
export function locateInStack(
	durations: readonly (number | null)[],
	progress: { pos: number; dur: number; part?: number | undefined }
): { part: number; pos: number } {
	const parts = durations.length
	const recordedPart = progress.part !== undefined && progress.part >= 0 && progress.part < parts ? progress.part : 0

	if (parts <= 1) {
		return { part: 0, pos: progress.pos }
	}

	const known = knownDurations(durations)

	if (known === null || !isCumulativeStackProgress(durations, progress.dur)) {
		const length = known?.[recordedPart]

		return { part: recordedPart, pos: length === undefined ? progress.pos : Math.max(0, Math.min(length, progress.pos)) }
	}

	let before = 0

	for (let i = 0; i < known.length; i++) {
		const length = known[i] ?? 0

		if (progress.pos < before + length || i === known.length - 1) {
			return { part: i, pos: Math.max(0, Math.min(length, progress.pos - before)) }
		}

		before += length
	}

	return { part: recordedPart, pos: 0 }
}

// Post-play's "Next": the first unwatched episode of the same season after the one that ended, else before it.
export function nextUnwatchedInSeason(
	episodes: readonly WatchItemTarget[],
	current: WatchItemTarget,
	played: (target: WatchItemTarget) => boolean
): WatchItemTarget | null {
	const season = current.episode?.season ?? null

	if (season === null) {
		return null
	}

	const same = episodes.filter(episode => episode.episode?.season === season && episode.key !== current.key && !played(episode))
	const order = current.episode?.order ?? 0

	return same.find(episode => (episode.episode?.order ?? 0) > order) ?? same[0] ?? null
}
