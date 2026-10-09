import { buildHomeIndex, selectUnwatchedCounts, type HomeEntry, type HomeIndex, type HomeVideo, type ItemRec } from "@filen/shared/videos"
import type { VideoHomeItem } from "@/workers/videoIndex/view.logic"

// What a title's card and the watch filters show of its watch state, derived from the library's videos and the
// watch-state document already held (no request).

export function toHomeVideos(items: readonly VideoHomeItem[], durationOf: (uuid: string) => number | null): HomeVideo[] {
	return items.map(item => ({
		id: item.stableUuid,
		uuid: item.uuid,
		size: item.size,
		itemClass: item.itemClass,
		seriesId: item.itemClass === "episode" ? item.titleId : null,
		season: item.itemClass === "episode" ? item.season : null,
		episode: item.itemClass === "episode" && item.episode > 0 ? item.episode : null,
		episodeEnd: item.itemClass === "episode" && item.episode > 0 ? item.episodeEnd : null,
		versionGroup: item.group,
		versionLabel: item.label,
		dateAdded: item.dateAdded,
		durationSec: durationOf(item.uuid)
	}))
}

export function buildVideoHomeIndex({
	items,
	watchItems,
	lastPlayedCeilingMs,
	durationOf
}: {
	items: readonly VideoHomeItem[]
	watchItems: Readonly<Record<string, ItemRec>>
	lastPlayedCeilingMs: number | undefined
	durationOf: (uuid: string) => number | null
}): HomeIndex {
	return buildHomeIndex({
		videos: toHomeVideos(items, durationOf),
		items: watchItems,
		lastPlayedCeilingMs
	})
}

export type WatchCategory = "unwatched" | "inProgress" | "watched"

export interface TitleWatch {
	category: WatchCategory
	// 0-1 when the title has a resume point (a movie's or other video's latest played version).
	progress: number | null
	// Seconds left of that resume point, when its duration is known.
	remainingSec: number | null
	// Series: unplayed episodes outside the specials.
	unwatched: number | null
	lastPlayedMs: number
}

// The title an entry belongs to: an episode's series, else its version group (the movie or other video).
export function entryTitleId(entry: HomeEntry): string {
	return entry.video.itemClass === "episode" ? (entry.video.seriesId ?? entry.group) : entry.group
}

export function entryProgress(entry: HomeEntry): { progress: number | null; remainingSec: number | null } {
	const { positionSec, durationSec, inProgress } = entry.state

	if (!inProgress || !(durationSec > 0)) {
		return { progress: null, remainingSec: null }
	}

	return {
		progress: Math.min(1, Math.max(0, positionSec / durationSec)),
		remainingSec: Math.max(0, durationSec - positionSec)
	}
}

// Every title's watch summary in one pass over the index.
export function buildTitleWatch(index: HomeIndex): Map<string, TitleWatch> {
	const out = new Map<string, TitleWatch>()
	const latest = new Map<string, HomeEntry>()

	for (const entry of index.entries) {
		if (entry.video.itemClass === "episode") {
			continue
		}

		const id = entryTitleId(entry)
		const current = out.get(id)
		const played = current?.category === "watched" || entry.state.played

		out.set(id, {
			category: played ? "watched" : "unwatched",
			progress: null,
			remainingSec: null,
			unwatched: null,
			lastPlayedMs: Math.max(current?.lastPlayedMs ?? 0, entry.state.lastPlayedMs)
		})

		const previous = latest.get(id)

		if (entry.state.lastPlayedMs > 0 && (previous === undefined || entry.state.lastPlayedMs > previous.state.lastPlayedMs)) {
			latest.set(id, entry)
		}
	}

	// The most recently played version decides the resume point, as Continue Watching does.
	for (const [id, entry] of latest) {
		const watch = out.get(id)

		if (watch === undefined || !entry.state.inProgress) {
			continue
		}

		const { progress, remainingSec } = entryProgress(entry)

		watch.category = "inProgress"
		watch.progress = progress
		watch.remainingSec = remainingSec
	}

	const unwatched = selectUnwatchedCounts(index)

	for (const series of index.series.values()) {
		const total = unwatched.get(series.id)?.total ?? 0
		const regular = series.units.filter(unit => unit.season !== 0)
		const started = series.units.some(unit => unit.played || unit.resume !== null)

		out.set(series.id, {
			category: regular.length > 0 && total === 0 ? "watched" : started ? "inProgress" : "unwatched",
			progress: null,
			remainingSec: null,
			unwatched: total,
			lastPlayedMs: series.lastPlayedMs
		})
	}

	return out
}

const UNWATCHED: TitleWatch = {
	category: "unwatched",
	progress: null,
	remainingSec: null,
	unwatched: null,
	lastPlayedMs: 0
}

export function titleWatchOf(watch: ReadonlyMap<string, TitleWatch>, titleId: string): TitleWatch {
	return watch.get(titleId) ?? UNWATCHED
}

export type CardBadge = { kind: "watched" } | { kind: "unwatched"; count: number } | null

// A series shows its unplayed episode count until every one is played; anything else a check once played.
export function cardBadge(kind: "movie" | "series" | "other", watch: TitleWatch): CardBadge {
	if (kind === "series" && watch.unwatched !== null && watch.unwatched > 0) {
		return { kind: "unwatched", count: watch.unwatched }
	}

	return watch.category === "watched" ? { kind: "watched" } : null
}
