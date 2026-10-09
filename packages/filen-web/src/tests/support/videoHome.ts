import { HLC_SPAN, itemKey, type ItemRec } from "@filen/shared/videos"
import type { VideoHomeItem, VideoTitleView } from "@/workers/videoIndex/view.logic"
import type { HomeData } from "@/features/videos/lib/homeRows.logic"
import { homeEntries } from "@/features/videos/lib/homeRows.logic"
import { buildTitleWatch, buildVideoHomeIndex } from "@/features/videos/lib/titleWatch.logic"

// Library fixtures for the Videos home, grid and search tests: titles, their videos and watch-state records.

export function title(id: string, patch: Partial<VideoTitleView> = {}): VideoTitleView {
	return {
		id,
		kind: "movie",
		title: id,
		sortName: id,
		year: 0,
		mergeKey: "",
		roots: ["root"],
		dateAdded: 0,
		poster: null,
		backdrop: null,
		logo: null,
		primary: { stableUuid: `${id}-v`, uuid: `${id}-u`, name: `${id}.mkv`, size: 100 },
		container: "mkv",
		height: 0,
		unsupported: false,
		disc: false,
		versionCount: 1,
		extraCount: 0,
		seasonCount: 0,
		episodeCount: 0,
		...patch
	}
}

export function movieItem(titleId: string, patch: Partial<VideoHomeItem> = {}): VideoHomeItem {
	return {
		titleId,
		stableUuid: `${titleId}-v`,
		uuid: `${titleId}-u`,
		size: 100,
		itemClass: "movie",
		season: 0,
		episode: 0,
		episodeEnd: 0,
		episodeTitle: "",
		group: titleId,
		label: "",
		dateAdded: 0,
		unsupported: false,
		folder: titleId,
		...patch
	}
}

export function episodeItem(seriesId: string, season: number, episode: number, patch: Partial<VideoHomeItem> = {}): VideoHomeItem {
	const id = `${seriesId}-s${String(season)}e${String(episode)}`

	return {
		titleId: seriesId,
		stableUuid: `${id}-v`,
		uuid: `${id}-u`,
		size: 100,
		itemClass: "episode",
		season,
		episode,
		episodeEnd: episode,
		episodeTitle: "",
		group: `${seriesId}/${String(season)}/${String(episode)}`,
		label: "",
		dateAdded: 0,
		unsupported: false,
		folder: `Season ${String(season)}`,
		...patch
	}
}

let clock = 1

// A progress register at `pos` of `dur` seconds, played at `at` (ms), for the file of `item`.
export function progress(item: VideoHomeItem, { pos, dur, at }: { pos: number; dur: number; at: number }): [string, ItemRec] {
	clock++

	return [
		itemKey(item.stableUuid),
		{
			p: { t: at * HLC_SPAN + clock, d: "dev", pos, dur, lp: at, u: item.uuid, sz: item.size }
		}
	]
}

// A watched register played at `at` (ms).
export function watched(item: VideoHomeItem, at: number): [string, ItemRec] {
	clock++

	return [itemKey(item.stableUuid), { w: { t: at * HLC_SPAN + clock, d: "dev", pl: true, pc: 1 } }]
}

export function homeData({
	titles,
	items,
	records = [],
	roots = ["root"],
	now = Date.UTC(2026, 9, 1)
}: {
	titles: VideoTitleView[]
	items: VideoHomeItem[]
	records?: [string, ItemRec][]
	roots?: string[]
	now?: number
}): HomeData {
	const index = buildVideoHomeIndex({
		items,
		watchItems: Object.fromEntries(records),
		lastPlayedCeilingMs: undefined,
		durationOf: () => null
	})

	return {
		titles,
		byId: new Map(titles.map(item => [item.id, item])),
		items,
		index,
		entries: homeEntries(index),
		watch: buildTitleWatch(index),
		roots,
		now
	}
}
