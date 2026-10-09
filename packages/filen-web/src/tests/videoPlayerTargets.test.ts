import { describe, expect, it } from "vitest"
import type { ItemRec } from "@filen/shared/videos"
import { planStart } from "@/features/videos/lib/player/resume.logic"
import {
	adjacentEpisode,
	allEpisodes,
	episodeCode,
	locateInStack,
	nextUnwatchedInSeason,
	resolveWatchTarget,
	stackProgress
} from "@/features/videos/lib/player/watchTarget.logic"
import type { VideoEpisodeView, VideoPartView, VideoTitleDetail, VideoTitleView } from "@/workers/videoIndex/view.logic"

function part(stableUuid: string): VideoPartView {
	return {
		file: { stableUuid, uuid: `${stableUuid}-u`, name: `${stableUuid}.mkv`, size: 1000 },
		container: "mkv",
		unsupported: false,
		height: 1080,
		subtitles: [],
		nfo: null,
		thumb: null,
		edl: null
	}
}

function title(id: string, kind: VideoTitleView["kind"], primary: string | null): VideoTitleView {
	return {
		id,
		kind,
		title: "Show",
		sortName: "show",
		year: 0,
		mergeKey: "",
		roots: [],
		dateAdded: 0,
		poster: null,
		backdrop: null,
		logo: null,
		primary: primary === null ? null : { stableUuid: primary, uuid: `${primary}-u`, name: "", size: 1000 },
		container: "mkv",
		height: 1080,
		unsupported: false,
		disc: false,
		versionCount: 1,
		extraCount: 0,
		seasonCount: 2,
		episodeCount: 4
	}
}

function episode(season: number, number: number, versions: Record<string, string[]>): VideoEpisodeView {
	return {
		season,
		episode: number,
		episodeEnd: number,
		date: 0,
		title: `Episode ${String(number)}`,
		dateAdded: 0,
		versions: Object.entries(versions).map(([label, parts]) => ({ label, parts: parts.map(part) }))
	}
}

const SERIES: VideoTitleDetail = {
	kind: "series",
	title: title("series-dir", "series", "s1e1"),
	seasons: [
		{ number: 0, dirUuid: "", poster: null, backdrop: null, nfo: null, episodes: [episode(0, 1, { "": ["sp1"] })] },
		{
			number: 2,
			dirUuid: "s2",
			poster: null,
			backdrop: null,
			nfo: null,
			episodes: [episode(2, 1, { "": ["s2e1"] })]
		},
		{
			number: 1,
			dirUuid: "s1",
			poster: null,
			backdrop: null,
			nfo: null,
			episodes: [
				episode(1, 2, { "1080p": ["s1e2"], "720p": ["s1e2-720"] }),
				episode(1, 1, { "1080p": ["s1e1"], "720p": ["s1e1-720"] })
			]
		}
	],
	extras: [],
	nfo: null
}

describe("watch targets", () => {
	it("plays the title's first episode without a file, and the named one with it", () => {
		expect(resolveWatchTarget(SERIES, null)?.key).toBe("s1e1")
		expect(resolveWatchTarget(SERIES, "s1e2-720")).toMatchObject({ key: "s1e2-720", versionLabel: "720p", seriesDirUuid: "series-dir" })
	})

	it("advances in season order, specials last and on their own, keeping the version", () => {
		const first = resolveWatchTarget(SERIES, "s1e1-720")

		expect(first).not.toBeNull()

		if (first === null) {
			return
		}

		const second = adjacentEpisode(SERIES, first, 1)

		expect(second?.key).toBe("s1e2-720")

		const third = second === null ? null : adjacentEpisode(SERIES, second, 1)

		expect(third?.key).toBe("s2e1")
		expect(third === null ? null : adjacentEpisode(SERIES, third, 1)).toBeNull()
		expect(adjacentEpisode(SERIES, first, -1)).toBeNull()
		expect(allEpisodes(SERIES, null).map(target => target.key)).toEqual(["s1e1", "s1e2", "s2e1", "sp1"])
	})

	it("offers the first unwatched episode of the season after one that ended", () => {
		const episodes = allEpisodes(SERIES, null)
		const last = episodes[1]

		expect(last).toBeDefined()

		if (last === undefined) {
			return
		}

		expect(nextUnwatchedInSeason(episodes, last, () => false)?.key).toBe("s1e1")
		expect(nextUnwatchedInSeason(episodes, last, target => target.key === "s1e1")).toBeNull()
	})

	it("resolves movies by version and extras by file", () => {
		const movie: VideoTitleDetail = {
			kind: "movie",
			title: title("m", "movie", "cd1"),
			versions: [
				{ label: "", parts: [part("cd1"), part("cd2")] },
				{ label: "4K", parts: [part("uhd")] }
			],
			extras: [{ type: "trailer", title: "Trailer", part: part("trailer") }],
			nfo: null
		}

		expect(resolveWatchTarget(movie, null)).toMatchObject({ key: "cd1", kind: "movie" })
		expect(resolveWatchTarget(movie, "cd2")?.parts.map(entry => entry.file.stableUuid)).toEqual(["cd1", "cd2"])
		expect(resolveWatchTarget(movie, "uhd")?.versionLabel).toBe("4K")
		expect(resolveWatchTarget(movie, "trailer")).toMatchObject({ key: "trailer", kind: "extra", itemClass: "extra" })
	})

	it("codes episodes", () => {
		expect(episodeCode({ season: 1, episode: 3, episodeEnd: 3, title: "", order: 0 })).toBe("S1:E3")
		expect(episodeCode({ season: 2, episode: 5, episodeEnd: 6, title: "", order: 0 })).toBe("S2:E5–E6")
		expect(episodeCode({ season: 0, episode: 12, episodeEnd: 12, title: "", order: 0 })).toBe("E12")
		expect(episodeCode({ season: 1, episode: 0, episodeEnd: 0, title: "", order: 0 })).toBe("")
	})
})

describe("stacks", () => {
	it("reports progress over the whole item once every part's duration is known", () => {
		expect(stackProgress([600, 900], 1, 100, 900)).toEqual({ pos: 700, dur: 1500, part: 1 })
		expect(stackProgress([null], 0, 100, 900)).toEqual({ pos: 100, dur: 900, part: 0 })
	})

	it("reports per part without the durations, and no duration (no completion) before the last part", () => {
		expect(stackProgress([600, null], 0, 590, 600)).toEqual({ pos: 590, dur: 0, part: 0 })
		expect(stackProgress([600, null], 1, 100, 900)).toEqual({ pos: 100, dur: 900, part: 1 })
	})

	it("maps stored progress back to a part, telling cumulative from per-part records by their duration", () => {
		expect(locateInStack([600, 900], { pos: 700, dur: 1500, part: 1 })).toEqual({ part: 1, pos: 100 })
		expect(locateInStack([600, 900], { pos: 300, dur: 1500 })).toEqual({ part: 0, pos: 300 })
		// Recorded per part while the durations were unknown, read back once they are known.
		expect(locateInStack([600, 900], { pos: 100, dur: 900, part: 1 })).toEqual({ part: 1, pos: 100 })
		expect(locateInStack([600, 900], { pos: 400, dur: 0, part: 0 })).toEqual({ part: 0, pos: 400 })
		expect(locateInStack([600, null], { pos: 100, dur: 900, part: 1 })).toEqual({ part: 1, pos: 100 })
		expect(locateInStack([null], { pos: 42, dur: 0 })).toEqual({ part: 0, pos: 42 })
	})
})

describe("resume", () => {
	const file = { uuid: "f-u", size: 1000 }

	function rec(pos: number, dur: number, part?: number): ItemRec {
		return { p: { t: 1000, d: "dev", pos, dur, lp: 1, u: "f-u", sz: 1000, ...(part === undefined ? {} : { part }) } }
	}

	it("resumes at a valid position and says so", () => {
		expect(planStart({ request: "resume", rec: rec(1200, 3000), file, partDurations: [3000] })).toEqual({
			part: 0,
			position: 1200,
			resumed: true
		})
	})

	it("starts over when asked, without progress, or outside the resumable window", () => {
		expect(planStart({ request: "start", rec: rec(1200, 3000), file, partDurations: [3000] }).position).toBe(0)
		expect(planStart({ request: "resume", rec: undefined, file, partDurations: [3000] }).resumed).toBe(false)
		expect(planStart({ request: "resume", rec: rec(60, 3000), file, partDurations: [3000] }).resumed).toBe(false)
		expect(planStart({ request: "resume", rec: rec(2950, 3000), file, partDurations: [3000] }).resumed).toBe(false)
		expect(planStart({ request: "resume", rec: rec(100, 200), file, partDurations: [200] }).resumed).toBe(false)
	})

	it("ignores progress recorded against a different file", () => {
		expect(
			planStart({ request: "resume", rec: rec(1200, 3000), file: { uuid: "other", size: 5 }, partDurations: [3000] }).resumed
		).toBe(false)
	})

	it("resumes a stack in the right part, from cumulative and per-part records", () => {
		expect(planStart({ request: "resume", rec: rec(2000, 3000, 1), file, partDurations: [1500, 1500] })).toEqual({
			part: 1,
			position: 500,
			resumed: true
		})
		// Per part, before the last part (no duration): still resumable.
		expect(planStart({ request: "resume", rec: rec(400, 0, 0), file, partDurations: [1500, null] })).toEqual({
			part: 0,
			position: 400,
			resumed: true
		})
		expect(planStart({ request: "resume", rec: rec(10, 0, 0), file, partDurations: [1500, null] }).resumed).toBe(false)
		// Past 90 % of the whole item: watched, not resumed.
		expect(planStart({ request: "resume", rec: rec(2900, 3000, 1), file, partDurations: [1500, 1500] }).resumed).toBe(false)
	})
})
