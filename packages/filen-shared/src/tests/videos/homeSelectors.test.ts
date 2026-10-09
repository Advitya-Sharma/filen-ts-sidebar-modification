import { describe, expect, it } from "vitest"
import {
	buildContinueWatchingRow,
	buildHomeIndex,
	hideFromContinueWatching,
	itemKey,
	LATEST_GROUP_WINDOW_MS,
	markWatched,
	selectContinueWatching,
	selectLatestEpisodes,
	selectLatestMovies,
	selectNextUp,
	selectSeriesPrimaryAction,
	selectUnwatchedCounts,
	withCompletion,
	withProgress,
	type HomeIndex,
	type HomeVideo,
	type ItemRec
} from "@filen/shared/videos"

const NOW = Date.UTC(2026, 9, 9)
const DAY = 86_400_000
const HOUR = 3_600_000

function movie(id: string, over: Partial<HomeVideo> = {}): HomeVideo {
	return {
		id,
		uuid: `${id}-file`,
		size: 1000,
		itemClass: "movie",
		dateAdded: NOW - DAY,
		durationSec: 3600,
		...over
	}
}

function episode(series: string, season: number | null, number: number | null, over: Partial<HomeVideo> = {}): HomeVideo {
	return {
		id: `${series}-s${season}e${number}`,
		uuid: `${series}-s${season}e${number}-file`,
		size: 1000,
		itemClass: "episode",
		seriesId: series,
		season,
		episode: number,
		dateAdded: NOW - 10 * DAY,
		durationSec: 1800,
		...over
	}
}

let counter = 0

// Stamped at wall-clock `at`, as a device writing then would.
function nextStamp(at = NOW): { t: number; d: string } {
	counter = (counter + 1) % 1000

	return {
		t: at * 1000 + counter,
		d: "dev"
	}
}

// Watch-state writes against a video at a given wall-clock time.
function played(video: HomeVideo, at: number): ItemRec {
	return withCompletion({
		rec: withProgress({
			rec: undefined,
			stamp: nextStamp(at),
			pos: (video.durationSec ?? 0) - 10,
			dur: video.durationSec ?? 0,
			lp: at,
			uuid: video.uuid,
			size: video.size
		}),
		stamp: nextStamp(at)
	})
}

function inProgress(video: HomeVideo, at: number, fraction = 0.5): ItemRec {
	return withProgress({
		rec: undefined,
		stamp: nextStamp(at),
		pos: (video.durationSec ?? 0) * fraction,
		dur: video.durationSec ?? 0,
		lp: at,
		uuid: video.uuid,
		size: video.size
	})
}

function index(videos: HomeVideo[], records: [HomeVideo, ItemRec][] = []): HomeIndex {
	const items: Record<string, ItemRec> = {}

	for (const [video, rec] of records) {
		items[itemKey(video.id)] = rec
	}

	return buildHomeIndex({
		videos,
		items
	})
}

function ids(entries: readonly { video: HomeVideo }[]): string[] {
	return entries.map(entry => entry.video.id)
}

describe("buildHomeIndex", () => {
	it("orders episode units by season and episode with unnumbered ones last and groups versions", () => {
		const a = episode("show", 2, 1)
		const b = episode("show", 1, 2)
		const c = episode("show", 1, null, {
			id: "loose"
		})
		const d1 = episode("show", 1, 1, {
			id: "e1-720",
			versionGroup: "e1"
		})
		const d2 = episode("show", 1, 1, {
			id: "e1-1080",
			versionGroup: "e1"
		})
		const special = episode("show", 0, 1)
		const built = index([a, b, c, d1, d2, special, movie("m")])
		const series = built.series.get("show")

		expect(built.entries).toHaveLength(7)
		expect(series?.units.map(unit => unit.group)).toEqual(["show-s0e1", "e1", "show-s1e2", "loose", "show-s2e1"])
		expect(series?.units[1]?.versions.map(entry => entry.video.id)).toEqual(["e1-720", "e1-1080"])
		expect(series?.seasons).toEqual(
			new Map([
				[2, 1],
				[1, 3],
				[0, 1]
			])
		)
	})
})

describe("selectContinueWatching", () => {
	it("lists resumable items by last play and collapses versions", () => {
		const m1 = movie("m1")
		const m2a = movie("m2a", {
			versionGroup: "m2"
		})
		const m2b = movie("m2b", {
			versionGroup: "m2"
		})
		const m3 = movie("m3")
		const m4 = movie("m4")
		const built = index(
			[m1, m2a, m2b, m3, m4],
			[
				[m1, inProgress(m1, NOW - 3 * HOUR)],
				[m2a, inProgress(m2a, NOW - 5 * HOUR)],
				[m2b, inProgress(m2b, NOW - HOUR)],
				[m3, inProgress(m3, NOW, 0.95)],
				[
					m4,
					hideFromContinueWatching({
						rec: inProgress(m4, NOW),
						stamp: nextStamp()
					})
				]
			]
		)

		expect(
			ids(
				selectContinueWatching({
					index: built
				})
			)
		).toEqual(["m2b", "m1"])
		expect(
			ids(
				selectContinueWatching({
					index: built,
					limit: 1
				})
			)
		).toEqual(["m2b"])
	})
})

describe("selectContinueWatching version groups", () => {
	it("follows the most recently played version, dropping the group when that one is finished", () => {
		const sd = movie("sd", {
			versionGroup: "film"
		})
		const hd = movie("hd", {
			versionGroup: "film"
		})
		const finishedLater = index(
			[sd, hd],
			[
				[sd, inProgress(sd, NOW - DAY)],
				[hd, played(hd, NOW)]
			]
		)

		expect(
			selectContinueWatching({
				index: finishedLater
			})
		).toEqual([])

		const resumedLater = index(
			[sd, hd],
			[
				[sd, played(sd, NOW - DAY)],
				[hd, inProgress(hd, NOW)]
			]
		)

		expect(
			ids(
				selectContinueWatching({
					index: resumedLater
				})
			)
		).toEqual(["hd"])
	})

	it("does not resume an older version of an episode once a newer one is finished", () => {
		const e1sd = episode("show", 1, 1, {
			id: "e1-sd",
			versionGroup: "e1"
		})
		const e1hd = episode("show", 1, 1, {
			id: "e1-hd",
			versionGroup: "e1"
		})
		const e2 = episode("show", 1, 2)
		const built = index(
			[e1sd, e1hd, e2],
			[
				[e1sd, inProgress(e1sd, NOW - DAY)],
				[e1hd, played(e1hd, NOW)]
			]
		)

		expect(built.series.get("show")?.units[0]?.resume).toBeNull()
		expect(
			selectSeriesPrimaryAction({
				index: built,
				seriesId: "show"
			})
		).toMatchObject({
			type: "play",
			entry: {
				video: {
					id: e2.id
				}
			}
		})
	})
})

describe("selectNextUp", () => {
	const show = [episode("show", 1, 1), episode("show", 1, 2), episode("show", 1, 3), episode("show", 2, 1)]
	const [e1, e2, e3, s2e1] = show as [HomeVideo, HomeVideo, HomeVideo, HomeVideo]

	it("offers the episode after the highest played one", () => {
		const built = index(show, [
			[e1, played(e1, NOW - DAY)],
			[e2, played(e2, NOW - 2 * DAY)]
		])
		const next = selectNextUp({
			index: built,
			now: NOW
		})

		expect(next).toHaveLength(1)
		expect(next[0]?.entry.video.id).toBe(e3.id)
		expect(next[0]?.lastPlayedMs).toBe(NOW - DAY)
	})

	it("moves on past gaps and into the next season", () => {
		const built = index(show, [
			[e1, played(e1, NOW)],
			[e3, played(e3, NOW)]
		])

		expect(
			selectNextUp({
				index: built,
				now: NOW
			})[0]?.entry.video.id
		).toBe(s2e1.id)
	})

	it("leaves a resumable next episode to Continue Watching", () => {
		const built = index(show, [
			[e1, played(e1, NOW - DAY)],
			[e2, inProgress(e2, NOW)]
		])

		expect(
			selectNextUp({
				index: built,
				now: NOW
			})
		).toEqual([])
		expect(
			ids(
				selectContinueWatching({
					index: built
				})
			)
		).toEqual([e2.id])
	})

	it("keeps a hidden in-progress next episode out of Next Up and the merged row until it is played again", () => {
		const hidden = hideFromContinueWatching({
			rec: inProgress(e2, NOW),
			stamp: nextStamp(NOW + 1)
		})
		const built = index(show, [
			[e1, played(e1, NOW - DAY)],
			[e2, hidden]
		])
		const resume = selectContinueWatching({
			index: built
		})
		const nextUp = selectNextUp({
			index: built,
			now: NOW
		})

		expect(resume).toEqual([])
		expect(nextUp).toEqual([])
		expect(
			buildContinueWatchingRow({
				resume,
				nextUp
			})
		).toEqual([])
		// The series page still offers to resume it
		expect(
			selectSeriesPrimaryAction({
				index: built,
				seriesId: "show"
			})
		).toMatchObject({
			type: "resume",
			entry: {
				video: {
					id: e2.id
				}
			}
		})

		const replayed = index(show, [
			[e1, played(e1, NOW - DAY)],
			[
				e2,
				withProgress({
					rec: hidden,
					stamp: nextStamp(NOW + 2),
					pos: 1000,
					dur: 1800,
					lp: NOW + 2,
					uuid: e2.uuid,
					size: e2.size
				})
			]
		])

		expect(
			ids(
				selectContinueWatching({
					index: replayed
				})
			)
		).toEqual([e2.id])
	})

	it("skips series with nothing played, everything played, or only specials played", () => {
		const special = episode("show", 0, 1)
		const nothing = index(show, [[e2, inProgress(e2, NOW)]])
		const all = index(
			show,
			show.map(video => [video, played(video, NOW)])
		)
		const specials = index([...show, special], [[special, played(special, NOW)]])

		for (const built of [nothing, all, specials]) {
			expect(
				selectNextUp({
					index: built,
					now: NOW
				})
			).toEqual([])
		}
	})

	it("advances past the end of a multi-episode file", () => {
		const double = episode("show", 1, 1, {
			episodeEnd: 2
		})
		const videos = [double, episode("show", 1, 2, { id: "dup-e2" }), episode("show", 1, 3)]
		const built = index(videos, [[double, played(double, NOW)]])

		expect(
			selectNextUp({
				index: built,
				now: NOW
			})[0]?.entry.video.id
		).toBe("show-s1e3")
	})

	it("supports absolute numbering", () => {
		const videos = [1, 2, 3].map(number =>
			episode("anime", null, number, {
				absolute: true,
				season: 3
			})
		)
		const built = index(videos, [[videos[0] as HomeVideo, played(videos[0] as HomeVideo, NOW)]])

		expect(built.series.get("anime")?.units.map(unit => unit.season)).toEqual([null, null, null])
		expect(
			selectNextUp({
				index: built,
				now: NOW
			})[0]?.entry.video.id
		).toBe("anime-snulle2")
	})

	it("keeps the version label played last", () => {
		const a1 = episode("show", 1, 1, {
			id: "a1-sd",
			versionGroup: "a1",
			versionLabel: "SD"
		})
		const a2sd = episode("show", 1, 2, {
			id: "a2-sd",
			versionGroup: "a2",
			versionLabel: "SD"
		})
		const a2hd = episode("show", 1, 2, {
			id: "a2-hd",
			versionGroup: "a2",
			versionLabel: "HD"
		})
		const built = index([a1, a2hd, a2sd], [[a1, played(a1, NOW)]])

		expect(
			selectNextUp({
				index: built,
				now: NOW
			})[0]?.entry.video.id
		).toBe("a2-sd")
	})

	it("applies the cutoff, with null meaning unlimited, and orders series by last play", () => {
		const other = [episode("other", 1, 1), episode("other", 1, 2)]
		const old = other[0] as HomeVideo
		const built = index(
			[...show, ...other],
			[
				[e1, played(e1, NOW - 10 * DAY)],
				[old, played(old, NOW - 400 * DAY)]
			]
		)

		expect(
			selectNextUp({
				index: built,
				now: NOW
			}).map(entry => entry.seriesId)
		).toEqual(["show"])
		expect(
			selectNextUp({
				index: built,
				now: NOW,
				cutoffDays: 5
			})
		).toEqual([])
		expect(
			selectNextUp({
				index: built,
				now: NOW,
				cutoffDays: null
			}).map(entry => entry.seriesId)
		).toEqual(["show", "other"])
	})
})

describe("buildContinueWatchingRow", () => {
	it("interleaves both sources by last play with one card per series", () => {
		const show = [episode("show", 1, 1), episode("show", 1, 2), episode("show", 1, 5), episode("show", 1, 6)]
		const [e1, e2, e5] = show as [HomeVideo, HomeVideo, HomeVideo]
		const other = [episode("other", 1, 1), episode("other", 1, 2)]
		const o1 = other[0] as HomeVideo
		const m = movie("m")
		const built = index(
			[...show, ...other, m],
			[
				[e1, played(e1, NOW - 5 * DAY)],
				[e2, inProgress(e2, NOW - 4 * DAY)],
				[e5, played(e5, NOW - DAY)],
				[o1, played(o1, NOW - 2 * DAY)],
				[m, inProgress(m, NOW - 3 * DAY)]
			]
		)
		const resume = selectContinueWatching({
			index: built
		})
		const nextUp = selectNextUp({
			index: built,
			now: NOW
		})
		const row = buildContinueWatchingRow({
			resume,
			nextUp
		})

		expect(ids(resume)).toEqual(["m", e2.id])
		expect(nextUp.map(entry => entry.entry.video.id)).toEqual(["show-s1e6", "other-s1e2"])
		expect(row.map(card => `${card.type}:${card.entry.video.id}`)).toEqual(["next:show-s1e6", "next:other-s1e2", "resume:m"])
		expect(
			buildContinueWatchingRow({
				resume,
				nextUp,
				limit: 2
			})
		).toHaveLength(2)
	})

	it("prefers the resume card on a tie", () => {
		const show = [episode("show", 1, 1), episode("show", 1, 2), episode("show", 1, 3)]
		const [e1, e2] = show as [HomeVideo, HomeVideo]
		const built = index(show, [
			[e2, played(e2, NOW - DAY)],
			[e1, inProgress(e1, NOW - DAY)]
		])
		const row = buildContinueWatchingRow({
			resume: selectContinueWatching({
				index: built
			}),
			nextUp: selectNextUp({
				index: built,
				now: NOW
			})
		})

		expect(row.map(card => `${card.type}:${card.entry.video.id}`)).toEqual([`resume:${e1.id}`])
	})
})

describe("selectLatestMovies", () => {
	it("lists movies newest first, one per version group", () => {
		const videos = [
			movie("old", {
				dateAdded: NOW - 10 * DAY
			}),
			movie("v-sd", {
				versionGroup: "v",
				dateAdded: NOW - 5 * DAY
			}),
			movie("v-hd", {
				versionGroup: "v",
				dateAdded: NOW - DAY
			}),
			movie("new", {
				dateAdded: NOW
			}),
			movie("extra", {
				itemClass: "extra",
				dateAdded: NOW
			}),
			episode("show", 1, 1, {
				dateAdded: NOW
			})
		]
		const built = index(videos)

		expect(
			ids(
				selectLatestMovies({
					index: built
				})
			)
		).toEqual(["new", "v-hd", "old"])
		expect(
			ids(
				selectLatestMovies({
					index: built,
					limit: 2
				})
			)
		).toEqual(["new", "v-hd"])
		expect(
			selectLatestMovies({
				index: built,
				limit: 0
			})
		).toEqual([])
	})
})

describe("selectLatestEpisodes", () => {
	const added = (at: number) => ({
		dateAdded: at
	})

	it("shows one new episode as an episode card", () => {
		const built = index([episode("show", 1, 1, added(NOW - 10 * DAY)), episode("show", 1, 2, added(NOW))])

		expect(
			selectLatestEpisodes({
				index: built
			})
		).toEqual([
			{
				type: "episode",
				seriesId: "show",
				entry: built.entries[1],
				addedMs: NOW
			}
		])
	})

	it("groups several new episodes of one season into a season card, or a series card for a single-season series", () => {
		const multi = index([
			episode("show", 1, 1, added(NOW - 10 * DAY)),
			episode("show", 2, 1, added(NOW - LATEST_GROUP_WINDOW_MS)),
			episode("show", 2, 2, added(NOW)),
			episode("show", 2, 3, added(NOW - LATEST_GROUP_WINDOW_MS - 1))
		])

		expect(
			selectLatestEpisodes({
				index: multi
			})
		).toEqual([
			{
				type: "season",
				seriesId: "show",
				season: 2,
				count: 2,
				addedMs: NOW
			}
		])

		const single = index([
			episode("show", 1, 1, added(NOW - 10 * DAY)),
			episode("show", 1, 2, added(NOW)),
			episode("show", 1, 3, added(NOW))
		])

		expect(
			selectLatestEpisodes({
				index: single
			})[0]?.type
		).toBe("series")
	})

	it("uses a season card when the one new episode is the whole season", () => {
		const built = index([episode("show", 1, 1, added(NOW - 10 * DAY)), episode("show", 2, 1, added(NOW))])

		expect(
			selectLatestEpisodes({
				index: built
			})
		).toEqual([
			{
				type: "season",
				seriesId: "show",
				season: 2,
				count: 1,
				addedMs: NOW
			}
		])
	})

	it("uses a series card when new episodes span seasons", () => {
		const built = index([
			episode("show", 1, 1, added(NOW)),
			episode("show", 2, 1, added(NOW - HOUR)),
			episode("show", 3, 1, added(NOW - 5 * DAY))
		])

		expect(
			selectLatestEpisodes({
				index: built
			})
		).toEqual([
			{
				type: "series",
				seriesId: "show",
				count: 2,
				addedMs: NOW
			}
		])
	})

	it("ranks series by their newest episode and limits the count", () => {
		const built = index([
			episode("a", 1, 1, added(NOW - 3 * DAY)),
			episode("b", 1, 1, added(NOW - DAY)),
			episode("b", 1, 2, added(NOW - 9 * DAY)),
			episode("c", 1, 1, added(NOW - 2 * DAY))
		])

		expect(
			selectLatestEpisodes({
				index: built,
				limit: 2
			}).map(card => card.seriesId)
		).toEqual(["b", "c"])
	})
})

describe("selectSeriesPrimaryAction", () => {
	const show = [episode("show", 0, 1), episode("show", 1, 1), episode("show", 1, 2), episode("show", 1, 3)]
	const [special, e1, e2, e3] = show as [HomeVideo, HomeVideo, HomeVideo, HomeVideo]
	const action = (records: [HomeVideo, ItemRec][]) => {
		const result = selectSeriesPrimaryAction({
			index: index(show, records),
			seriesId: "show"
		})

		return result === null ? null : `${result.type}:${result.entry.video.id}`
	}

	it("plays the first episode when nothing is watched", () => {
		expect(action([])).toBe(`playFirst:${e1.id}`)
		expect(action([[special, played(special, NOW)]])).toBe(`playFirst:${e1.id}`)
	})

	it("resumes the next episode when it is in progress", () => {
		expect(action([[e1, inProgress(e1, NOW)]])).toBe(`resume:${e1.id}`)
		expect(
			action([
				[e1, played(e1, NOW)],
				[e2, inProgress(e2, NOW)]
			])
		).toBe(`resume:${e2.id}`)
	})

	it("plays the next episode after the highest watched one", () => {
		expect(action([[e2, played(e2, NOW)]])).toBe(`play:${e3.id}`)
	})

	it("starts over once everything is watched", () => {
		expect(action([e1, e2, e3].map(video => [video, played(video, NOW)]))).toBe(`playFirst:${e1.id}`)
	})

	it("is null for an unknown series", () => {
		expect(
			selectSeriesPrimaryAction({
				index: index(show),
				seriesId: "missing"
			})
		).toBeNull()
	})
})

describe("selectUnwatchedCounts", () => {
	it("counts unplayed episode slots per season and series, specials only per season", () => {
		const show = [
			episode("show", 0, 1),
			episode("show", 1, 1),
			episode("show", 1, 2, {
				id: "e2-a",
				versionGroup: "e2"
			}),
			episode("show", 1, 2, {
				id: "e2-b",
				versionGroup: "e2"
			}),
			episode("show", 2, 1)
		]
		const [, e1, , e2b] = show as [HomeVideo, HomeVideo, HomeVideo, HomeVideo]
		const counts = selectUnwatchedCounts(
			index(show, [
				[
					e1,
					markWatched({
						rec: undefined,
						stamp: nextStamp()
					})
				],
				[e2b, inProgress(e2b, NOW)]
			])
		)

		expect(counts.get("show")).toEqual({
			total: 2,
			seasons: new Map([
				[0, 1],
				[1, 1],
				[2, 1]
			])
		})
	})
})

describe("performance", () => {
	it("derives every home row for 20k items quickly", () => {
		const videos: HomeVideo[] = []
		const items: Record<string, ItemRec> = {}

		for (let i = 0; i < 4000; i++) {
			const video = movie(`movie-${i}`, {
				versionGroup: `group-${Math.floor(i / 2)}`,
				dateAdded: NOW - i * HOUR
			})

			videos.push(video)

			if (i % 3 === 0) {
				items[itemKey(video.id)] = i % 2 === 0 ? played(video, NOW - i * 60_000) : inProgress(video, NOW - i * 60_000)
			}
		}

		for (let s = 0; s < 400; s++) {
			for (let season = 1; season <= 4; season++) {
				for (let number = 1; number <= 10; number++) {
					const video = episode(`series-${s}`, season, number, {
						dateAdded: NOW - (s * 40 + season * 10 + number) * HOUR
					})

					videos.push(video)

					if (season <= 2 && (s + number) % 4 !== 0) {
						items[itemKey(video.id)] =
							number === 10 ? inProgress(video, NOW - s * HOUR) : played(video, NOW - s * HOUR - number)
					}
				}
			}
		}

		expect(videos).toHaveLength(20_000)

		const started = performance.now()
		const built = buildHomeIndex({
			videos,
			items
		})
		const resume = selectContinueWatching({
			index: built
		})
		const nextUp = selectNextUp({
			index: built,
			now: NOW,
			cutoffDays: null
		})
		const row = buildContinueWatchingRow({
			resume,
			nextUp,
			limit: 40
		})
		const latestMovies = selectLatestMovies({
			index: built,
			limit: 40
		})
		const latestEpisodes = selectLatestEpisodes({
			index: built,
			limit: 40
		})
		const counts = selectUnwatchedCounts(built)
		const elapsed = performance.now() - started

		console.info(`home rows for ${videos.length} items: ${elapsed.toFixed(1)} ms`)
		expect(row).toHaveLength(40)
		expect(latestMovies).toHaveLength(40)
		expect(latestEpisodes).toHaveLength(40)
		expect(counts.size).toBe(400)
		expect(nextUp.length).toBeGreaterThan(0)
		expect(elapsed).toBeLessThan(500)
	})
})
