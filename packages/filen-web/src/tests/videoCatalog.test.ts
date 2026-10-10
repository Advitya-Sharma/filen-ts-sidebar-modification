import { describe, expect, it } from "vitest"
import {
	HLC_SPAN,
	buildContinueWatchingRow,
	itemKey,
	selectContinueWatching,
	selectLatestEpisodes,
	selectLatestMovies,
	selectNextUp,
	type ItemRec
} from "@filen/shared/videos"
import {
	VideoCatalog,
	allItems,
	allTitles,
	arrayItems,
	arrayTitles,
	catalogContent,
	groupByTitle,
	type CatalogContent,
	type GridQuery
} from "@/workers/videoIndex/catalog.logic"
import { CATALOG_HEADER_BYTES, catalogKey, catalogSlotOrder, decodeCatalog, encodeCatalog } from "@/workers/videoIndex/catalogFile.logic"
import type { VideoHomeItem, VideoTitleView } from "@/workers/videoIndex/view.logic"
import {
	HOME_ROOT_ROW_CARDS,
	HOME_ROW_MAX_CARDS,
	buildHomeRows,
	byTitleName,
	itemCard,
	newestFirst,
	titleCard,
	type HomeRowModel
} from "@/features/videos/lib/homeRows.logic"
import { buildTitleWatch, buildVideoHomeIndex, homeEntries, titleWatchOf } from "@/features/videos/lib/titleWatch.logic"
import { DEFAULT_GRID_PREFS, GRID_SORTS, gridView, inGridSource, type GridPrefs } from "@/features/videos/lib/videoGrid.logic"
import { buildVideoSearchIndex, searchVideos } from "@/features/videos/lib/search.logic"
import { episodeItem, movieItem, title } from "@/tests/support/videoHome"

// The library catalog the SDK worker serves the screens from: what it shows must be exactly what one index over the
// whole library would show (the previous main-thread derivation, rebuilt below as the reference), and its persisted
// form must round-trip and reject anything torn or written under other rules.

const NOW = Date.UTC(2026, 9, 1)
const DAY = 86_400_000

function rng(seed: number): () => number {
	let state = seed >>> 0 || 1

	return () => {
		state ^= state << 13
		state ^= state >>> 17
		state ^= state << 5

		return (state >>> 0) / 0x1_0000_0000
	}
}

interface Library {
	titles: VideoTitleView[]
	items: VideoHomeItem[]
	records: Record<string, ItemRec>
}

// A library with ties in names and dates, multi-version movies, series with specials and unnumbered episodes, other
// videos, hidden titles, two roots; and watch records of every kind on random videos.
function randomLibrary(seed: number): Library {
	const random = rng(seed)
	const pick = <T>(values: readonly T[]): T => values[Math.floor(random() * values.length)] as T
	const titles: VideoTitleView[] = []
	const items: VideoHomeItem[] = []
	const names = ["Alien", "Arrival", "Heat", "The Matrix", "Dune", "Ronin", "Up", "Zodiac"]

	for (let i = 0; i < 30; i++) {
		const id = `m${String(i)}`
		const dateAdded = Math.floor(random() * 6) * DAY
		const hidden = random() < 0.1

		titles.push(
			title(id, {
				title: pick(names),
				sortName: pick(names),
				year: pick([0, 1979, 1995, 2016]),
				dateAdded,
				roots: random() < 0.5 ? ["r1"] : random() < 0.5 ? ["r2"] : ["r1", "r2"],
				height: pick([0, 720, 1080, 2160]),
				backdrop: random() < 0.3 ? { stableUuid: `${id}-b`, uuid: `${id}-b`, name: "fanart.jpg", size: 1 } : null,
				...(hidden ? { hidden: true } : {})
			})
		)

		if (hidden) {
			continue
		}

		items.push(movieItem(id, { dateAdded }))

		if (random() < 0.3) {
			items.push(movieItem(id, { stableUuid: `${id}-v2`, uuid: `${id}-u2`, label: "4K", dateAdded: dateAdded + DAY }))
		}
	}

	for (let s = 0; s < 6; s++) {
		const id = `s${String(s)}`

		titles.push(title(id, { kind: "series", primary: null, title: pick(names), sortName: pick(names), roots: [pick(["r1", "r2"])] }))

		for (let season = random() < 0.5 ? 0 : 1; season <= 2; season++) {
			for (let episode = 1; episode <= 4; episode++) {
				items.push(episodeItem(id, season, random() < 0.1 ? 0 : episode, { dateAdded: Math.floor(random() * 8) * DAY }))
			}
		}
	}

	for (let o = 0; o < 5; o++) {
		const id = `o${String(o)}`
		const dateAdded = Math.floor(random() * 4) * DAY

		titles.push(title(id, { kind: "other", title: `clip ${String(o)}`, sortName: `clip ${String(o)}`, dateAdded, roots: ["r2"] }))
		items.push(movieItem(id, { itemClass: "other", dateAdded }))
	}

	const records: Record<string, ItemRec> = {}
	let clock = 1

	for (const item of items) {
		const roll = random()
		const at = NOW - Math.floor(random() * 400) * DAY
		const stamp = (): { t: number; d: string } => ({ t: at * HLC_SPAN + clock++, d: "dev" })

		if (roll < 0.15) {
			records[itemKey(item.stableUuid)] = { p: { ...stamp(), pos: 600, dur: 3000, lp: at, u: item.uuid, sz: item.size } }
		} else if (roll < 0.3) {
			records[itemKey(item.stableUuid)] = { w: { ...stamp(), pl: true, pc: 1 } }
		} else if (roll < 0.35) {
			const p = { ...stamp(), pos: 900, dur: 3000, lp: at, u: item.uuid, sz: item.size }

			records[itemKey(item.stableUuid)] = { p, h: { t: p.t + 1, d: "dev", v: true } }
		}
	}

	return { titles, items, records }
}

function catalogOf({ titles, items, records }: Library): VideoCatalog {
	const catalog = new VideoCatalog()

	catalog.setContent(catalogContent(titles, items))
	catalog.setWatch({ reset: true, upserts: records, removed: [], lastPlayedCeilingMs: undefined })

	return catalog
}

// One index over the whole library, every row selected from it: what the page derived before the catalog.
function referenceRows({ titles, items, records }: Library, roots: readonly string[]): HomeRowModel[] {
	const index = buildVideoHomeIndex({ items, watchItems: records, lastPlayedCeilingMs: undefined, durationOf: () => null })
	const data = { index, entries: homeEntries(index), watch: buildTitleWatch(index) }
	const byId = new Map(titles.map(item => [item.id, item]))
	const byStable = new Map(items.map(item => [item.stableUuid, item]))
	const shown = titles.filter(item => item.hidden !== true)

	return buildHomeRows({
		roots,
		continueWatching: buildContinueWatchingRow({
			resume: selectContinueWatching({ index, limit: HOME_ROW_MAX_CARDS }),
			nextUp: selectNextUp({ index, now: NOW, limit: HOME_ROW_MAX_CARDS }),
			limit: HOME_ROW_MAX_CARDS
		}),
		latestMovies: selectLatestMovies({ index, limit: HOME_ROW_MAX_CARDS }).flatMap(entry => byId.get(entry.group) ?? []),
		latestEpisodes: selectLatestEpisodes({ index, limit: HOME_ROW_MAX_CARDS }),
		topByName: (kind, limit) =>
			shown
				.filter(item => item.kind === kind)
				.sort(byTitleName)
				.slice(0, limit),
		topNewest: (filter, limit) =>
			shown
				.filter(item => ("root" in filter ? item.roots.includes(filter.root) : item.kind === filter.kind))
				.sort(newestFirst)
				.slice(0, limit),
		title: id => byId.get(id),
		item: stableUuid => byStable.get(stableUuid),
		watchData: () => data
	})
}

const NO_ART = { at: () => null }

describe("library catalog", () => {
	it.each([1, 2, 3, 4, 5, 6, 7, 8])("builds the home rows one whole-library index would (seed %i)", seed => {
		const library = randomLibrary(seed)
		const rows = catalogOf(library).home({ roots: ["r1", "r2"], now: NOW, heroPreviousId: null, heroSeed: 0 }).rows

		expect(rows).toEqual(referenceRows(library, ["r1", "r2"]))
		expect(rows.some(row => row.kind === "continue")).toBe(true)
	})

	it.each([1, 2, 3])("pages every grid sort and filter as one whole-library view would (seed %i)", seed => {
		const library = randomLibrary(seed)
		const catalog = catalogOf(library)
		const index = buildVideoHomeIndex({
			items: library.items,
			watchItems: library.records,
			lastPlayedCeilingMs: undefined,
			durationOf: () => null
		})
		const data = { index, entries: homeEntries(index), watch: buildTitleWatch(index) }
		const shown = library.titles.filter(item => item.hidden !== true)
		const variants: GridPrefs[] = [
			...GRID_SORTS.map(sort => ({ ...DEFAULT_GRID_PREFS, sort })),
			{ ...DEFAULT_GRID_PREFS, watch: "watched" },
			{ ...DEFAULT_GRID_PREFS, watch: "inProgress", sort: "played" },
			{ ...DEFAULT_GRID_PREFS, watch: "unwatched", root: "r2" }
		]

		for (const source of [{ kind: "movies" }, { kind: "shows" }, { kind: "root", uuid: "r1" }] as const) {
			for (const prefs of variants) {
				const query: GridQuery = { source, prefs, query: "", seed: 9 }
				const expected = gridView({
					titles: shown.filter(item => inGridSource(item, source)),
					prefs,
					watchOf: id => titleWatchOf(data.watch, id),
					facts: new Map(),
					seed: 9
				}).titles
				const first = catalog.grid(query, 0, 7)
				const rest = catalog.grid(query, 7, 1000)

				expect([...first.cards, ...rest.cards].map(card => card.title.id)).toEqual(expected.map(item => item.id))
				expect(first.total).toBe(expected.length)
				expect(first.cards).toEqual(expected.slice(0, 7).map(item => titleCard({ data, title: item, shape: first.shape })))
			}
		}
	})

	it("searches as one whole-library index would, with the results' cards", () => {
		const library = randomLibrary(4)
		const catalog = catalogOf(library)
		const index = buildVideoHomeIndex({
			items: library.items,
			watchItems: library.records,
			lastPlayedCeilingMs: undefined,
			durationOf: () => null
		})
		const data = { index, entries: homeEntries(index), watch: buildTitleWatch(index) }
		const search = buildVideoSearchIndex({ titles: library.titles.filter(item => item.hidden !== true), items: library.items })

		for (const query of ["heat", "a", "clip", "1995"]) {
			expect(catalog.search(query).groups).toEqual(
				searchVideos({ index: search, query }).map(group => ({
					kind: group.kind,
					total: group.total,
					cards: group.results.map(result =>
						result.item === null
							? titleCard({ data, title: result.title, shape: result.title.kind === "other" ? "landscape" : "poster" })
							: itemCard({ data, item: result.item, title: result.title })
					)
				}))
			)
		}
	})

	it("follows a watch-state change in the rows and the title's own cards", () => {
		const heat = title("heat")
		const dune = title("dune")
		const heatItem = movieItem("heat")
		const catalog = catalogOf({ titles: [heat, dune], items: [heatItem, movieItem("dune")], records: {} })
		const query: GridQuery = { source: { kind: "movies" }, prefs: DEFAULT_GRID_PREFS, query: "", seed: 0 }

		expect(catalog.grid(query, 0, 10).cards.find(card => card.title.id === "heat")?.resumeSec).toBe(0)

		catalog.setWatch({
			reset: false,
			upserts: {
				[itemKey(heatItem.stableUuid)]: {
					p: { t: NOW * HLC_SPAN, d: "dev", pos: 700, dur: 6000, lp: NOW, u: heatItem.uuid, sz: 100 }
				}
			},
			removed: [],
			lastPlayedCeilingMs: undefined
		})

		expect(catalog.grid(query, 0, 10).cards.find(card => card.title.id === "heat")?.resumeSec).toBe(700)
		expect(catalog.home({ roots: ["root"], now: NOW, heroPreviousId: null, heroSeed: 0 }).rows[0]?.cards[0]?.title.id).toBe("heat")

		catalog.setWatch({ reset: false, upserts: {}, removed: [itemKey(heatItem.stableUuid)], lastPlayedCeilingMs: undefined })

		expect(catalog.grid(query, 0, 10).cards.find(card => card.title.id === "heat")?.resumeSec).toBe(0)
	})

	it("says which titles a new classification changed, and whether anything did", () => {
		const library = randomLibrary(2)
		const catalog = catalogOf(library)
		const items = groupByTitle(library.titles, library.items)

		expect(catalog.diff(library.titles, items, NO_ART)).toEqual({ changed: false, titles: [] })

		// A title whose card art alone changed (a poster added).
		const poster = {
			candidates: [{ source: "sidecar" as const, file: 0 }],
			files: [{ stableUuid: "p", uuid: "p", name: "poster.jpg", size: 1 }]
		}

		expect(catalog.diff(library.titles, items, { at: (index, slot) => (index === 2 && slot === "p" ? poster : null) })).toEqual({
			changed: true,
			titles: [library.titles[2]?.id]
		})

		const retitled = library.titles.map(item => (item.id === "m3" ? { ...item, title: "Renamed" } : item))

		expect(catalog.diff(retitled, items, NO_ART)).toEqual({ changed: true, titles: ["m3"] })

		// A new version of a video: no title view changes, the catalog does.
		const rotated = items.map((item, index) => (index === 0 ? { ...item, uuid: "rotated" } : item))

		expect(catalog.diff(library.titles, rotated, NO_ART)).toEqual({ changed: true, titles: [] })
		expect(catalog.diff(library.titles.slice(1), items, NO_ART).titles).toEqual([library.titles[0]?.id])
	})

	it.each([1, 2, 3, 4])("groups videos by title in title order, each title's own order kept (seed %i)", seed => {
		const random = rng(seed)
		const titles = Array.from({ length: 50 }, (_, i) => title(`t${String(i)}`))
		const ordered = random() < 0.5
		const items = Array.from({ length: 400 }, (_, i) => {
			const owner = ordered ? Math.floor((i / 400) * 52) : Math.floor(random() * 52)

			// Titles t50 and t51 don't exist: their videos are left out.
			return movieItem(`t${String(owner)}`, { stableUuid: `v${String(i)}`, uuid: `u${String(i)}` })
		})
		const expected = titles.flatMap(item => items.filter(video => video.titleId === item.id))
		const view = arrayItems(titles, items)

		expect(groupByTitle(titles, items)).toEqual(expected)
		expect(Array.from({ length: view.length }, (_, index) => [view.at(index), titles[view.title(index)]?.id])).toEqual(
			expected.map(video => [video, video.titleId])
		)
	})

	it("caps rows by kind and directory from its orders, never sorting at query time", () => {
		const titles = Array.from({ length: 60 }, (_, i) => title(`m${String(i).padStart(2, "0")}`, { dateAdded: i }))
		const catalog = catalogOf({ titles, items: titles.map(item => movieItem(item.id, { dateAdded: item.dateAdded })), records: {} })
		const rows = catalog.home({ roots: ["root"], now: NOW, heroPreviousId: null, heroSeed: 0 }).rows

		expect(rows.find(row => row.kind === "movies")?.cards).toHaveLength(HOME_ROW_MAX_CARDS)
		expect(rows.find(row => row.kind === "root")?.cards.map(card => card.title.id)).toEqual(
			titles
				.slice(-HOME_ROOT_ROW_CARDS)
				.reverse()
				.map(item => item.id)
		)
	})
})

describe("library catalog grid orders", () => {
	const warm = (catalog: VideoCatalog, query: GridQuery, cancelled = () => false): void => {
		const steps = catalog.gridWarmSteps(query, cancelled)

		while (steps.next().done !== true) {
			// Each step is one slice of the worker's.
		}
	}
	const query = (prefs: Partial<GridPrefs>, text = ""): GridQuery => ({
		source: { kind: "movies" },
		prefs: { ...DEFAULT_GRID_PREFS, ...prefs },
		query: text,
		seed: 9
	})

	it("keeps an order across watch and probe changes it doesn't read, and drops it on those it does", () => {
		const library = randomLibrary(4)
		const catalog = catalogOf(library)
		const byTitle = query({})
		const played = query({ sort: "played" })
		const runtime = query({ sort: "runtime" })
		const movie = library.items.find(item => item.itemClass === "movie")

		for (const each of [byTitle, played, runtime]) {
			warm(catalog, each)
		}

		expect([catalog.hasGrid(byTitle), catalog.hasGrid(played), catalog.hasGrid(runtime)]).toEqual([true, true, true])

		catalog.setWatch({
			reset: false,
			upserts: { [itemKey("new")]: { w: { t: 1, d: "dev", pl: true, pc: 1 } } },
			removed: [],
			lastPlayedCeilingMs: undefined
		})

		expect([catalog.hasGrid(byTitle), catalog.hasGrid(played), catalog.hasGrid(runtime)]).toEqual([true, false, true])

		// Facts it already holds change nothing; new ones drop the orders reading them.
		const facts = { reset: false, uuids: [movie?.uuid ?? ""], durations: [5400], heights: [1080] }

		expect(catalog.setFacts(facts)).toMatchObject({ changed: true })

		warm(catalog, played)
		warm(catalog, runtime)

		expect(catalog.setFacts(facts)).toEqual({ changed: false, affectsWatch: false })
		expect([catalog.hasGrid(byTitle), catalog.hasGrid(played), catalog.hasGrid(runtime)]).toEqual([true, true, true])
		expect(catalog.setFacts({ ...facts, durations: [60] }).changed).toBe(true)
		expect([catalog.hasGrid(byTitle), catalog.hasGrid(runtime)]).toEqual([true, false])
	})

	it("keeps the base orders while searches come and go, and keeps nothing for a cancelled build", () => {
		const catalog = catalogOf(randomLibrary(5))
		const base = query({})

		warm(catalog, base)

		for (const text of ["a", "al", "ali", "alie", "alien"]) {
			warm(catalog, query({}, text))
		}

		expect(catalog.hasGrid(base)).toBe(true)
		expect(catalog.hasGrid(query({}, "alien"))).toBe(true)
		expect(catalog.hasGrid(query({}, "a"))).toBe(false)

		const cancelled = query({ sort: "year" })

		warm(catalog, cancelled, () => true)

		expect(catalog.hasGrid(cancelled)).toBe(false)
	})
})

describe("persisted library catalog", () => {
	const library = randomLibrary(3)
	const content: CatalogContent = catalogContent(library.titles, library.items)
	const key = catalogKey("roots+corrections")

	it("round-trips every title, video and selection, unicode included", () => {
		const titles = allTitles(content.titles).map((item, index) =>
			index === 0 ? { ...item, title: "Amélie – 東京 🎬", sortName: "Amélie", detailHash: 4_000_000_000 } : item
		)
		const bytes = encodeCatalog({ content: { ...content, titles: arrayTitles(titles) }, key, generation: 5n })
		const decoded = decodeCatalog(bytes, key)

		expect(decoded === null ? null : allTitles(decoded.titles)).toEqual(titles)
		expect(decoded?.titles.id(2)).toBe(titles[2]?.id)
		expect(decoded === null ? null : allItems(decoded.items)).toEqual(allItems(content.items))
		// Built only when read; a title's videos name it by index.
		expect(decoded?.items.title(0)).toBe(content.items.title(0))
		expect(decoded?.items.stableUuid(1)).toBe(content.items.stableUuid(1))
		expect(decoded?.derived).toEqual(content.derived)
	})

	it("round-trips each title's card art candidates", () => {
		const poster = { stableUuid: "p-s", uuid: "p-u", name: "poster.jpg", size: 12 }
		const titles = allTitles(content.titles)
		const art = titles.map((_, index) =>
			index === 0
				? [
						{
							candidates: [
								{ source: "sidecar" as const, file: 0 },
								{ source: "typographic" as const },
								{ source: "frame" as const, file: 1 }
							],
							files: [poster, titles[0]?.primary ?? poster]
						},
						null
					]
				: [null, { candidates: [{ source: "gradient" as const }], files: [] }]
		)
		const decoded = decodeCatalog(
			encodeCatalog({
				content: { ...content, art: { at: (title, slot) => art[title]?.[slot === "p" ? 0 : 1] ?? null } },
				key,
				generation: 2n
			}),
			key
		)

		expect(decoded?.art.at(0, "p")).toEqual(art[0]?.[0])
		expect(decoded?.art.at(0, "s")).toBeNull()
		expect(decoded?.art.at(1, "s")).toEqual({ candidates: [{ source: "gradient" }], files: [] })
		expect(decoded?.art.at(1, "p")).toBeNull()
	})

	it("rejects a torn or truncated write and one written under another key", () => {
		const bytes = encodeCatalog({ content, key, generation: 1n })
		const flipped = bytes.slice()

		flipped[CATALOG_HEADER_BYTES + 3] = (flipped[CATALOG_HEADER_BYTES + 3] ?? 0) ^ 1

		expect(decodeCatalog(flipped, key)).toBeNull()
		expect(decodeCatalog(bytes.slice(0, bytes.length - 8), key)).toBeNull()
		expect(decodeCatalog(bytes, catalogKey("other roots"))).toBeNull()
		expect(decodeCatalog(new Uint8Array(16), key)).toBeNull()
	})

	it("orders the slots to read newest first, leaving out a missing header", () => {
		expect(catalogSlotOrder([4n, 5n])).toEqual([1, 0])
		expect(catalogSlotOrder([7n, null])).toEqual([0])
		expect(catalogSlotOrder([null, null])).toEqual([])
	})
})
