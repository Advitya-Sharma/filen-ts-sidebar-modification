import { describe, expect, it } from "vitest"
import { HLC_SPAN, itemKey, type ItemRec } from "@filen/shared/videos"
import { VideoCatalog, catalogContent, type CatalogContent, type GridQuery } from "@/workers/videoIndex/catalog.logic"
import { catalogKey, decodeCatalog, encodeCatalog } from "@/workers/videoIndex/catalogFile.logic"
import type { VideoHomeItem, VideoTitleView } from "@/workers/videoIndex/view.logic"
import { DEFAULT_GRID_PREFS, GRID_SORTS, type GridPrefs } from "@/features/videos/lib/videoGrid.logic"
import { episodeItem, movieItem, title } from "@/tests/support/videoHome"
import { referenceGrid, referenceHome, referenceSearch, type ReferenceLibrary } from "@/tests/support/videoReference/compose"

// The worker's catalog against the page's own composition from when it held the whole library (frozen under
// tests/support/videoReference): the same rows, hero, grids and search results for the same library and state.

const NOW = Date.UTC(2026, 9, 1)
const DAY = 86_400_000

function rng(seed: number): () => number {
	let state = seed >>> 0

	return () => {
		state = (state * 1_664_525 + 1_013_904_223) >>> 0

		return state / 2 ** 32
	}
}

// Ties in names and dates, versions, series with specials, multi-part and titled episodes, other videos, hidden
// titles and shown titles without a video, two roots; watch records of every kind, probe facts on some files, and a
// last-played ceiling on some seeds.
function randomLibrary(seed: number): ReferenceLibrary {
	const random = rng(seed)
	const pick = <T>(values: readonly T[]): T => values[Math.floor(random() * values.length)] as T
	const titles: VideoTitleView[] = []
	const items: VideoHomeItem[] = []
	const names = ["Alien", "Arrival", "Heat", "The Matrix", "Dune", "Ronin", "Up", "Zodiac", "Léon"]

	for (let i = 0; i < 40; i++) {
		const id = `m${String(i)}`
		const dateAdded = Math.floor(random() * 6) * DAY
		const hidden = random() < 0.1
		const empty = !hidden && random() < 0.08

		titles.push(
			title(id, {
				title: pick(names),
				sortName: pick(names),
				year: pick([0, 1979, 1995, 2016]),
				dateAdded,
				roots: random() < 0.5 ? ["r1"] : random() < 0.5 ? ["r2"] : ["r1", "r2"],
				height: pick([0, 0, 720, 1080, 2160]),
				backdrop: random() < 0.3 ? { stableUuid: `${id}-b`, uuid: `${id}-b`, name: "fanart.jpg", size: 1 } : null,
				unsupported: random() < 0.05,
				...(hidden ? { hidden: true } : {})
			})
		)

		if (hidden || empty) {
			continue
		}

		items.push(movieItem(id, { dateAdded }))

		if (random() < 0.3) {
			items.push(movieItem(id, { stableUuid: `${id}-v2`, uuid: `${id}-u2`, label: "4K", dateAdded: dateAdded + DAY }))
		}
	}

	for (let s = 0; s < 8; s++) {
		const id = `s${String(s)}`

		titles.push(
			title(id, {
				kind: "series",
				primary: null,
				title: pick(names),
				sortName: pick(names),
				roots: [pick(["r1", "r2"])],
				backdrop: random() < 0.4 ? { stableUuid: `${id}-b`, uuid: `${id}-b`, name: "fanart.jpg", size: 1 } : null
			})
		)

		if (random() < 0.1) {
			continue
		}

		for (let season = random() < 0.5 ? 0 : 1; season <= 2; season++) {
			for (let episode = 1; episode <= 4; episode++) {
				const twoPart = random() < 0.15

				items.push(
					episodeItem(id, season, random() < 0.1 ? 0 : episode, {
						dateAdded: Math.floor(random() * 8) * DAY,
						...(twoPart ? { episodeEnd: episode + 1 } : {}),
						...(random() < 0.3 ? { episodeTitle: pick(["Pilot", "Finale", "Homecoming"]) } : {})
					})
				)

				if (twoPart) {
					episode++
				}
			}
		}
	}

	for (let o = 0; o < 6; o++) {
		const id = `o${String(o)}`
		const dateAdded = Math.floor(random() * 4) * DAY

		titles.push(title(id, { kind: "other", title: `clip ${String(o)}`, sortName: `clip ${String(o)}`, dateAdded, roots: ["r2"] }))
		items.push(movieItem(id, { itemClass: "other", dateAdded, folder: pick(["Camera", "Holiday Lisbon"]) }))
	}

	const records: Record<string, ItemRec> = {}
	const facts = new Map<string, { durationSec: number | null; height: number }>()
	let clock = 1

	for (const item of items) {
		const roll = random()
		const at = NOW - Math.floor(random() * 400) * DAY
		const stamp = (): { t: number; d: string } => ({ t: at * HLC_SPAN + clock++, d: "dev" })
		// Some records carry no runtime: the probed one decides then.
		const dur = random() < 0.3 ? 0 : 3000

		if (roll < 0.15) {
			records[itemKey(item.stableUuid)] = { p: { ...stamp(), pos: 600, dur, lp: at, u: item.uuid, sz: item.size } }
		} else if (roll < 0.3) {
			records[itemKey(item.stableUuid)] = { w: { ...stamp(), pl: true, pc: 1 } }
		} else if (roll < 0.35) {
			const p = { ...stamp(), pos: 2900, dur, lp: at, u: item.uuid, sz: item.size }

			records[itemKey(item.stableUuid)] = { p, h: { t: p.t + 1, d: "dev", v: true } }
		}

		if (random() < 0.5) {
			facts.set(item.uuid, {
				durationSec: random() < 0.2 ? null : 600 + Math.floor(random() * 6000),
				height: pick([0, 480, 720, 1080, 2160])
			})
		}
	}

	return { titles, items, records, facts, lastPlayedCeilingMs: seed % 3 === 0 ? NOW - 200 * DAY : undefined }
}

function catalogFor(library: ReferenceLibrary, content: CatalogContent = catalogContent(library.titles, library.items)): VideoCatalog {
	const catalog = new VideoCatalog()
	const facts = [...(library.facts ?? new Map<string, { durationSec: number | null; height: number }>()).entries()]

	catalog.setContent(content)
	catalog.setWatch({ reset: true, upserts: library.records, removed: [], lastPlayedCeilingMs: library.lastPlayedCeilingMs })
	catalog.setFacts({
		reset: true,
		uuids: facts.map(([uuid]) => uuid),
		durations: facts.map(([, fact]) => fact.durationSec),
		heights: facts.map(([, fact]) => fact.height)
	})

	return catalog
}

// The catalog as the next session loads it: written to the persisted file and read back.
function persisted(library: ReferenceLibrary): CatalogContent {
	const key = catalogKey("roots")
	const content = decodeCatalog(encodeCatalog({ content: catalogContent(library.titles, library.items), key, generation: 1n }), key)

	if (content === null) {
		throw new Error("the persisted catalog did not read back")
	}

	return content
}

const PREFS: GridPrefs[] = [
	...GRID_SORTS.map(sort => ({ ...DEFAULT_GRID_PREFS, sort })),
	{ ...DEFAULT_GRID_PREFS, watch: "watched" },
	{ ...DEFAULT_GRID_PREFS, watch: "inProgress", sort: "played" },
	{ ...DEFAULT_GRID_PREFS, watch: "unwatched", root: "r2" },
	{ ...DEFAULT_GRID_PREFS, quality: "1080" },
	{ ...DEFAULT_GRID_PREFS, quality: "2160", sort: "runtime" },
	{ ...DEFAULT_GRID_PREFS, quality: "sd", watch: "unwatched" }
]
const SOURCES: GridQuery["source"][] = [{ kind: "movies" }, { kind: "shows" }, { kind: "other" }, { kind: "root", uuid: "r1" }]
const TEXT = ["", "al", "the matrix", "1995", "clip 3", "lisbon", "zz"]
const SEARCHES = ["alien", "heat 1995", "pilot", "homecoming", "clip", "lisbon", "season", "zz"]

function expectSameScreens(catalog: VideoCatalog, library: ReferenceLibrary): void {
	for (const heroPreviousId of [null, "m3", "s1", "o2"]) {
		for (const heroSeed of [0, 0.37, 0.99]) {
			const query = { roots: ["r1", "r2"], now: NOW, heroPreviousId, heroSeed }

			expect(catalog.home(query)).toEqual(referenceHome(library, query))
		}
	}

	for (const source of SOURCES) {
		for (const prefs of PREFS) {
			for (const text of TEXT) {
				const query: GridQuery = { source, prefs, query: text, seed: 9 }
				const expected = referenceGrid(library, query)
				const page = catalog.grid(query, 0, 1000)

				expect(page.cards).toEqual(expected.cards)
				expect({ total: page.total, coverage: page.coverage, shape: page.shape, sourceCount: page.sourceCount }).toEqual({
					total: expected.titles.length,
					coverage: expected.coverage,
					shape: expected.shape,
					sourceCount: expected.sourceCount
				})
			}
		}
	}

	for (const text of SEARCHES) {
		expect(catalog.search(text).groups).toEqual(referenceSearch(library, text))
	}
}

describe("library catalog against the page's own composition", () => {
	it.each([1, 2, 3, 4, 5, 6])("answers the home, grids and search the page did (seed %i)", seed => {
		const library = randomLibrary(seed)

		expectSameScreens(catalogFor(library), library)
	})

	it.each([7, 8, 9])("answers them from the persisted catalog as well (seed %i)", seed => {
		const library = randomLibrary(seed)

		expectSameScreens(catalogFor(library, persisted(library)), library)
	})

	it("covers what the comparison is meant to reach", () => {
		const library = randomLibrary(3)
		const catalog = catalogFor(library)
		const home = catalog.home({ roots: ["r1", "r2"], now: NOW, heroPreviousId: null, heroSeed: 0.5 })
		const shown = library.titles.filter(item => item.hidden !== true)

		expect(library.lastPlayedCeilingMs).toBeDefined()
		expect(home.hero).not.toBeNull()
		expect(home.rows.map(row => row.kind)).toEqual(
			expect.arrayContaining(["continue", "latestMovies", "latestEpisodes", "movies", "shows", "root", "other"])
		)
		expect(library.items.some(item => item.episodeEnd > item.episode)).toBe(true)
		expect(shown.some(item => !library.items.some(video => video.titleId === item.id))).toBe(true)
		expect(
			catalog.grid({ source: { kind: "movies" }, prefs: { ...DEFAULT_GRID_PREFS, sort: "runtime" }, query: "", seed: 9 }, 0, 0)
				.coverage?.known
		).toBeGreaterThan(0)
	})

	it("follows watch and probe changes as a page rebuilding from scratch would", () => {
		const library = randomLibrary(4)
		const catalog = catalogFor(library)
		const movie = library.items.find(item => item.itemClass === "movie" && library.records[itemKey(item.stableUuid)] !== undefined)
		const records = Object.fromEntries(Object.entries(library.records).filter(([key]) => key !== itemKey(movie?.stableUuid ?? "")))

		const facts = new Map(library.facts)

		facts.set(movie?.uuid ?? "", { durationSec: 42, height: 2160 })
		catalog.setWatch({ reset: false, upserts: {}, removed: [itemKey(movie?.stableUuid ?? "")], lastPlayedCeilingMs: NOW - 30 * DAY })
		catalog.setFacts({ reset: false, uuids: [movie?.uuid ?? ""], durations: [42], heights: [2160] })

		expectSameScreens(catalog, { ...library, records, facts, lastPlayedCeilingMs: NOW - 30 * DAY })
	})
})
