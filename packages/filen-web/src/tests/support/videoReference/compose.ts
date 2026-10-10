import type { ItemRec } from "@filen/shared/videos"
import type { VideoHomeItem, VideoTitleView } from "@/workers/videoIndex/view.logic"
import { buildTitleWatch, buildVideoHomeIndex } from "@/features/videos/lib/titleWatch.logic"
import { buildVideoSearchIndex, filterTitlesByQuery, searchVideos, titleSearchTokens } from "@/features/videos/lib/search.logic"
import {
	buildHomeRows,
	homeEntries,
	itemCard,
	pickHero,
	rowShape,
	titleCard,
	type CardShape,
	type HeroModel,
	type HomeRowModel,
	type VideoCardModel
} from "@/tests/support/videoReference/homeRows"
import {
	gridView,
	sourceTitles,
	type GridPrefs,
	type GridSource,
	type GridView,
	type ProbeFacts
} from "@/tests/support/videoReference/videoGrid"

// How the page composed its screens when it held the whole library (one watch index over every video, the titles in
// the order the worker sent them): what the worker's catalog must answer for the same library and state.

export interface ReferenceLibrary {
	// Every title, hidden ones included, in the order the worker sent them.
	titles: readonly VideoTitleView[]
	items: readonly VideoHomeItem[]
	records: Record<string, ItemRec>
	lastPlayedCeilingMs?: number | undefined
	facts?: ReadonlyMap<string, ProbeFacts>
}

function homeData(library: ReferenceLibrary) {
	const facts = library.facts ?? new Map<string, ProbeFacts>()
	const index = buildVideoHomeIndex({
		items: library.items,
		watchItems: library.records,
		lastPlayedCeilingMs: library.lastPlayedCeilingMs,
		durationOf: uuid => facts.get(uuid)?.durationSec ?? null
	})

	return {
		shown: library.titles.filter(title => title.hidden !== true),
		byId: new Map(library.titles.map(title => [title.id, title])),
		index,
		entries: homeEntries(index),
		watch: buildTitleWatch(index)
	}
}

export function referenceHome(
	library: ReferenceLibrary,
	query: { roots: readonly string[]; now: number; heroPreviousId: string | null; heroSeed: number }
): { rows: HomeRowModel[]; hero: HeroModel | null } {
	const data = homeData(library)
	const rows = buildHomeRows({
		titles: data.shown,
		byId: data.byId,
		items: library.items,
		index: data.index,
		entries: data.entries,
		watch: data.watch,
		roots: query.roots,
		now: query.now
	})

	return { rows, hero: pickHero({ rows, titles: data.shown, previousId: query.heroPreviousId, seed: query.heroSeed }) }
}

export function referenceGrid(
	library: ReferenceLibrary,
	query: { source: GridSource; prefs: GridPrefs; query: string; seed: number }
): { titles: VideoTitleView[]; coverage: GridView["coverage"]; shape: CardShape; sourceCount: number; cards: VideoCardModel[] } {
	const data = homeData(library)
	const titles = sourceTitles(data.shown, query.source)
	const searched = filterTitlesByQuery({ titles, query: query.query, tokens: titleSearchTokens(titles) })
	const view = gridView({ titles: searched, prefs: query.prefs, watch: data.watch, facts: library.facts ?? new Map(), seed: query.seed })
	const shape: CardShape = query.source.kind === "other" ? "landscape" : query.source.kind === "root" ? rowShape(titles) : "poster"

	return {
		titles: view.titles,
		coverage: view.coverage,
		shape,
		sourceCount: titles.length,
		cards: view.titles.map(title => titleCard({ data, title, shape }))
	}
}

export function referenceSearch(library: ReferenceLibrary, query: string): { kind: string; total: number; cards: VideoCardModel[] }[] {
	const data = homeData(library)
	const index = buildVideoSearchIndex({ titles: data.shown, items: library.items })

	return searchVideos({ index, query }).map(group => ({
		kind: group.kind,
		total: group.total,
		cards: group.results.map(result =>
			result.item === null
				? titleCard({ data, title: result.title, shape: result.title.kind === "other" ? "landscape" : "poster" })
				: itemCard({ data, item: result.item, title: result.title })
		)
	}))
}
