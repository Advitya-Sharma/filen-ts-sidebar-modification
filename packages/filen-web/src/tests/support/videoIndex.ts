import type { CatalogStore, VideoIndexEngine } from "@/workers/videoIndex/engine"
import { readCatalogHeader } from "@/workers/videoIndex/catalogFile.logic"
import { DEFAULT_GRID_PREFS } from "@/features/videos/lib/videoGrid.logic"
import type { VideoTitleView } from "@/workers/videoIndex/view.logic"

// The library index's persisted catalog in memory, and every title the engine's catalog holds, through its public
// queries (a grid per kind plus the hidden ones).

export function memoryCatalogStore(): CatalogStore & { slots: [Uint8Array | null, Uint8Array | null]; writes: number; reads: number } {
	const store = {
		slots: [null, null] as [Uint8Array | null, Uint8Array | null],
		writes: 0,
		reads: 0,
		readSlot: (slot: 0 | 1) => {
			store.reads++

			return Promise.resolve(store.slots[slot])
		},
		readGenerations: () =>
			Promise.resolve(
				store.slots.map(slot => (slot === null ? null : (readCatalogHeader(slot)?.generation ?? null))) as [
					bigint | null,
					bigint | null
				]
			),
		write: (slot: 0 | 1, bytes: Uint8Array<ArrayBuffer>, cancelled: () => boolean) => {
			if (!cancelled()) {
				store.slots[slot] = bytes
				store.writes++
			}

			return Promise.resolve()
		}
	}

	return store
}

export async function catalogTitles(engine: VideoIndexEngine): Promise<VideoTitleView[]> {
	const out: VideoTitleView[] = []

	for (const kind of ["movies", "shows", "other"] as const) {
		const page = await engine.catalogGrid({ source: { kind }, prefs: DEFAULT_GRID_PREFS, query: "", seed: 0 }, 0, 1_000_000)

		out.push(...page.cards.map(card => card.title))
	}

	return [...out, ...(await engine.catalogHidden())]
}

// Keeps `into` holding the catalog's titles, refreshed on each catalog change.
export function trackCatalogTitles(engine: VideoIndexEngine, into: Map<string, VideoTitleView>): void {
	engine.subscribe(update => {
		if (update.type !== "catalog") {
			return
		}

		void catalogTitles(engine).then(titles => {
			into.clear()

			for (const title of titles) {
				into.set(title.id, title)
			}
		})
	})
}
