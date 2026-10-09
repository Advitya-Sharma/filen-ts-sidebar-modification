import { useEffect, useState, useSyncExternalStore } from "react"
import { videoLibrary } from "@/features/videos/lib/index"
import { watchState } from "@/features/videos/lib/watchState"
import { probeService } from "@/features/videos/lib/probeService"
import { buildTitleWatch, buildVideoHomeIndex } from "@/features/videos/lib/titleWatch.logic"
import { homeEntries } from "@/features/videos/lib/homeRows.logic"

function probedDuration(uuid: string): number | null {
	return probeService.peek(uuid)?.probe?.durationSec ?? null
}

// The library's titles and videos joined with the watch state: what the home rows, the grids' watch filters and
// search read. All of it is held already; the videos are read from the index's worker once per classification.
export function useVideoHomeData() {
	const snapshot = useSyncExternalStore(videoLibrary.subscribe, videoLibrary.getSnapshot)
	const items = useSyncExternalStore(videoLibrary.subscribe, videoLibrary.getHomeItems)
	const watch = useSyncExternalStore(watchState.subscribe, watchState.getSnapshot)
	// Fixed per mount, so the rows don't move while they are looked at.
	const [now] = useState(() => Date.now())

	// The worker sends the library's videos while a screen that reads them is mounted.
	useEffect(() => videoLibrary.retainHome(), [])

	const index = buildVideoHomeIndex({
		items: items ?? [],
		watchItems: watch.doc.items,
		lastPlayedCeilingMs: watch.lastPlayedCeilingMs,
		durationOf: probedDuration
	})

	return {
		snapshot,
		items,
		index,
		entries: homeEntries(index),
		watch: buildTitleWatch(index),
		now
	}
}
