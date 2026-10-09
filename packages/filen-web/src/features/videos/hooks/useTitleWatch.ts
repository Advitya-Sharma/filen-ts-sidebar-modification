import { useSyncExternalStore } from "react"
import { watchState } from "@/features/videos/lib/watchState"
import { probeService } from "@/features/videos/lib/probeService"
import { buildVideoHomeIndex } from "@/features/videos/lib/titleWatch.logic"
import { entriesByStableUuid, titleHomeItems } from "@/features/videos/lib/titleDetail.logic"
import type { VideoTitleDetail } from "@/workers/videoIndex/view.logic"

function probedDuration(uuid: string): number | null {
	return probeService.peek(uuid)?.probe?.durationSec ?? null
}

// One title's watch state: its own videos joined with the records held already (no request, and nothing of the
// rest of the library).
export function useTitleWatch(detail: VideoTitleDetail) {
	const watch = useSyncExternalStore(watchState.subscribe, watchState.getSnapshot)
	const index = buildVideoHomeIndex({
		items: titleHomeItems(detail),
		watchItems: watch.doc.items,
		lastPlayedCeilingMs: watch.lastPlayedCeilingMs,
		durationOf: probedDuration
	})

	return { index, entries: entriesByStableUuid(index), doc: watch.doc }
}
