import { useTranslation } from "react-i18next"
import { useNavigate } from "@tanstack/react-router"
import { toast } from "sonner"
import { log } from "@/lib/log"
import { watchState } from "@/features/videos/lib/watchState"
import { markTargets } from "@/features/videos/lib/videoCard.logic"
import type { VideoCardModel } from "@/features/videos/lib/homeRows.logic"
import type { VideoCardAction } from "@/features/videos/components/videoCard"
import type { VideoHomeItem } from "@/workers/videoIndex/view.logic"

// What a card's Play button, menu and keys do: open the player, or change the watch state (journaled locally, synced
// on the watch state's own schedule).
export function useVideoCardActions(items: readonly VideoHomeItem[] | null) {
	const { t } = useTranslation("videos")
	const navigate = useNavigate()

	function report(work: Promise<void>): void {
		work.catch((e: unknown) => {
			log.warn("videos.card", "watch-state change failed", e)
			toast.error(t("videosCardWatchStateFailed"))
		})
	}

	return (action: VideoCardAction, card: VideoCardModel): void => {
		switch (action) {
			case "play":
				if (card.play !== null) {
					void navigate({ to: "/videos/watch/$uuid", params: { uuid: card.title.id }, search: { file: card.play } })
				}

				return
			case "markWatched":
				report(watchState.markWatched({ stableUuids: markTargets(card, items ?? []) }))

				return
			case "markUnwatched":
				report(watchState.markUnwatched({ stableUuids: markTargets(card, items ?? []) }))

				return
			case "hide": {
				const stableUuid = card.item?.stableUuid ?? card.play

				if (stableUuid !== null) {
					report(watchState.hide({ stableUuid }))
				}

				return
			}
		}
	}
}
