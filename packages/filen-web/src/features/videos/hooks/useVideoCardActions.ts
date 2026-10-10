import { useTranslation } from "react-i18next"
import { useNavigate } from "@tanstack/react-router"
import { toast } from "sonner"
import { log } from "@/lib/log"
import { watchState } from "@/features/videos/lib/watchState"
import { videoLibrary } from "@/features/videos/lib/index"
import type { VideoCardModel } from "@/features/videos/lib/homeRows.logic"
import type { VideoCardAction } from "@/features/videos/components/videoCard"

// The videos a card's Mark watched / unwatched applies to: every version of its episode or movie, every episode of
// its series (looked up in the library's worker).
function markTargets(card: Pick<VideoCardModel, "title" | "item">): Promise<string[]> {
	return videoLibrary.markTargets(card.title.id, card.item?.group ?? null)
}

// What a card's Play button, menu and keys do: open the player, or change the watch state (journaled locally, synced
// on the watch state's own schedule).
export function useVideoCardActions() {
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
				report(markTargets(card).then(stableUuids => watchState.markWatched({ stableUuids })))

				return
			case "markUnwatched":
				report(markTargets(card).then(stableUuids => watchState.markUnwatched({ stableUuids })))

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
