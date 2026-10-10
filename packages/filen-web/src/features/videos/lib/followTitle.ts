import { sdkApi } from "@/lib/sdk/client"
import { log } from "@/lib/log"
import { videoLibrary } from "@/features/videos/lib/index"

// After a correction the item may be shown as another title; the page follows once the index has it (a
// classification is local and quick), and stays where it is if it never turns up.
export const FOLLOW_TIMEOUT_MS = 5_000

// Asks the library's worker again each time its catalog changes, until the title is there or the wait is over.
// Returns the cancel: the page that asked calls it when it goes, so `then` never runs on another page.
export function whenTitleKnown(id: string, then: () => void): () => void {
	let done = false
	let version = -1
	let unsubscribe: () => void = () => undefined
	const finish = (): void => {
		done = true
		clearTimeout(timer)
		unsubscribe()
	}
	const timer = setTimeout(finish, FOLLOW_TIMEOUT_MS)
	const check = (): void => {
		const current = videoLibrary.getSnapshot().version

		if (done || current === version) {
			return
		}

		version = current
		sdkApi.videoCatalogTitle(id).then(
			title => {
				if (title !== null && !done) {
					finish()
					then()
				}
			},
			(error: unknown) => {
				log.warn("videos.detail", "corrected title lookup failed", error)
			}
		)
	}

	unsubscribe = videoLibrary.subscribe(check)
	check()

	return finish
}
