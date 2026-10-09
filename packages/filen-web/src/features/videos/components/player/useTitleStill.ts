import { useEffect, useState } from "react"
import { log } from "@/lib/log"
import { readStoreBlob } from "@/lib/storage/opfsBlobStore"
import { videoLibrary } from "@/features/videos/lib/index"
import { artLoader } from "@/features/videos/lib/artwork"
import { ART_STORE, artKey } from "@/features/videos/lib/artStore.logic"

// The Up Next card's still, decoded only while the card shows: the next item's own frame when this browser has
// one stored (captured on a pause, or generated for its card), else the title's still. Stored frames are local
// reads; nothing is generated or downloaded for the card.
export function useUpNextStill(params: { titleId: string; title: string; fileUuid: string | null; enabled: boolean }): string | null {
	const { titleId, title, fileUuid, enabled } = params
	const [url, setUrl] = useState<string | null>(null)

	useEffect(() => {
		if (!enabled) {
			return
		}

		const controller = new AbortController()
		// A function, so each check reads the signal anew across the awaits.
		const aborted = (): boolean => controller.signal.aborted
		let created: string | null = null

		const show = (blob: Blob): void => {
			created = URL.createObjectURL(blob)
			setUrl(created)
		}

		void (async () => {
			try {
				if (fileUuid !== null) {
					for (const variant of ["capture", "frame"] as const) {
						const stored = await readStoreBlob(ART_STORE, artKey(fileUuid, "s", variant))

						if (aborted()) {
							return
						}

						if (stored !== null) {
							show(stored)

							return
						}
					}
				}

				const art = await videoLibrary.titleArt(titleId, "s")

				if (art === null || aborted()) {
					return
				}

				const result = await artLoader.load({
					candidates: art.candidates,
					slot: "s",
					title,
					resolveFile: art.resolveFile,
					signal: controller.signal
				})

				if (result?.type === "image" && !aborted()) {
					show(result.blob)
				}
			} catch (error) {
				log.warn("videos.player", "up next still failed", error)
			}
		})()

		return () => {
			controller.abort()

			if (created !== null) {
				URL.revokeObjectURL(created)
			}

			setUrl(null)
		}
	}, [titleId, title, fileUuid, enabled])

	return url
}
