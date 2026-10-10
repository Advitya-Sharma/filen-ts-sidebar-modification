import { useEffect, useState } from "react"
import { log } from "@/lib/log"
import { readStoreBlob } from "@/lib/storage/opfsBlobStore"
import { videoLibrary } from "@/features/videos/lib/index"
import { artLoader } from "@/features/videos/lib/artwork"
import { ART_STORE, artKey } from "@/features/videos/lib/artStore.logic"

// The Up Next card's still, decoded only while the card shows: the next item's own frame when this browser has
// one stored (captured on a pause, or generated for its card), else the title's still. Stored frames are local
// reads, as is the title's still (an image stored earlier or the drive's cached thumbnail); nothing is generated or
// downloaded for the card.

// The stored frame of `fileUuid`, else the title's still; null when there is none or the card went away first.
async function loadStill({
	titleId,
	title,
	fileUuid,
	signal
}: {
	titleId: string
	title: string
	fileUuid: string | null
	signal: AbortSignal
}): Promise<Blob | null> {
	if (fileUuid !== null) {
		for (const variant of ["capture", "frame"] as const) {
			const stored = await readStoreBlob(ART_STORE, artKey(fileUuid, "s", variant))

			if (signal.aborted || stored !== null) {
				return stored
			}
		}
	}

	const art = await videoLibrary.titleArt(titleId, "s")

	if (art === null || signal.aborted) {
		return null
	}

	// Local only: a candidate that would need reads ("network") leaves the card without a still.
	const result = await artLoader.loadLocal({
		candidates: art.candidates,
		slot: "s",
		title,
		resolveFile: art.resolveFile,
		fileUuid: art.fileUuid,
		signal
	})

	return result !== null && result !== "network" && result.type === "image" ? result.blob : null
}

export function useUpNextStill(params: { titleId: string; title: string; fileUuid: string | null; enabled: boolean }): string | null {
	const { titleId, title, fileUuid, enabled } = params
	const [url, setUrl] = useState<string | null>(null)

	useEffect(() => {
		if (!enabled) {
			return
		}

		const controller = new AbortController()
		let created: string | null = null

		loadStill({ titleId, title, fileUuid, signal: controller.signal }).then(
			blob => {
				if (blob !== null && !controller.signal.aborted) {
					created = URL.createObjectURL(blob)
					setUrl(created)
				}
			},
			(error: unknown) => {
				log.warn("videos.player", "up next still failed", error)
			}
		)

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
