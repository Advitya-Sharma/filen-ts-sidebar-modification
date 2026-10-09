import { useEffect, useState } from "react"
import { log } from "@/lib/log"
import type { Rgb } from "@/lib/media/framePick.logic"
import { videoLibrary, type VideoArtCandidates } from "@/features/videos/lib/index"
import { artLoader, type ArtLoadRequest, type ArtResult } from "@/features/videos/lib/artwork"
import type { ArtSlotCode } from "@/features/videos/lib/artStore.logic"
import { CARD_PROBE_DWELL_MS, PROBE_RANK } from "@/features/videos/lib/probeService.logic"
import { visibleCardObserver } from "@/features/videos/lib/cardObserver"

export type VideoArt =
	// Not known yet: the placeholder shows.
	| { kind: "none" }
	| { kind: "image"; url: string; fit: "cover" | "contain"; dominant: Rgb | null }
	// No image: the title-set card.
	| { kind: "text" }

const NONE: VideoArt = { kind: "none" }

export interface UseVideoArtParams {
	titleId: string
	// An episode's (or any one video's) own art; null for the title's.
	stableUuid: string | null
	slot: ArtSlotCode
	title: string
	// Within reach of the screen: stored art may be read. Nothing is read before.
	near: boolean
	// The element whose time on screen allows a download (after the dwell).
	element: Element | null
}

function candidatesFor({ titleId, stableUuid, slot }: Pick<UseVideoArtParams, "titleId" | "stableUuid" | "slot">) {
	return stableUuid === null ? videoLibrary.titleArt(titleId, slot) : videoLibrary.itemArt(titleId, stableUuid, slot)
}

// A card's artwork: what is stored (or the drive's cached thumbnail) as soon as the card is near; anything that needs
// a download only once the card has been on screen for the dwell, cancelled when it leaves or unmounts. The image's
// object URL lives as long as the card shows it.
export function useVideoArt({ titleId, stableUuid, slot, title, near, element }: UseVideoArtParams): VideoArt {
	const key = `${titleId}:${stableUuid ?? ""}:${slot}`
	const [state, setState] = useState<{ key: string; art: VideoArt }>({ key: "", art: NONE })

	useEffect(() => {
		if (!near || element === null) {
			return
		}

		const controller = new AbortController()
		// A function, so each check reads the signal anew across the awaits.
		const aborted = (): boolean => controller.signal.aborted
		let url: string | null = null
		let visible = false
		let dwell: ReturnType<typeof setTimeout> | null = null
		let unobserve: (() => void) | null = null

		const show = (result: ArtResult): void => {
			if (aborted()) {
				return
			}

			if (result.type !== "image") {
				setState({ key, art: { kind: "text" } })

				return
			}

			url = URL.createObjectURL(result.blob)
			setState({ key, art: { kind: "image", url, fit: result.fit, dominant: result.dominant } })
		}

		const loadFull = (request: ArtLoadRequest): void => {
			artLoader
				.load(request)
				.then(result => {
					if (result !== null) {
						show(result)
					}
				})
				.catch((e: unknown) => {
					log.warn("videos.art", "artwork load failed", e)
				})
		}

		// Observed until the card goes: before the dwell, leaving the screen resets it; once loading, it only moves the
		// request's rank (read whenever a slot frees). Leaving the near zone unmounts this and aborts the load.
		const waitForDwell = (request: ArtLoadRequest): void => {
			let started = false

			unobserve = visibleCardObserver().observe(element, inView => {
				visible = inView

				if (started) {
					return
				}

				if (!inView) {
					if (dwell !== null) {
						clearTimeout(dwell)
						dwell = null
					}

					return
				}

				dwell ??= setTimeout(() => {
					started = true
					loadFull(request)
				}, CARD_PROBE_DWELL_MS)
			})
		}

		const run = async (): Promise<void> => {
			const art: VideoArtCandidates | null = await candidatesFor({ titleId, stableUuid, slot })

			if (aborted()) {
				return
			}

			if (art === null) {
				setState({ key, art: { kind: "text" } })

				return
			}

			const request: ArtLoadRequest & { fileUuid: (file: number) => string | null } = {
				candidates: art.candidates,
				slot,
				title,
				resolveFile: art.resolveFile,
				fileUuid: art.fileUuid,
				signal: controller.signal,
				rank: () => (visible ? PROBE_RANK.visible : PROBE_RANK.nearViewport)
			}
			const local = await artLoader.loadLocal(request)

			if (local === null || aborted()) {
				return
			}

			if (local === "network") {
				waitForDwell(request)

				return
			}

			show(local)
		}

		run().catch((e: unknown) => {
			log.warn("videos.art", "artwork lookup failed", e)
		})

		return () => {
			controller.abort()

			if (dwell !== null) {
				clearTimeout(dwell)
			}

			unobserve?.()

			if (url !== null) {
				URL.revokeObjectURL(url)
				// The revoked image must not show if the card comes near again before its next load.
				setState(previous => (previous.key === key ? { key: "", art: NONE } : previous))
			}
		}
	}, [key, titleId, stableUuid, slot, title, near, element])

	return near && state.key === key ? state.art : NONE
}
