import { rangeReads, type RangeReadBudget } from "@/lib/media/rangeReads"

// While media plays, or waits for data to play on, the background media features give up their range
// reads to it (rangeReads); a pause, the end, an emptied or failed element, or the returned dispose hands
// them back. Driven by the element's state changes only, never per timeupdate.
const HOLD_EVENTS = ["playing", "waiting"] as const
const RELEASE_EVENTS = ["pause", "ended", "emptied", "error"] as const

export function holdReadsWhilePlaying(media: HTMLMediaElement, budget: Pick<RangeReadBudget, "holdPlayback"> = rangeReads): () => void {
	let release: (() => void) | null = null
	const hold = (): void => {
		release ??= budget.holdPlayback()
	}
	const drop = (): void => {
		release?.()
		release = null
	}

	for (const type of HOLD_EVENTS) {
		media.addEventListener(type, hold)
	}

	for (const type of RELEASE_EVENTS) {
		media.addEventListener(type, drop)
	}

	if (!media.paused) {
		hold()
	}

	return () => {
		for (const type of HOLD_EVENTS) {
			media.removeEventListener(type, hold)
		}

		for (const type of RELEASE_EVENTS) {
			media.removeEventListener(type, drop)
		}

		drop()
	}
}
