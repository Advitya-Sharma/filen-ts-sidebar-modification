import { useEffect } from "react"
import { holdReadsWhilePlaying } from "@/lib/media/playbackReads"

// Playback first for range reads while `media` plays (holdReadsWhilePlaying), until it unmounts.
export function useHoldReadsWhilePlaying(media: HTMLMediaElement | null): void {
	useEffect(() => (media === null ? undefined : holdReadsWhilePlaying(media)), [media])
}
