import { useEffect } from "react"
import { createScreenWakeLock } from "@/lib/media/wakeLock"
import { playbackEnvironment } from "@/lib/media/playback/capabilities"

// Keeps the screen on while the stage's video plays (Chromium does that itself for a playing video).
export function useStageWakeLock(video: HTMLVideoElement | null, extra = false): void {
	useEffect(() => {
		if (video === null) {
			return
		}

		const browser = playbackEnvironment().browser
		const lock = createScreenWakeLock({ needed: browser !== "chrome" && browser !== "edge" })
		const element = video

		function update(): void {
			lock.set(extra || (!element.paused && !element.ended))
		}

		update()

		const events = ["play", "pause", "ended", "emptied"]

		for (const event of events) {
			element.addEventListener(event, update)
		}

		return () => {
			for (const event of events) {
				element.removeEventListener(event, update)
			}

			lock.dispose()
		}
	}, [video, extra])
}
