import { bufferedAheadOf, DemotionWatchLogic, WATCH_SAMPLE_MS, type WatchConfig } from "@/lib/media/playback/demotionWatch.logic"
import type { PlaybackTrigger } from "@/lib/media/playback/types"

// Engines that count decoded audio (Chromium, WebKit); elsewhere silent audio goes undetected. The
// typings lack the counter; readyState only makes this assignable from an HTMLVideoElement.
interface AudioCountingElement {
	readyState: number
	webkitAudioDecodedByteCount?: number
}

export interface DemotionWatch {
	// The source is back after a reconnect: the watchdogs start over.
	restart(): void
	suspend(): void
	dispose(): void
}

function rangesOf(ranges: TimeRanges): [number, number][] {
	const out: [number, number][] = []

	for (let index = 0; index < ranges.length; index++) {
		out.push([ranges.start(index), ranges.end(index)])
	}

	return out
}

// Watches `video` for the signs that its current rung cannot play it (DemotionWatchLogic); calls
// `onTrigger` once. Samples only while the element plays, so a paused player costs nothing.
// `onTrigger`'s `trusted` is false when the tab was hidden while the watchdogs could run. `onFirstFrame`
// is called once the element presents its first frame (requestVideoFrameCallback; without it, the first
// "playing").
export function watchForDemotion(
	video: HTMLVideoElement,
	config: WatchConfig,
	onTrigger: (trigger: PlaybackTrigger, trusted: boolean) => void,
	onFirstFrame: () => void = () => undefined
): DemotionWatch {
	const logic = new DemotionWatchLogic(config)
	const listeners: [string, () => void][] = []
	let timer: ReturnType<typeof setInterval> | undefined
	let frameRequest: number | null = null
	let disposed = false
	const counting: AudioCountingElement = video
	const sample = () => {
		if (disposed) {
			return
		}

		const trigger = logic.sample({
			now: performance.now(),
			paused: video.paused,
			seeking: video.seeking,
			readyState: video.readyState,
			currentTime: video.currentTime,
			bufferedAhead: bufferedAheadOf(rangesOf(video.buffered), video.currentTime),
			audioDecodedBytes: typeof counting.webkitAudioDecodedByteCount === "number" ? counting.webkitAudioDecodedByteCount : null
		})

		if (trigger !== null) {
			stopSampling()
			onTrigger(trigger, logic.trusted)
		}
	}
	// Sampling runs while the element plays, or while a deadline is open.
	const syncSampling = () => {
		const wanted = !disposed && (!video.paused || logic.needsClock)

		if (wanted && timer === undefined) {
			timer = setInterval(sample, WATCH_SAMPLE_MS)
		} else if (!wanted) {
			stopSampling()
		}
	}
	const stopSampling = () => {
		clearInterval(timer)
		timer = undefined
	}
	const listen = (type: string, listener: () => void) => {
		video.addEventListener(type, listener)
		listeners.push([type, listener])
	}

	logic.start(performance.now())
	logic.visibility(document.hidden, performance.now())

	let sawFrame = false
	const firstFrame = () => {
		if (!sawFrame && !disposed) {
			sawFrame = true
			onFirstFrame()
		}
	}

	if (config.frameCallback) {
		frameRequest = video.requestVideoFrameCallback(() => {
			frameRequest = null
			logic.frame()
			firstFrame()
		})
	}

	const onVisibility = () => {
		logic.visibility(document.hidden, performance.now())
		syncSampling()
	}

	document.addEventListener("visibilitychange", onVisibility)

	listen("loadedmetadata", () => {
		logic.loadedMetadata()
		syncSampling()
	})
	listen("canplay", () => {
		logic.canPlayEvent(performance.now(), video.paused)
	})
	listen("playing", () => {
		logic.playing(performance.now())
		syncSampling()

		if (!config.frameCallback) {
			firstFrame()
		}
	})
	listen("pause", () => {
		logic.paused()
		syncSampling()
	})
	listen("ended", syncSampling)
	syncSampling()

	return {
		restart: () => {
			logic.start(performance.now())

			if (video.readyState >= video.HAVE_METADATA) {
				logic.loadedMetadata()
			}

			syncSampling()
		},
		suspend: () => {
			logic.suspend()
			syncSampling()
		},
		dispose: () => {
			disposed = true
			stopSampling()
			document.removeEventListener("visibilitychange", onVisibility)

			if (frameRequest !== null) {
				video.cancelVideoFrameCallback(frameRequest)
			}

			for (const [type, listener] of listeners) {
				video.removeEventListener(type, listener)
			}
		}
	}
}
