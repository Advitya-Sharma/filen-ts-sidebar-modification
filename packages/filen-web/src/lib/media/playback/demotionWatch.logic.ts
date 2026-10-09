import type { PlaybackTrigger } from "@/lib/media/playback/types"

// The watchdogs that tell a rung that loads but does not play from slow loading, as a pure state machine
// fed the element's events and a periodic sample. Times are milliseconds on one clock.

// A direct stream with no metadata by then is not going to play.
export const METADATA_TIMEOUT_MS = 6_000
// canplay (while playing) without a presented frame by then: the decoder took the data and shows nothing.
export const FIRST_FRAME_TIMEOUT_MS = 4_000
// Data buffered ahead, yet below HAVE_FUTURE_DATA for this long: decoding, not loading, is stuck.
export const DECODE_STALL_MS = 3_000
export const DECODE_STALL_BUFFERED_SECONDS = 2
// Media played for this long with audio expected and none decoded.
export const SILENT_AUDIO_SECONDS = 3
// How often the element is sampled while playing.
export const WATCH_SAMPLE_MS = 500

const HAVE_FUTURE_DATA = 3
// A larger jump between two samples is a seek, not playback.
const MAX_PLAYED_STEP_SECONDS = 1.5

export interface WatchSample {
	now: number
	paused: boolean
	seeking: boolean
	readyState: number
	currentTime: number
	// Seconds buffered ahead of currentTime in the range holding it.
	bufferedAhead: number
	// Bytes of audio decoded so far where the engine says (webkitAudioDecodedByteCount), else null.
	audioDecodedBytes: number | null
}

export interface WatchConfig {
	// Only a direct stream waits for metadata on a timer; a remux session's own pipeline reports.
	metadataTimeout: boolean
	// Frames are observable (requestVideoFrameCallback).
	frameCallback: boolean
	expectAudio: boolean
}

export class DemotionWatchLogic {
	private readonly config: WatchConfig
	private metadataDeadline: number | null = null
	private frameDeadline: number | null = null
	private sawFrame = false
	private canPlay = false
	private stallSince: number | null = null
	private playedSeconds = 0
	private lastTime: number | null = null
	private audioChecked = false
	private suspended = false
	private fired = false
	private hidden = false
	private hiddenAt: number | null = null
	private playingNow = false
	// The tab was hidden while a frame or stall timer could run: what they say then is not trusted.
	private ranHidden = false

	public constructor(config: WatchConfig) {
		this.config = config
	}

	// The source was set (or set again after a reconnect).
	public start(now: number): void {
		this.suspended = false
		this.metadataDeadline = this.config.metadataTimeout ? now + METADATA_TIMEOUT_MS : null
		this.frameDeadline = null
		this.stallSince = null
		this.lastTime = null
	}

	// While a lost stream is re-established nothing counts.
	public suspend(): void {
		this.suspended = true
		this.metadataDeadline = null
		this.frameDeadline = null
		this.stallSince = null
	}

	// A deadline runs that needs sampling even while paused.
	public get needsClock(): boolean {
		return this.metadataDeadline !== null && !this.fired
	}

	public loadedMetadata(): void {
		this.metadataDeadline = null
	}

	public canPlayEvent(now: number, paused: boolean): void {
		this.canPlay = true
		this.armFrame(now, paused)
	}

	public playing(now: number): void {
		this.playingNow = true
		this.armFrame(now, false)
	}

	public paused(): void {
		this.playingNow = false
	}

	// A hidden tab presents no frames (requestVideoFrameCallback stops) and may starve the decoder, so the
	// frame and stall timers stop while hidden and start over once visible.
	public visibility(hidden: boolean, now: number): void {
		if (hidden === this.hidden) {
			return
		}

		this.hidden = hidden

		if (hidden) {
			this.ranHidden ||= this.playingNow || this.frameDeadline !== null
			this.hiddenAt = now
			this.frameDeadline = null
			this.stallSince = null

			return
		}

		if (this.hiddenAt !== null && this.metadataDeadline !== null) {
			this.metadataDeadline += now - this.hiddenAt
		}

		this.hiddenAt = null
		this.lastTime = null
		this.armFrame(now, !this.playingNow)
	}

	// Whether a trigger fired from this watch can be trusted as a verdict on the browser.
	public get trusted(): boolean {
		return !this.ranHidden
	}

	public frame(): void {
		this.sawFrame = true
		this.frameDeadline = null
	}

	// The trigger due at this sample, once; null while all is well.
	public sample(sample: WatchSample): PlaybackTrigger | null {
		if (this.fired || this.suspended || this.hidden) {
			return null
		}

		const trigger = this.check(sample)

		if (trigger !== null) {
			this.fired = true
		}

		return trigger
	}

	private armFrame(now: number, paused: boolean): void {
		if (
			this.config.frameCallback &&
			this.canPlay &&
			!this.sawFrame &&
			!paused &&
			this.frameDeadline === null &&
			!this.suspended &&
			!this.hidden
		) {
			this.frameDeadline = now + FIRST_FRAME_TIMEOUT_MS
		}
	}

	private check(sample: WatchSample): PlaybackTrigger | null {
		const { now } = sample

		if (this.metadataDeadline !== null && now >= this.metadataDeadline) {
			return "no-metadata"
		}

		if (this.frameDeadline !== null && now >= this.frameDeadline && !sample.paused) {
			return "no-first-frame"
		}

		const playing = !sample.paused && !sample.seeking

		if (playing && sample.readyState < HAVE_FUTURE_DATA && sample.bufferedAhead > DECODE_STALL_BUFFERED_SECONDS) {
			this.stallSince ??= now

			if (now - this.stallSince >= DECODE_STALL_MS) {
				return "decode-stall"
			}
		} else {
			this.stallSince = null
		}

		if (playing && this.lastTime !== null) {
			const step = sample.currentTime - this.lastTime

			if (step > 0 && step <= MAX_PLAYED_STEP_SECONDS) {
				this.playedSeconds += step
			}
		}

		this.lastTime = sample.currentTime

		if (
			this.config.expectAudio &&
			!this.audioChecked &&
			sample.audioDecodedBytes !== null &&
			this.playedSeconds >= SILENT_AUDIO_SECONDS
		) {
			this.audioChecked = true

			if (sample.audioDecodedBytes === 0) {
				return "silent-audio"
			}
		}

		return null
	}
}

// Seconds buffered ahead of `time` in the range holding it (with a small tolerance for a range starting
// just after it).
export function bufferedAheadOf(ranges: readonly (readonly [number, number])[], time: number): number {
	for (const [start, end] of ranges) {
		if (start <= time + 0.1 && end > time) {
			return end - time
		}
	}

	return 0
}
