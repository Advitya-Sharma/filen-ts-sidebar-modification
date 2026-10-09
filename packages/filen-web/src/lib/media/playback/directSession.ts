import { directErrorAction, KEEPALIVE_IDLE_PAUSE_MS, RECOVERY_TIMEOUT_MS } from "@/lib/media/playback/directSession.logic"
import type { PlaybackSession, SessionAudio, SessionCallbacks } from "@/lib/media/playback/types"

// A natively playable file streamed by the service worker's range route straight into the element.

// The service worker's side of a direct stream.
export interface DirectSource {
	// Registers the file under `id` (again, after the worker lost it) and resolves its URL.
	register(id: string): Promise<string>
	// The element is done with `id`.
	release(id: string): void
	// Keeps the worker alive until the returned release is called.
	holdAlive(): () => void
	// Whether `url` answers a range request (a two-byte read): the stream is there, so an error the element
	// reported is the file's.
	reachable(url: string): Promise<boolean>
	// A registration made ahead of playback (prewarm), taken over instead of a new one.
	presetId?: string
}

interface AudioTrackLike {
	enabled: boolean
}

interface AudioTrackListLike {
	readonly length: number
	[index: number]: AudioTrackLike | undefined
}

// HTMLVideoElement's typings lack audioTracks; currentTime only makes this assignable from it.
interface ElementWithAudioTracks {
	currentTime: number
	audioTracks?: AudioTrackListLike
}

// Safari's own audio track switching on a direct stream; undefined elsewhere.
function audioTracksOf(video: HTMLVideoElement | null): AudioTrackListLike | undefined {
	const element: ElementWithAudioTracks | null = video

	return element?.audioTracks
}

export function openDirectSession(params: {
	source: DirectSource
	startTime: number
	audioIndex: number
	callbacks: SessionCallbacks
	now?: () => number
}): PlaybackSession {
	return new DirectSession(params)
}

class DirectSession implements PlaybackSession {
	public readonly path = "direct"
	private readonly source: DirectSource
	private readonly callbacks: SessionCallbacks
	private readonly now: () => number
	private readonly listeners: [string, () => void][] = []
	private video: HTMLVideoElement | null = null
	private id: string | null = null
	private audioIndex: number
	// Where the element goes once it has metadata: the start position, or the position a recovery saved.
	private pendingSeek: number | null
	private resumeAfterLoad = false
	private hadMetadata = false
	private reconnecting = false
	private lastRecoveryAt: number | null = null
	// What the user last asked for, which a reload's own pause does not change.
	private wantsPlay = false
	// The pause a reload of a playing element fires on its own, not the user's.
	private reloadPause = false
	private url: string | null = null
	private releaseAlive: (() => void) | null = null
	private idleTimer: ReturnType<typeof setTimeout> | undefined
	private recoveryTimer: ReturnType<typeof setTimeout> | undefined
	private disposed = false

	public constructor(params: {
		source: DirectSource
		startTime: number
		audioIndex: number
		callbacks: SessionCallbacks
		now?: () => number
	}) {
		this.source = params.source
		this.callbacks = params.callbacks
		this.now = params.now ?? (() => performance.now())
		this.audioIndex = params.audioIndex
		this.pendingSeek = params.startTime > 0 ? params.startTime : null
	}

	public async attach(video: HTMLVideoElement): Promise<void> {
		const id = this.source.presetId ?? crypto.randomUUID()
		let url: string

		try {
			url = await this.source.register(id)
		} catch (error) {
			if (!this.disposed) {
				this.callbacks.onFailure({ trigger: "stream-unavailable", detail: error instanceof Error ? error.message : String(error) })
			}

			return
		}

		if (this.disposed) {
			this.source.release(id)

			return
		}

		this.id = id
		this.url = url
		this.video = video
		// A remote player could reach neither the service worker's URL nor its decryption.
		video.disableRemotePlayback = true
		this.listen("loadedmetadata", () => {
			this.onMetadata()
		})
		this.listen("canplay", () => {
			if (this.reconnecting) {
				this.reconnecting = false
				clearTimeout(this.recoveryTimer)
				this.callbacks.onReconnecting(false)

				// Paused while it reconnected: the hold goes after the idle pause again.
				if (video.paused) {
					this.armIdleRelease()
				}
			}
		})
		this.listen("error", () => {
			this.onError()
		})
		this.listen("play", () => {
			this.wantsPlay = true
			this.holdAlive()
		})
		this.listen("pause", () => {
			if (this.reloadPause) {
				this.reloadPause = false

				return
			}

			this.wantsPlay = false
			this.armIdleRelease()
		})
		this.holdAlive()

		if (video.paused) {
			this.armIdleRelease()
		}

		video.src = url
	}

	public selectAudio(audio: SessionAudio): Promise<boolean> {
		const tracks = audioTracksOf(this.video)

		if (audio.index === this.audioIndex) {
			return Promise.resolve(true)
		}

		if (tracks === undefined || audio.index >= tracks.length) {
			return Promise.resolve(false)
		}

		this.audioIndex = audio.index
		this.applyAudioTrack()

		return Promise.resolve(true)
	}

	public dispose(): void {
		if (this.disposed) {
			return
		}

		this.disposed = true
		clearTimeout(this.idleTimer)
		clearTimeout(this.recoveryTimer)
		this.releaseAlive?.()
		this.releaseAlive = null

		const video = this.video

		for (const [type, listener] of this.listeners) {
			video?.removeEventListener(type, listener)
		}

		// Emptied before the registration goes, so no range request of this element can still reach it.
		if (video !== null) {
			video.removeAttribute("src")
			video.load()
		}

		if (this.id !== null) {
			this.source.release(this.id)
		}
	}

	private listen(type: string, listener: () => void): void {
		this.video?.addEventListener(type, listener)
		this.listeners.push([type, listener])
	}

	private holdAlive(): void {
		clearTimeout(this.idleTimer)
		this.idleTimer = undefined
		this.releaseAlive ??= this.source.holdAlive()
	}

	// A long pause lets the worker go; playing again re-holds it, and a stream it lost meanwhile recovers.
	private armIdleRelease(): void {
		clearTimeout(this.idleTimer)
		this.idleTimer = setTimeout(() => {
			this.releaseAlive?.()
			this.releaseAlive = null
		}, KEEPALIVE_IDLE_PAUSE_MS)
	}

	private onMetadata(): void {
		const video = this.video

		if (video === null) {
			return
		}

		this.hadMetadata = true
		this.applyAudioTrack()

		if (this.pendingSeek !== null) {
			video.currentTime = this.pendingSeek
			this.pendingSeek = null
		}

		// The play state asked for last, a pause during the reconnect included.
		if (this.resumeAfterLoad) {
			this.resumeAfterLoad = false

			if (this.wantsPlay) {
				void video.play().catch(() => undefined)
			}
		}
	}

	private applyAudioTrack(): void {
		const tracks = audioTracksOf(this.video)

		if (tracks === undefined || this.audioIndex < 0 || this.audioIndex >= tracks.length) {
			return
		}

		for (let index = 0; index < tracks.length; index++) {
			const track = tracks[index]

			if (track !== undefined) {
				track.enabled = index === this.audioIndex
			}
		}
	}

	private onError(): void {
		const video = this.video
		const error = video?.error

		if (video === null || error === null || error === undefined || this.disposed) {
			return
		}

		const action = directErrorAction({
			code: error.code,
			hadMetadata: this.hadMetadata,
			recovering: this.reconnecting,
			now: this.now(),
			lastRecoveryAt: this.lastRecoveryAt
		})
		const detail = `MediaError ${String(error.code)}${error.message === "" ? "" : `: ${error.message}`}`

		if (action === "recover") {
			void this.recover(video)
		} else if (action === "verify") {
			void this.verify(video, detail)
		} else if (action === "demote") {
			this.callbacks.onFailure({ trigger: "media-error", detail })
		} else if (action === "fail") {
			clearTimeout(this.recoveryTimer)
			this.callbacks.onFailure({ trigger: "network", detail })
		}
	}

	// A source refused before any metadata: the file when the stream answers, the stream otherwise.
	private async verify(video: HTMLVideoElement, detail: string): Promise<void> {
		const url = this.url
		const reachable = url !== null && (await this.source.reachable(url).catch(() => false))

		if (this.disposed) {
			return
		}

		if (reachable) {
			this.callbacks.onFailure({ trigger: "media-error", detail })
		} else if (this.lastRecoveryAt === null) {
			void this.recover(video)
		} else {
			this.callbacks.onFailure({ trigger: "network", detail })
		}
	}

	// The worker lost the registration (terminated while idle) or the stream broke: register the same id
	// again, reload, and go back to the saved position and play state.
	private async recover(video: HTMLVideoElement): Promise<void> {
		const id = this.id

		if (id === null) {
			return
		}

		this.lastRecoveryAt = this.now()
		this.reconnecting = true
		this.callbacks.onReconnecting(true)
		this.recoveryTimer = setTimeout(() => {
			if (this.reconnecting && !this.disposed) {
				this.callbacks.onFailure({ trigger: "network", detail: "the stream did not come back" })
			}
		}, RECOVERY_TIMEOUT_MS)

		const time = video.currentTime

		try {
			await this.source.register(id)
		} catch (error) {
			clearTimeout(this.recoveryTimer)

			if (!this.disposed) {
				this.callbacks.onFailure({ trigger: "network", detail: error instanceof Error ? error.message : String(error) })
			}

			return
		}

		if (this.disposed) {
			return
		}

		this.holdAlive()
		this.pendingSeek = time > 0 ? time : this.pendingSeek
		this.resumeAfterLoad = true
		this.reloadPause = !video.paused
		video.load()
	}
}
