import { openRemuxSession, type MseFailure, type RemuxSession, type RemuxSessionOptions } from "@/lib/media/mseController"
import { isIdrUnreachable, isSourceReadError, SourceReadError, type RemuxTextCue } from "@/lib/media/remuxTypes"
import type { ProbeIndex } from "@/lib/media/probe/types"
import type {
	PlaybackFailure,
	PlaybackSession,
	PlaybackTrigger,
	SessionAudio,
	SessionCallbacks,
	SessionPath
} from "@/lib/media/playback/types"

// A file the remux worker turns into fragmented MP4 for MediaSource, read from the SDK worker over a range
// port. A port whose serving worker dies is noticed through that worker's life lock (lib/media/rangePort),
// so a session fails at once rather than at a read timeout, and a dead remux worker through its own.

// The SDK worker's side of a remux session.
export interface RemuxSource {
	// A range port serving the file's decrypted bytes.
	openPort(): Promise<MessagePort>
	// Stops serving the file: every read in flight aborts.
	close(): void
	// The file was opened in the remux worker ahead of playback under this key (prewarm).
	warmKey?: string
}

// The SDK worker answers a port request at once; one that has not by then is gone (a dead worker never
// answers a call at all).
export const OPEN_PORT_TIMEOUT_MS = 10_000

const FAILURE_TRIGGERS: Record<MseFailure, PlaybackTrigger> = {
	keyframeIntervalTooLarge: "keyframe-interval",
	stalled: "decode-stall",
	remux: "worker-error",
	media: "media-error",
	source: "network",
	openGopStart: "open-gop-start"
}

function errorDetail(error: unknown): string {
	return error instanceof Error ? error.message : String(error)
}

// What a failure to open or attach says about the rung.
export function remuxOpenFailure(error: unknown): PlaybackFailure {
	const detail = errorDetail(error)

	if (isIdrUnreachable(error)) {
		return { trigger: "open-gop-start", detail }
	}

	if (isSourceReadError(error)) {
		return { trigger: "network", detail }
	}

	// addSourceBuffer refusing the output's codecs.
	if (error instanceof DOMException && error.name === "NotSupportedError") {
		return { trigger: "media-error", detail }
	}

	// The worker died, its pipeline threw, or no MediaSource exists.
	return { trigger: "worker-error", detail }
}

export interface RemuxSessionParams {
	path: Exclude<SessionPath, "direct">
	source: RemuxSource
	size: number
	index: ProbeIndex | null
	startTime: number
	// Null: video alone.
	audio: SessionAudio | null
	// Video starts only at IDR pictures (Firefox with HEVC).
	idrOnlyStart: boolean
	callbacks: SessionCallbacks
	// Embedded text subtitle cues as the remux reads them.
	onTextCues?: ((cues: RemuxTextCue[]) => void) | undefined
	open?: (options: RemuxSessionOptions) => Promise<RemuxSession>
}

export function openRemuxPlaybackSession(params: RemuxSessionParams): PlaybackSession {
	return new RemuxPlaybackSession(params)
}

class RemuxPlaybackSession implements PlaybackSession {
	public readonly path: Exclude<SessionPath, "direct">
	private readonly params: RemuxSessionParams
	private session: RemuxSession | null = null
	private audio: SessionAudio | null
	private failed = false
	private disposed = false
	private portOpen = false

	public constructor(params: RemuxSessionParams) {
		this.params = params
		this.path = params.path
		this.audio = params.audio
	}

	public async attach(video: HTMLVideoElement): Promise<void> {
		const { source, size, index, startTime, idrOnlyStart } = this.params
		const open = this.params.open ?? openRemuxSession

		try {
			let timer: ReturnType<typeof setTimeout> | undefined
			const port = await Promise.race([
				source.openPort(),
				new Promise<never>((_resolve, reject) => {
					timer = setTimeout(() => {
						reject(new SourceReadError(new Error("the file's byte source did not answer")))
					}, OPEN_PORT_TIMEOUT_MS)
				})
			]).finally(() => {
				clearTimeout(timer)
			})

			this.portOpen = true

			if (this.isDisposed()) {
				port.close()
				this.closeSource()

				return
			}

			const session = await open({
				port,
				size,
				startTime,
				idrOnlyStart,
				...(index === null ? {} : { index }),
				audio: this.audio === null ? { trackId: -1, mode: "none" } : { trackId: this.audio.trackId, mode: this.audio.mode },
				onFailure: (reason, detail) => {
					this.fail({ trigger: FAILURE_TRIGGERS[reason], detail })
				},
				...(this.params.onTextCues === undefined ? {} : { onTextCues: this.params.onTextCues }),
				...(source.warmKey === undefined ? {} : { warmKey: source.warmKey })
			})

			if (this.isDisposed()) {
				session.dispose()

				return
			}

			this.session = session
			await session.attach(video)
		} catch (error) {
			this.fail(remuxOpenFailure(error))
		}
	}

	public async selectAudio(audio: SessionAudio): Promise<boolean> {
		const session = this.session

		if (session === null || audio.mode === "none" || this.audio === null) {
			return false
		}

		await session.selectAudio(audio.trackId, audio.mode)
		this.audio = audio

		return true
	}

	// A method rather than a field read: it changes across awaits.
	private isDisposed(): boolean {
		return this.disposed
	}

	public dispose(): void {
		if (this.disposed) {
			return
		}

		this.disposed = true
		this.session?.dispose()
		this.session = null
		this.closeSource()
	}

	private closeSource(): void {
		if (this.portOpen) {
			this.portOpen = false
			this.params.source.close()
		}
	}

	private fail(failure: PlaybackFailure): void {
		if (this.failed || this.disposed) {
			return
		}

		this.failed = true
		this.params.callbacks.onFailure(failure)
	}
}
