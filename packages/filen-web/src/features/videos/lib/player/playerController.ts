import type { File as SdkFile } from "@filen/sdk-rs"
import { CueBuilder, itemKey, seriesKey, type CueStore, type ItemRec, type LadderSubtitle } from "@filen/shared/videos"
import type { RemuxTextCue } from "@/lib/media/remuxTypes"
import { openPlayback, type Playback, type PlaybackSnapshot } from "@/lib/media/playback/openPlayback"
import type { VideoProbe } from "@/lib/media/probe/types"
import { createListenerSet } from "@/lib/listenerSet"
import { log } from "@/lib/log"
import { videoLibrary } from "@/features/videos/lib/index"
import { probeService } from "@/features/videos/lib/probeService"
import { PROBE_RANK } from "@/features/videos/lib/probeService.logic"
import { watchState, type WatchItem } from "@/features/videos/lib/watchState"
import { captureFrameFromVideo } from "@/features/videos/lib/frames"
import { storeCapturedFrame } from "@/features/videos/lib/artwork"
import { planStart, type StartRequest } from "@/features/videos/lib/player/resume.logic"
import { stackProgress, type WatchItemTarget } from "@/features/videos/lib/player/watchTarget.logic"
import { defaultAudioIndex } from "@/features/videos/lib/player/audioChoice.logic"

// Plays watch targets on the stage's one <video> element: finds the files, probes them (cached), opens the
// playback ladder at the resume point, follows stacked parts, and reports progress to the watch state.
// Its snapshot changes on discrete events only (an item opening, a status change); position never goes
// through it.

export type PlayerPhase = "idle" | "loading" | "ready" | "missing" | "probeFailed" | "unsupported" | "failed"

export interface PlayerSnapshot {
	phase: PlayerPhase
	target: WatchItemTarget | null
	part: number
	probe: VideoProbe | null
	playback: PlaybackSnapshot | null
	audioIndex: number
	partDurations: readonly (number | null)[]
	// Started from a resume point (the "Resumed at" pill), with the position it started at.
	resumedAt: number | null
	// play() was refused (no user activation): the caller offers a play button.
	autoplayBlocked: boolean
	// Bumped each time an item or part starts, so the UI can reset per-item state.
	epoch: number
	// Bumped whenever embedded subtitle cues arrived (embeddedCues).
	embeddedVersion: number
}

const IDLE: PlayerSnapshot = {
	phase: "idle",
	target: null,
	part: 0,
	probe: null,
	playback: null,
	audioIndex: -1,
	partDurations: [],
	resumedAt: null,
	autoplayBlocked: false,
	epoch: 0,
	embeddedVersion: 0
}

// Audio codecs every browser decodes without transcoding: what a default track choice prefers.
// The player waits this long at most for the watch state's first read (another device's newer resume point)
// before it starts from what it has; a flush in another tab can hold the lock much longer.
export const WATCH_STATE_WAIT_MS = 1500

// A pause this early or this close to the end leaves no still worth keeping.
const CAPTURE_MIN_POSITION_SECONDS = 5
const CAPTURE_END_FRACTION = 0.95

function userLanguages(): readonly string[] {
	return typeof navigator === "undefined" ? [] : navigator.languages
}

export interface OpenRequest {
	target: WatchItemTarget
	start: StartRequest
	// Started by the user (a click) rather than by autoplay: play() is expected to be allowed.
	autoplay: boolean
	// An audio track picked before opening (the detail page); the automatic choice when absent or out of range.
	audioIndex?: number
}

export class PlayerController {
	private snapshot: PlayerSnapshot = IDLE
	private readonly changes = createListenerSet("videos.player")
	private video: HTMLVideoElement | null = null
	private playback: Playback | null = null
	private unsubscribePlayback: (() => void) | null = null
	private files: SdkFile[] = []
	private generation = 0
	private readonly elementListeners: [string, () => void][] = []
	private onEnded: (() => void) | null = null
	private captureDone = false
	// What this item last reported to the watch state, ended with a stop whatever happens next.
	private lastReported: { item: WatchItem; pos: number; dur: number } | null = null
	// Where the element last really was (part, seconds), for a retry; null before the first position.
	private lastPosition: { key: string; part: number; time: number } | null = null
	// The item was already reported as finished (Up Next advanced past its credits).
	private finished = false
	// Embedded text cues of the open file, by track number, as the remux has read them.
	private embedded = new Map<number, RemuxTextCue[]>()
	private embeddedStores = new Map<number, { version: number; store: CueStore }>()
	private embeddedKeys = new Set<string>()
	// What the ladder is told about subtitles: an embedded text track needs the remux rung.
	private subtitleKind: LadderSubtitle = "none"
	// The audio track picked before opening, for the next part opened only.
	private pickedAudio: number | null = null

	public readonly subscribe = (listener: () => void): (() => void) => this.changes.subscribe(listener)

	public readonly getSnapshot = (): PlayerSnapshot => this.snapshot

	// The whole item (every part) ended; the caller decides what comes next.
	public setOnEnded(listener: (() => void) | null): void {
		this.onEnded = listener
	}

	public attach(video: HTMLVideoElement | null): void {
		if (this.video === video) {
			return
		}

		this.detachElement()
		this.video = video

		if (video === null) {
			return
		}

		const listen = (type: string, listener: () => void): void => {
			video.addEventListener(type, listener)
			this.elementListeners.push([type, listener])
		}

		listen("timeupdate", () => {
			this.report(false)
		})
		listen("play", () => {
			this.report(true)
			this.captureDone = false

			if (this.snapshot.autoplayBlocked) {
				this.update({ autoplayBlocked: false })
			}
		})
		listen("pause", () => {
			this.report(true)
			this.captureStill()
		})
		listen("ended", () => {
			this.handleEnded()
		})
	}

	public async open(request: OpenRequest, partOverride?: { part: number; position: number }): Promise<void> {
		const video = this.video

		if (video === null) {
			return
		}

		const generation = ++this.generation

		this.stopCurrent()

		const { target } = request
		const keepTarget = this.snapshot.target?.key === target.key

		this.pickedAudio = request.audioIndex ?? null

		this.embedded = new Map()
		this.embeddedStores = new Map()
		this.embeddedKeys = new Set()
		this.update({
			...IDLE,
			phase: "loading",
			target,
			part: partOverride?.part ?? 0,
			epoch: this.snapshot.epoch + 1,
			embeddedVersion: this.snapshot.embeddedVersion + 1,
			partDurations: keepTarget ? this.snapshot.partDurations : target.parts.map(() => null)
		})

		// The resume point comes from the watch state: local journal plus the newest save from other devices, waited
		// for briefly. A part switch or a retry starts at a known position and doesn't wait.
		const [files] = await Promise.all([
			videoLibrary.playableFiles(target.parts.map(part => part.file.stableUuid)),
			partOverride === undefined ? this.awaitWatchState() : Promise.resolve()
		])

		if (generation !== this.generation) {
			return
		}

		const ordered = target.parts.map(part => files.find(file => file.stableUUID === part.file.stableUuid))

		if (ordered.some(file => file === undefined)) {
			this.update({ phase: "missing" })

			return
		}

		this.files = ordered.filter(file => file !== undefined)

		const partDurations = await this.partDurations(generation)

		if (generation !== this.generation) {
			return
		}

		const first = this.files[0]

		if (first === undefined) {
			this.update({ phase: "missing" })

			return
		}

		const plan =
			partOverride === undefined
				? planStart({
						request: request.start,
						rec: this.record(target.key),
						file: { uuid: first.uuid, size: Number(first.size) },
						partDurations
					})
				: { ...partOverride, resumed: false }

		await this.openPart(
			generation,
			target,
			plan.part,
			plan.position,
			request.autoplay,
			plan.resumed ? plan.position : null,
			partDurations
		)
	}

	// Plays the same item again from where it last really was (after a failure, a long stall, a failed probe), else
	// from its resume point. The element itself is no witness: a failed or restarting element reads 0.
	public async retry(): Promise<void> {
		const target = this.snapshot.target

		if (target === null || this.video === null) {
			return
		}

		const last = this.lastPosition?.key === target.key ? this.lastPosition : null

		await (last === null
			? this.open({ target, start: "resume", autoplay: true })
			: this.open({ target, start: "resume", autoplay: true }, { part: last.part, position: last.time }))
	}

	// An embedded text track was chosen (or not): it needs the remux rung, which a direct stream (an MKV played as
	// it is) restarts on, at the same position.
	public setSubtitleKind(kind: LadderSubtitle): void {
		if (kind === this.subtitleKind) {
			return
		}

		this.subtitleKind = kind

		if (kind === "embedded-text" && this.snapshot.playback?.decision.path === "direct") {
			void this.retry()
		}
	}

	// The cues of embedded track `track` read so far; null before any. `version` is the snapshot's embeddedVersion
	// the caller rendered with, so a memoizing caller asks again when it changes.
	public embeddedCues(track: number, version: number): CueStore | null {
		const cues = this.embedded.get(track)

		if (cues === undefined || cues.length === 0) {
			return null
		}

		const cached = this.embeddedStores.get(track)

		if (cached?.version === version) {
			return cached.store
		}

		const builder = new CueBuilder()

		for (const cue of cues) {
			builder.add(cue.start, cue.end, cue.text, cue.hint)
		}

		const store = builder.finish()

		this.embeddedStores.set(track, { version, store })

		return store
	}

	// The autoplay-blocked card's Play: a user gesture, so play() is allowed now.
	public play(): void {
		const video = this.video

		if (video !== null) {
			video.play().catch((error: unknown) => {
				log.warn("videos.player", "play refused", error)
			})
		}
	}

	// Up Next moved on past the credits: the item counts as played to its end.
	public finish(): void {
		const item = this.watchItem()
		const position = this.currentProgress()

		if (item === null || position === null || this.finished) {
			return
		}

		this.finished = true
		void watchState.reportStop({ item, pos: position.pos, dur: position.dur, ended: true })
	}

	public async selectAudio(index: number): Promise<void> {
		const playback = this.playback
		const target = this.snapshot.target

		if (playback === null || target === null || index === this.snapshot.audioIndex) {
			return
		}

		this.update({ audioIndex: index })
		await playback.selectAudio(index)

		const lang = this.snapshot.probe?.audio[index]?.lang ?? null

		if (target.seriesDirUuid !== null && lang !== null) {
			void watchState.setSeriesLanguages({ seriesDirUuid: target.seriesDirUuid, audio: lang })
		}
	}

	// Ends playback of the current item (reported as stopped) and leaves the element empty.
	public stop(): void {
		this.generation++
		this.stopCurrent()
		this.update({ ...IDLE, epoch: this.snapshot.epoch + 1, embeddedVersion: this.snapshot.embeddedVersion + 1 })
	}

	public dispose(): void {
		this.stop()
		this.detachElement()
		this.video = null
		this.onEnded = null
		this.changes.clear()
	}

	// The watch item being played, in watch-state terms.
	public watchItem(): WatchItem | null {
		const target = this.snapshot.target
		const first = this.files[0]

		if (target === null || first === undefined) {
			return null
		}

		return {
			stableUuid: target.key,
			uuid: first.uuid,
			size: Number(first.size),
			itemClass: target.itemClass,
			...(target.parts.length > 1 ? { part: this.snapshot.part } : {})
		}
	}

	private async awaitWatchState(): Promise<void> {
		const wait: { timer: ReturnType<typeof setTimeout> | null } = { timer: null }

		await Promise.race([
			watchState.open().catch((error: unknown) => {
				log.warn("videos.player", "watch state not opened", error)
			}),
			new Promise<void>(resolve => {
				wait.timer = setTimeout(resolve, WATCH_STATE_WAIT_MS)
			})
		])

		if (wait.timer !== null) {
			clearTimeout(wait.timer)
		}
	}

	// The element's position as the item's progress, when the element is really playing this item; else what was
	// last reported.
	private currentProgress(): { pos: number; dur: number } | null {
		const video = this.video

		if (video !== null && this.playing() && video.readyState >= video.HAVE_METADATA && Number.isFinite(video.duration)) {
			const progress = stackProgress(this.snapshot.partDurations, this.snapshot.part, video.currentTime, video.duration)

			return { pos: progress.pos, dur: progress.dur }
		}

		return this.lastReported === null ? null : { pos: this.lastReported.pos, dur: this.lastReported.dur }
	}

	// The current session plays (not opening, restarting, reconnecting or failed).
	private playing(): boolean {
		return this.snapshot.phase === "ready" && this.playback?.getSnapshot().status === "ready"
	}

	private record(key: string): ItemRec | undefined {
		return watchState.getSnapshot().doc.items[itemKey(key)]
	}

	// Durations of a stack's parts (the card tier answers them; usually cached already). One part needs none.
	private async partDurations(generation: number): Promise<(number | null)[]> {
		if (this.files.length <= 1) {
			return [null]
		}

		const records = await Promise.all(
			this.files.map(file => probeService.request({ uuid: file.uuid, file }, "card", { rank: () => PROBE_RANK.explicit }))
		)
		const durations = records.map(record => record?.probe?.durationSec ?? null)

		if (generation === this.generation) {
			this.update({ partDurations: durations })
		}

		return durations
	}

	private async openPart(
		generation: number,
		target: WatchItemTarget,
		part: number,
		position: number,
		autoplay: boolean,
		resumedAt: number | null,
		partDurations: readonly (number | null)[]
	): Promise<void> {
		const video = this.video
		const file = this.files[part]

		if (video === null || file === undefined) {
			return
		}

		const record = await probeService.request({ uuid: file.uuid, file }, "full", { rank: () => PROBE_RANK.explicit })

		if (generation !== this.generation) {
			return
		}

		const probe = record?.probe ?? null

		if (probe === null) {
			this.update({ phase: "probeFailed", part })

			return
		}

		const series = target.seriesDirUuid === null ? undefined : watchState.getSnapshot().doc.series[seriesKey(target.seriesDirUuid)]
		const picked = this.pickedAudio

		this.pickedAudio = null

		const audioIndex =
			picked !== null && picked < probe.audio.length
				? picked
				: defaultAudioIndex({ probe, rememberedLang: series?.al?.v ?? null, userLanguages: userLanguages() })

		this.update({ probe, part, audioIndex, resumedAt, partDurations, autoplayBlocked: false })

		let playback: Playback

		try {
			playback = await openPlayback(
				video,
				file,
				probe,
				{ audioIndex, subtitle: this.subtitleKind },
				{
					name: target.parts[part]?.file.name ?? "",
					startTime: position,
					autoplay: false,
					onTextCues: cues => {
						if (generation === this.generation) {
							this.addTextCues(cues)
						}
					}
				}
			)
		} catch (error) {
			log.warn("videos.player", "opening playback failed", error)

			if (generation === this.generation) {
				this.update({ phase: "failed" })
			}

			return
		}

		if (generation !== this.generation) {
			playback.dispose()

			return
		}

		this.playback = playback
		this.unsubscribePlayback = playback.subscribe(() => {
			this.syncPlayback()
		})
		this.syncPlayback()

		if (autoplay) {
			this.playWhenReady(generation)
		}
	}

	// play() once the element has metadata; a refusal (no user activation) is reported, never retried.
	private playWhenReady(generation: number): void {
		const video = this.video

		if (video === null) {
			return
		}

		const start = (): void => {
			if (generation !== this.generation) {
				return
			}

			video.play().then(
				() => {
					if (generation === this.generation && this.snapshot.autoplayBlocked) {
						this.update({ autoplayBlocked: false })
					}
				},
				(error: unknown) => {
					if (generation === this.generation && error instanceof DOMException && error.name === "NotAllowedError") {
						this.update({ autoplayBlocked: true })
					}
				}
			)
		}

		if (video.readyState >= video.HAVE_METADATA) {
			start()
		} else {
			video.addEventListener("loadedmetadata", start, { once: true })
		}
	}

	// A restarted remux (an audio change, a demotion) reads the same blocks again: cues are kept once.
	private addTextCues(cues: readonly RemuxTextCue[]): void {
		let added = false

		for (const cue of cues) {
			const key = `${String(cue.track)}:${String(Math.round(cue.start * 1000))}`

			if (this.embeddedKeys.has(key)) {
				continue
			}

			this.embeddedKeys.add(key)
			added = true

			const list = this.embedded.get(cue.track)

			if (list === undefined) {
				this.embedded.set(cue.track, [cue])
			} else {
				list.push(cue)
			}
		}

		if (added) {
			this.update({ embeddedVersion: this.snapshot.embeddedVersion + 1 })
		}
	}

	private syncPlayback(): void {
		const playback = this.playback

		if (playback === null) {
			return
		}

		const state = playback.getSnapshot()
		const phase: PlayerPhase =
			state.status === "unsupported"
				? "unsupported"
				: state.status === "failed"
					? "failed"
					: state.status === "opening"
						? "loading"
						: "ready"

		this.update({ playback: state, phase, audioIndex: state.selection.audioIndex })
	}

	// Never reports a position the element doesn't really hold: while opening or restarting it reads 0.
	private report(stateChanged: boolean): void {
		const video = this.video
		const item = this.watchItem()
		const target = this.snapshot.target

		if (
			video === null ||
			item === null ||
			target === null ||
			this.finished ||
			!this.playing() ||
			video.readyState < video.HAVE_METADATA ||
			video.seeking
		) {
			return
		}

		this.lastPosition = { key: target.key, part: this.snapshot.part, time: video.currentTime }

		if (!stateChanged && video.paused) {
			return
		}

		const partDuration = Number.isFinite(video.duration) ? video.duration : 0
		const progress = stackProgress(this.snapshot.partDurations, this.snapshot.part, video.currentTime, partDuration)

		this.lastReported = { item, pos: progress.pos, dur: progress.dur }
		watchState.reportProgress({ item, pos: progress.pos, dur: progress.dur, playing: !video.paused && !video.ended })
	}

	private handleEnded(): void {
		const target = this.snapshot.target
		const item = this.watchItem()
		const video = this.video

		if (target === null || item === null || video === null) {
			return
		}

		const nextPart = this.snapshot.part + 1

		if (nextPart < target.parts.length) {
			void this.open({ target, start: "resume", autoplay: true }, { part: nextPart, position: 0 })

			return
		}

		const progress = stackProgress(this.snapshot.partDurations, this.snapshot.part, video.currentTime, video.duration)

		if (!this.finished) {
			this.finished = true
			void watchState.reportStop({ item, pos: progress.pos, dur: progress.dur, ended: true })
		}

		this.lastReported = null
		this.onEnded?.()
	}

	// The frame on screen as the item's still (zero reads), once per pause.
	private captureStill(): void {
		const video = this.video
		const first = this.files[0]

		if (video === null || first === undefined || this.captureDone || video.ended || this.snapshot.phase !== "ready") {
			return
		}

		const duration = video.duration

		if (
			video.currentTime < CAPTURE_MIN_POSITION_SECONDS ||
			(Number.isFinite(duration) && video.currentTime > duration * CAPTURE_END_FRACTION)
		) {
			return
		}

		this.captureDone = true
		void captureFrameFromVideo(video, "s")
			.then(frame => (frame === null ? undefined : storeCapturedFrame(first.uuid, "s", frame)))
			.catch((error: unknown) => {
				log.warn("videos.player", "still capture failed", error)
			})
	}

	// Ends the item in the watch state whatever the state of the session (a stop at the last real position, or a
	// drop when it never had one), keeps a still, and lets go of the stream.
	private stopCurrent(): void {
		const item = this.watchItem() ?? this.lastReported?.item ?? null
		const progress = this.currentProgress()

		if (item !== null && !this.finished) {
			if (progress !== null && (this.lastReported !== null || this.playing())) {
				this.captureStill()
				void watchState.reportStop({ item, pos: progress.pos, dur: progress.dur })
			} else {
				watchState.dropPlayback({ stableUuid: item.stableUuid })
			}
		}

		this.lastReported = null
		this.finished = false
		this.unsubscribePlayback?.()
		this.unsubscribePlayback = null
		this.playback?.dispose()
		this.playback = null
	}

	private detachElement(): void {
		const video = this.video

		if (video !== null) {
			for (const [type, listener] of this.elementListeners) {
				video.removeEventListener(type, listener)
			}
		}

		this.elementListeners.length = 0
	}

	private update(patch: Partial<PlayerSnapshot>): void {
		this.snapshot = { ...this.snapshot, ...patch }
		this.changes.emit()
	}
}
