import type { CueStore } from "@filen/shared/videos"

// Subtitles as the element's own text tracks (addTextTrack + VTTCue: no blob URL, no fetch). Normally the
// track is "hidden" and a DOM overlay draws the cues (its cuechange is the overlay's clock); where an overlay
// can't be seen (picture-in-picture, iPhone's native fullscreen) the browser shows the track itself.
// addTextTrack's tracks can't be removed from an element, so they are pooled: one is filled, the rest stay
// disabled and empty.

type CueConstructor = new (
	start: number,
	end: number,
	text: string
) => TextTrackCue & { line: number | "auto"; snapToLines: boolean; align: string }

function cueConstructor(): CueConstructor | null {
	const global = globalThis as { VTTCue?: CueConstructor }

	return global.VTTCue ?? null
}

function clearCues(track: TextTrack): void {
	const cues = track.cues

	if (cues === null) {
		return
	}

	for (let i = cues.length - 1; i >= 0; i--) {
		const cue = cues[i]

		if (cue !== undefined) {
			track.removeCue(cue)
		}
	}
}

// Empties a track and turns it off. The cues go first: a disabled track exposes no cue list, so cues left on it
// would pile up again every time it is reused.
function retire(track: TextTrack): void {
	if (track.mode === "disabled") {
		track.mode = "hidden"
	}

	clearCues(track)
	track.mode = "disabled"
}

// Numpad alignment hint to the cue's line: 7-9 top, 4-6 middle, otherwise the default bottom placement.
function placeCue(cue: InstanceType<CueConstructor>, hint: number): void {
	if (hint >= 7 && hint <= 9) {
		cue.snapToLines = true
		cue.line = 0
	} else if (hint >= 4 && hint <= 6) {
		cue.snapToLines = false
		cue.line = 50
	}

	const column = hint >= 1 && hint <= 9 ? (hint - 1) % 3 : 1

	cue.align = column === 0 ? "start" : column === 2 ? "end" : "center"
}

export class SubtitleTrackPool {
	private readonly video: HTMLVideoElement
	private readonly tracks: TextTrack[] = []
	private current: TextTrack | null = null
	private native = false
	private cueListener: (() => void) | null = null

	public constructor(video: HTMLVideoElement) {
		this.video = video
	}

	// Fills a track with `store`'s cues shifted by `offset` seconds and makes it the current one; null clears.
	public show(store: CueStore | null, offset: number): void {
		const previous = this.current

		if (previous !== null) {
			this.detachListener(previous)
			retire(previous)
			this.current = null
		}

		if (store === null || store.texts.length === 0) {
			return
		}

		const Cue = cueConstructor()

		if (Cue === null) {
			return
		}

		const track = this.tracks.find(entry => entry.mode === "disabled") ?? this.addTrack()

		// A disabled track's cue list reads null: it is enabled (hidden) before anything is added or removed.
		track.mode = "hidden"
		clearCues(track)

		for (let i = 0; i < store.texts.length; i++) {
			const start = Math.max(0, (store.starts[i] ?? 0) + offset)
			const end = Math.max(start, (store.ends[i] ?? 0) + offset)
			const cue = new Cue(start, end, store.texts[i] ?? "")
			const hint = store.hints?.[i] ?? 0

			cue.id = String(i)

			if (hint !== 0) {
				placeCue(cue, hint)
			}

			track.addCue(cue)
		}

		this.current = track
		track.mode = this.native ? "showing" : "hidden"
		this.attachListener(track)
	}

	// The browser draws the cues itself (picture-in-picture, iPhone fullscreen), or leaves them to the overlay.
	public setNative(native: boolean): void {
		this.native = native

		if (this.current !== null) {
			this.current.mode = native ? "showing" : "hidden"
		}
	}

	// Called whenever the set of showing cues changes, and when the current track changes.
	public onCueChange(listener: (() => void) | null): void {
		if (this.current !== null) {
			this.detachListener(this.current)
		}

		this.cueListener = listener

		if (this.current !== null) {
			this.attachListener(this.current)
		}
	}

	public dispose(): void {
		if (this.current !== null) {
			this.detachListener(this.current)
		}

		for (const track of this.tracks) {
			retire(track)
		}

		this.current = null
		this.cueListener = null
	}

	private addTrack(): TextTrack {
		const track = this.video.addTextTrack("subtitles")

		track.mode = "disabled"
		this.tracks.push(track)

		return track
	}

	private attachListener(track: TextTrack): void {
		if (this.cueListener !== null) {
			track.addEventListener("cuechange", this.cueListener)
		}
	}

	private detachListener(track: TextTrack): void {
		if (this.cueListener !== null) {
			track.removeEventListener("cuechange", this.cueListener)
		}
	}
}
