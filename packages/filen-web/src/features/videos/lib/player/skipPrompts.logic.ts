import {
	SKIP_BUTTON_HIDE_MS,
	defaultSkipMode,
	segmentAction,
	segmentAt,
	skipTarget,
	type SegmentKind,
	type SkipMode,
	type SkipSegment
} from "@filen/shared/videos"

// Skip Intro / Skip Credits as playback moves through a file's segments (chapter skip types, chapter names,
// an .edl). Each segment acts once: a prompt shows on entering it and hides after a while unless the
// controls are up; an auto-skip jumps past it. Coming back into a segment already met (a seek back) never
// prompts or skips again. An outro running to the end belongs to Up Next instead.

export interface SkipView {
	// The segment the Skip button is for, null when none shows.
	prompt: SkipSegment | null
	// Where to jump now (an auto-skip), null for nothing.
	jumpTo: number | null
}

const NONE: SkipView = { prompt: null, jumpTo: null }

function segmentId(segment: SkipSegment): string {
	return `${String(segment.start)}:${String(segment.end)}`
}

export class SkipPromptTracker {
	private readonly segments: readonly SkipSegment[]
	private readonly duration: number
	private readonly modes: (kind: SegmentKind) => SkipMode
	// Segments met before; never acted on again.
	private readonly met = new Set<string>()
	private promptId: string | null = null
	private promptSince = 0

	public constructor(params: { segments: readonly SkipSegment[]; duration: number; modes?: (kind: SegmentKind) => SkipMode }) {
		this.segments = params.segments
		this.duration = params.duration
		this.modes = params.modes ?? defaultSkipMode
	}

	// Called as the playhead moves (a few times a second at most). `controlsShown`: the prompt stays while
	// the controls are up.
	public update(params: { time: number; now: number; controlsShown: boolean }): SkipView {
		const { time, now, controlsShown } = params
		const segment = segmentAt(this.segments, time)

		if (segment === null) {
			this.promptId = null

			return NONE
		}

		const id = segmentId(segment)

		if (this.promptId === id) {
			if (controlsShown) {
				this.promptSince = now
			}

			return now - this.promptSince < SKIP_BUTTON_HIDE_MS && skipTarget(segment, time) !== null
				? { prompt: segment, jumpTo: null }
				: NONE
		}

		if (this.met.has(id)) {
			return NONE
		}

		this.met.add(id)

		const action = segmentAction(segment, this.modes(segment.kind), this.duration)

		if (action === "autoskip") {
			const target = skipTarget(segment, time)

			return target === null ? NONE : { prompt: null, jumpTo: target }
		}

		if (action === "prompt" && skipTarget(segment, time) !== null) {
			this.promptId = id
			this.promptSince = now

			return { prompt: segment, jumpTo: null }
		}

		return NONE
	}

	// The Skip button was pressed (or S): where to seek, null when too little of the segment is left.
	public skip(time: number): number | null {
		const segment = this.promptId === null ? null : segmentAt(this.segments, time)

		this.promptId = null

		return segment === null ? null : skipTarget(segment, time)
	}

	// The prompt was dismissed without skipping.
	public dismiss(): void {
		this.promptId = null
	}
}

// Where Up Next may start early: the start of an outro that runs to the end of the file.
export function creditsStart(
	segments: readonly SkipSegment[],
	duration: number,
	modes: (kind: SegmentKind) => SkipMode = defaultSkipMode
): number | null {
	for (const segment of segments) {
		if (segment.kind === "outro" && segmentAction(segment, modes(segment.kind), duration) === "upNext") {
			return segment.start
		}
	}

	return null
}
