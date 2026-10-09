// Relative seeks from held or repeated keys. The first press seeks at once; presses following within the
// window collapse into one trailing seek to where they add up, so a stream reads one position rather than
// every one passed over (each seek abandons the previous one's reads).

export const SEEK_COALESCE_MS = 200

export interface SeekCoalescerDeps<Timer> {
	// The element's position when no seek is pending.
	current: () => number
	duration: () => number
	seek: (seconds: number) => void
	// Where the pending seek will land, null once it has been applied (for a scrubber preview).
	onPending?: (seconds: number | null) => void
	now: () => number
	setTimer: (callback: () => void, ms: number) => Timer
	clearTimer: (timer: Timer) => void
}

export function clampSeek(seconds: number, duration: number): number {
	if (!Number.isFinite(seconds)) {
		return 0
	}

	return Math.max(0, duration > 0 ? Math.min(duration, seconds) : seconds)
}

export class SeekCoalescer<Timer = ReturnType<typeof setTimeout>> {
	private readonly deps: SeekCoalescerDeps<Timer>
	private lastPressAt = -Infinity
	private pending: number | null = null
	private timer: Timer | null = null

	public constructor(deps: SeekCoalescerDeps<Timer>) {
		this.deps = deps
	}

	// Where the playhead is headed: the pending target, else the element's position.
	public target(): number {
		return this.pending ?? this.deps.current()
	}

	public seekBy(delta: number): void {
		this.seekTo(this.target() + delta)
	}

	public seekTo(seconds: number): void {
		const now = this.deps.now()
		const target = clampSeek(seconds, this.deps.duration())
		const inWindow = now - this.lastPressAt < SEEK_COALESCE_MS

		this.lastPressAt = now

		if (!inWindow && this.pending === null) {
			this.deps.seek(target)

			return
		}

		this.pending = target
		this.deps.onPending?.(target)

		if (this.timer !== null) {
			this.deps.clearTimer(this.timer)
		}

		this.timer = this.deps.setTimer(() => {
			this.timer = null
			this.flush()
		}, SEEK_COALESCE_MS)
	}

	// Applies a pending seek now.
	public flush(): void {
		if (this.timer !== null) {
			this.deps.clearTimer(this.timer)
			this.timer = null
		}

		const pending = this.pending

		if (pending === null) {
			return
		}

		this.pending = null
		this.deps.seek(pending)
		this.deps.onPending?.(null)
	}

	// Drops a pending seek (the item changed).
	public cancel(): void {
		if (this.timer !== null) {
			this.deps.clearTimer(this.timer)
			this.timer = null
		}

		if (this.pending !== null) {
			this.pending = null
			this.deps.onPending?.(null)
		}

		this.lastPressAt = -Infinity
	}
}
