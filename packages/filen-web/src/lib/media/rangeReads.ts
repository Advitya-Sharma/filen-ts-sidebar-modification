import { RankedGate } from "@/features/drive/lib/rankedGate"

// One cap on the ranged downloads the background media features keep in flight together (video probes,
// artwork, generated frames, audio metadata), so a scrolled grid never queues more than a connection
// serves well. Each job holds one permit for its reads, which run one after another. Playback is never
// gated: while it runs, background work gets only what playback leaves over.
export const RANGE_READ_LIMIT = 6
export const BACKGROUND_READS_DURING_PLAYBACK = 2

// Lets the visible work mounting in the same pass claim free permits before work further down the queue.
const RANGE_READ_HOLD_MS = 50

export class RangeReadBudget {
	private readonly gate: RankedGate
	private readonly limit: number
	private readonly playbackLimit: number
	private playbacks = 0
	private readonly idleWaiters = new Set<() => void>()

	public constructor(limit = RANGE_READ_LIMIT, playbackLimit = BACKGROUND_READS_DURING_PLAYBACK, holdMs = RANGE_READ_HOLD_MS) {
		this.limit = limit
		this.playbackLimit = playbackLimit
		this.gate = new RankedGate(limit, holdMs)
	}

	// Runs `read` in a permit; `rank` orders the waiters as RankedGate does (0 and below first).
	public run<T>(rank: () => number, read: () => Promise<T>): Promise<T> {
		return this.gate.withPermit(rank, read)
	}

	public acquire(rank: () => number): Promise<void> {
		return this.gate.acquire(rank)
	}

	public release(): void {
		this.gate.release()
	}

	public get playing(): boolean {
		return this.playbacks > 0
	}

	// Held by a playback session for its whole length; the returned release is idempotent.
	public holdPlayback(): () => void {
		let released = false

		this.playbacks++

		if (this.playbacks === 1) {
			this.gate.setLimit(this.playbackLimit)
		}

		return () => {
			if (released) {
				return
			}

			released = true
			this.playbacks--

			if (this.playbacks === 0) {
				this.gate.setLimit(this.limit)

				for (const wake of [...this.idleWaiters]) {
					wake()
				}
			}
		}
	}

	// Resolves once no playback runs; rejects with the signal's reason if it aborts first.
	public whenPlaybackIdle(signal?: AbortSignal): Promise<void> {
		if (this.playbacks === 0) {
			return Promise.resolve()
		}

		return new Promise((resolve, reject) => {
			const settle = (): void => {
				this.idleWaiters.delete(settle)
				signal?.removeEventListener("abort", onAbort)
				resolve()
			}
			const onAbort = (): void => {
				this.idleWaiters.delete(settle)
				reject(signal?.reason instanceof Error ? signal.reason : new DOMException("aborted", "AbortError"))
			}

			if (signal?.aborted === true) {
				onAbort()

				return
			}

			this.idleWaiters.add(settle)
			signal?.addEventListener("abort", onAbort, { once: true })
		})
	}
}

export const rangeReads = new RangeReadBudget()
