import { onAbort } from "@/lib/abort"

export interface SharedWorkContext {
	// Aborts once every caller has given up (a caller without a signal never does).
	signal: AbortSignal
	// The most urgent rank among the callers still interested; 0 for one without a rank.
	rank: () => number
}

export interface SharedWorkOptions {
	signal?: AbortSignal | undefined
	rank?: (() => number) | undefined
}

interface Interest {
	controller: AbortController
	count: number
	pinned: boolean
	ranks: Set<() => number>
}

interface Entry<V> {
	promise: Promise<V>
	interest: Interest
}

function firstInLine(): number {
	return 0
}

// One run of a job per key, joined by every caller that asks for the same key while it runs, so two
// cards showing one title share one download and one decode.
export class SharedWork<V> {
	private readonly entries = new Map<string, Entry<V>>()

	public run(key: string, options: SharedWorkOptions, work: (context: SharedWorkContext) => Promise<V>): Promise<V> {
		const existing = this.entries.get(key)

		if (existing !== undefined && !existing.interest.controller.signal.aborted) {
			this.join(existing.interest, options)

			return existing.promise
		}

		const interest: Interest = { controller: new AbortController(), count: 0, pinned: false, ranks: new Set() }
		const rank = (): number => {
			let best = interest.pinned && interest.ranks.size === 0 ? 0 : Number.POSITIVE_INFINITY

			for (const each of interest.ranks) {
				best = Math.min(best, each())
			}

			return best
		}

		this.join(interest, options)

		const entry: Entry<V> = { promise: work({ signal: interest.controller.signal, rank }), interest }

		this.entries.set(key, entry)
		void entry.promise.then(
			() => {
				this.forget(key, entry)
			},
			() => {
				this.forget(key, entry)
			}
		)

		return entry.promise
	}

	private forget(key: string, entry: Entry<V>): void {
		if (this.entries.get(key) === entry) {
			this.entries.delete(key)
		}
	}

	private join(interest: Interest, { signal, rank }: SharedWorkOptions): void {
		const ranked = rank ?? firstInLine

		if (signal === undefined) {
			interest.pinned = true
			interest.ranks.add(ranked)

			return
		}

		if (signal.aborted) {
			if (!interest.pinned && interest.count === 0) {
				interest.controller.abort()
			}

			return
		}

		interest.count++
		interest.ranks.add(ranked)
		onAbort(signal, () => {
			interest.count--
			interest.ranks.delete(ranked)

			if (!interest.pinned && interest.count <= 0) {
				interest.controller.abort()
			}
		})
	}
}
