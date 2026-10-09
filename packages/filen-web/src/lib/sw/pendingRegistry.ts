// The service worker's registry of resolved, ready-to-stream downloads, keyed by opaque token. Split
// out of sw.ts purely so this bookkeeping is unit-testable — a worker module cannot be imported under
// node. Deliberately generic: the entry type (decrypted SDK handles) never crosses this boundary.
//
// Retention is bounded rather than explicit: entries hold decrypted key material, no page-side signal
// says a download finished, and an entry must survive repeated GETs (Safari probes a Range, then
// re-fetches) — so neither a completion signal nor a GET can drive eviction. Past the ceiling the LEAST
// RECENTLY READ entry goes, except one whose stream is still in flight: dropping that 404s its next
// range request mid-playback. A preview whose element is done hands its entry back (delete).
//
// Entries live in buckets, each with its own ceiling, and only crowd out entries of their own bucket: a
// burst of short-lived registrations (thumbnails) can never evict a paused player's entry.
export const DEFAULT_BUCKET = "default"

export class PendingRegistry<T> {
	private readonly buckets = new Map<string, Map<string, T>>()
	private readonly limits = new Map<string, number>()
	private readonly bucketOf = new Map<string, string>()
	// Per-id in-flight stream count: one id can be read twice at once (Safari's range probe overlapping
	// the real fetch), so this counts rather than flags.
	private readonly streams = new Map<string, number>()
	private inFlight = 0

	// `max` is the default bucket's ceiling; `extraBuckets` names the others. An unknown bucket name
	// registers into the default one.
	public constructor(max: number, extraBuckets: Readonly<Record<string, number>> = {}) {
		this.limits.set(DEFAULT_BUCKET, max)

		for (const [name, limit] of Object.entries(extraBuckets)) {
			this.limits.set(name, limit)
		}
	}

	public set(id: string, entry: T, bucket: string = DEFAULT_BUCKET): void {
		const name = this.limits.has(bucket) ? bucket : DEFAULT_BUCKET
		const max = this.limits.get(name) ?? 0
		let entries = this.buckets.get(name)

		if (entries === undefined) {
			entries = new Map()
			this.buckets.set(name, entries)
		}

		// Re-registering an existing id (a preview revisit reuses its id) must count as fresh too, or the
		// entry keeps its old position and the next unrelated registration could evict it before its GET.
		this.delete(id)
		entries.set(id, entry)
		this.bucketOf.set(id, name)

		if (entries.size <= max) {
			return
		}

		for (const key of entries.keys()) {
			// The entry just registered is never its own eviction victim; nor is one still being streamed
			// (both are, by definition, the ones in use).
			if (key === id || this.streams.has(key)) {
				continue
			}

			this.delete(key)

			return
		}
	}

	// Reading an entry refreshes its recency (Map preserves insertion order, so re-inserting moves it to
	// the end) — an entry still being fetched must never be the next eviction candidate.
	public get(id: string): T | undefined {
		const entries = this.entriesOf(id)
		const entry = entries?.get(id)

		if (entries === undefined || entry === undefined) {
			return undefined
		}

		entries.delete(id)
		entries.set(id, entry)

		return entry
	}

	// Drops one entry by id. A stream still reading from it keeps its own reference and is counted until it
	// ends, as before.
	public delete(id: string): void {
		this.entriesOf(id)?.delete(id)
		this.bucketOf.delete(id)
	}

	public beginStream(id: string): void {
		this.streams.set(id, (this.streams.get(id) ?? 0) + 1)
		this.inFlight++
	}

	public endStream(id: string): void {
		const count = this.streams.get(id) ?? 0

		if (count > 1) {
			this.streams.set(id, count - 1)
		} else {
			this.streams.delete(id)
		}

		this.inFlight = Math.max(0, this.inFlight - 1)
	}

	// Total streams currently being pumped, across every id — the "an update must not truncate a running
	// save" gate. In-flight counts deliberately survive clear(): those pumps still hold their own SDK
	// handles and settle on their own.
	public get activeStreams(): number {
		return this.inFlight
	}

	public get size(): number {
		return this.bucketOf.size
	}

	public clear(): void {
		this.buckets.clear()
		this.bucketOf.clear()
	}

	private entriesOf(id: string): Map<string, T> | undefined {
		const bucket = this.bucketOf.get(id)

		return bucket === undefined ? undefined : this.buckets.get(bucket)
	}
}
