// One recursive listing per directory at a time, shared by every feature that walks it (Photos, the Videos
// index): a second caller joins the walk under way, and a settled result is reused briefly. Only while no drive
// event has arrived since the walk began, and the socket hasn't dropped or reconnected: a caller asking after an
// event must see what it changed, and a walk already under way may have read the tree before it. A caller can also
// refuse any walk begun before a point of its own (mark).

export const WALK_SETTLED_REUSE_MS = 10_000

interface Entry<T> {
	promise: Promise<T>
	// Walks started before this one, plus one.
	tick: number
}

export class WalkCoordinator<T> {
	private readonly entries = new Map<string, Entry<T>>()
	private readonly run: (dirUuid: string) => Promise<T>
	private started = 0

	public constructor(run: (dirUuid: string) => Promise<T>) {
		this.run = run
	}

	// A point in time: a walk joined with `startedAfter` set to it began after it.
	public mark(): number {
		return this.started
	}

	public walk(dirUuid: string, startedAfter = -1): Promise<T> {
		const existing = this.entries.get(dirUuid)

		if (existing !== undefined && existing.tick > startedAfter) {
			return existing.promise
		}

		this.started++

		const entry: Entry<T> = {
			promise: this.run(dirUuid),
			tick: this.started
		}

		this.entries.set(dirUuid, entry)

		// A settled result stays joinable for a short while; a failure doesn't.
		void entry.promise.then(
			() => {
				setTimeout(() => {
					if (this.entries.get(dirUuid) === entry) {
						this.entries.delete(dirUuid)
					}
				}, WALK_SETTLED_REUSE_MS)
			},
			() => {
				if (this.entries.get(dirUuid) === entry) {
					this.entries.delete(dirUuid)
				}
			}
		)

		return entry.promise
	}

	// Every drive event, socket drop and reconnect, before anything else handles it; and logout: walks begun before
	// can no longer be joined.
	public invalidate(): void {
		this.entries.clear()
	}
}
