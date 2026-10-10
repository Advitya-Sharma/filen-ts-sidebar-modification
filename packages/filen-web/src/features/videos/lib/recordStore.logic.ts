import type { Type } from "arktype"
import { InFlight, planSizeCapEviction } from "@filen/shared"
import type { kvDelete, kvDeleteIfAt, kvEntriesJson, kvGetJson, kvKeys, kvSetJson } from "@/lib/storage/adapter"
import { log } from "@/lib/log"
import { readSignOutPending } from "@/lib/signOutPending"

// Small per-file records in kv under one prefix (probe records, artwork metadata, frame hints): read one
// at a time and only when asked for, kept in a bounded memory cache, written behind, and held to a record
// count by evicting the least recently written. Nothing is hydrated up front: a library of 20k files
// would otherwise parse megabytes of rows nobody looks at.

export interface StoredRecord {
	// When the record was written; eviction goes oldest first.
	at: number
}

export interface RecordStoreKv {
	kvGetJson: typeof kvGetJson
	kvSetJson: typeof kvSetJson
	kvKeys: typeof kvKeys
	kvEntriesJson: typeof kvEntriesJson
	kvDelete: typeof kvDelete
	// Evicts many rows at once, each only while it still holds the `at` planned; without it, each is read first.
	kvDeleteIfAt?: typeof kvDeleteIfAt
}

export interface RecordStoreOptions<T extends StoredRecord> {
	prefix: string
	schema: Type<T>
	maxRecords: number
	// Records kept in memory, most recently used last.
	memoryRecords: number
	// Writes between two checks of the record count (the first write of a session checks too).
	writesPerSweep: number
	scope: string
	kv: RecordStoreKv
	now?: () => number
}

// The ids to drop so that at most `target` of `records` remain, oldest write first; none while there are
// at most `max`. Evicting to below the cap leaves room, so the next check is many writes away.
export function planCountEviction(records: readonly (readonly [string, number])[], max: number, target: number): string[] {
	if (records.length <= max) {
		return []
	}

	return planSizeCapEviction(
		records.map(([id, at]) => ({ id, size: 1, timestamp: at })),
		target
	)
}

// How long a record known to be absent stays known: another tab may write it meanwhile.
export const ABSENT_TTL_MS = 30_000

// kv round trips a sweep keeps in flight, so evicting thousands of rows never floods the storage worker.
const SWEEP_CONCURRENCY = 8
// Rows one conditional eviction drops in one storage transaction.
const SWEEP_BATCH = 500

// Every store of the page, so a sign-out holds them all at once.
const stores = new Set<{ hold: () => void }>()

// A sign-out started, here or in another tab: no store writes to kv again (the page reloads once it is done), so
// nothing lands after the wipe.
export function holdRecordStores(): void {
	for (const store of stores) {
		store.hold()
	}
}

export class RecordStore<T extends StoredRecord> {
	private readonly options: RecordStoreOptions<T>
	private readonly kv: RecordStoreKv
	private readonly now: () => number
	private readonly memory = new Map<string, T>()
	// id → when it was read as absent.
	private readonly absent = new Map<string, number>()
	private readonly reads = new InFlight<string, T | null>()
	private writesSinceSweep = 0
	private sweptThisSession = false
	private sweeping = false
	// Bumped by reset, so a read settling after logout writes nothing back.
	private epoch = 0
	private held = false

	public constructor(options: RecordStoreOptions<T>) {
		this.options = options
		this.kv = options.kv
		this.now = options.now ?? Date.now
		stores.add(this)
	}

	// Writes go to memory only from now on; reset doesn't lift it.
	public hold(): void {
		this.held = true
	}

	// Whether writes still reach kv: false once a sign-out started, here or in another tab.
	public writable(): boolean {
		if (!this.held && readSignOutPending()) {
			this.held = true
		}

		return !this.held
	}

	private key(id: string): string {
		return `${this.options.prefix}${id}`
	}

	private bound<V>(map: Map<string, V>): void {
		while (map.size > this.options.memoryRecords) {
			const oldest = map.keys().next()

			if (oldest.done === true) {
				break
			}

			map.delete(oldest.value)
		}
	}

	private remember(id: string, record: T | null): void {
		this.memory.delete(id)
		this.absent.delete(id)

		if (record === null) {
			this.absent.set(id, this.now())
			this.bound(this.absent)
		} else {
			this.memory.set(id, record)
			this.bound(this.memory)
		}
	}

	// The record in memory: undefined when unknown (or absent too long ago to trust), null when recently
	// read as absent.
	public peek(id: string): T | null | undefined {
		const record = this.memory.get(id)

		if (record !== undefined) {
			this.memory.delete(id)
			this.memory.set(id, record)

			return record
		}

		const absentAt = this.absent.get(id)

		return absentAt !== undefined && this.now() - absentAt <= ABSENT_TTL_MS ? null : undefined
	}

	public async get(id: string): Promise<T | null> {
		const known = this.peek(id)

		return known === undefined ? this.read(id) : known
	}

	// The record from memory, or else from kv even when it was recently read as absent: for the moment
	// before paying for what the record would save, which another tab may have done meanwhile.
	public async refresh(id: string): Promise<T | null> {
		return this.memory.get(id) ?? this.read(id)
	}

	private read(id: string): Promise<T | null> {
		const epoch = this.epoch

		return this.reads.coalesce(id, async () => {
			const record = await this.kv.kvGetJson(this.key(id), this.options.schema).catch((error: unknown) => {
				log.warn(this.options.scope, "record read failed", error)

				return null
			})

			if (epoch !== this.epoch) {
				return record
			}

			// A write that landed while the read was in flight is newer.
			const written = this.memory.get(id)

			if (written !== undefined) {
				return written
			}

			this.remember(id, record)

			return record
		})
	}

	public put(id: string, record: T): void {
		this.remember(id, record)

		if (!this.writable()) {
			return
		}

		void this.kv.kvSetJson(this.key(id), record).catch((error: unknown) => {
			log.warn(this.options.scope, "record write failed", error)
		})

		this.writesSinceSweep++

		if (!this.sweptThisSession || this.writesSinceSweep >= this.options.writesPerSweep) {
			this.sweptThisSession = true
			void this.sweep()
		}
	}

	public delete(id: string): void {
		this.remember(id, null)

		if (!this.writable()) {
			return
		}

		void this.kv.kvDelete(this.key(id)).catch((error: unknown) => {
			log.warn(this.options.scope, "record delete failed", error)
		})
	}

	// Deletes `key` only if it still holds the row the plan saw (`at`, or an invalid row for null): a
	// row rewritten since, by this tab or another, stays.
	private async deleteIfUnchanged(key: string, at: number | null): Promise<void> {
		const current = await this.kv.kvGetJson(key, this.options.schema)

		if (this.held || (current !== null && current.at !== at)) {
			return
		}

		this.memory.delete(key.slice(this.options.prefix.length))
		await this.kv.kvDelete(key)
	}

	// Holds the prefix to its record count. Keys alone are listed first: only an overfull prefix pays for
	// reading every row. Rows that no longer validate go too.
	public async sweep(): Promise<void> {
		if (this.sweeping || !this.writable()) {
			return
		}

		this.sweeping = true
		this.writesSinceSweep = 0

		try {
			const { prefix, maxRecords, schema } = this.options
			const keys = await this.kv.kvKeys(prefix)

			if (keys.length <= maxRecords) {
				return
			}

			const rows = await this.kv.kvEntriesJson(prefix, schema)
			const planned = new Map(rows.map(([key, record]) => [key, record.at]))
			const evicted = planCountEviction(
				rows.map(([key, record]) => [key, record.at] as const),
				maxRecords,
				Math.floor(maxRecords * 0.9)
			)
			// Rows that no longer validate are read again before they go (rare: a schema change).
			const evict: [string, number | null][] = keys.filter(key => !planned.has(key)).map(key => [key, null] as [string, null])
			const deleteIfAt = this.kv.kvDeleteIfAt

			if (deleteIfAt === undefined) {
				evict.push(...evicted.map(key => [key, planned.get(key) ?? null] as [string, number | null]))
			} else {
				for (let start = 0; start < evicted.length && !this.held; start += SWEEP_BATCH) {
					const batch = evicted.slice(start, start + SWEEP_BATCH).flatMap(key => {
						const at = planned.get(key)

						return at === undefined ? [] : [[key, at] as const]
					})

					for (const key of await deleteIfAt(batch)) {
						this.memory.delete(key.slice(prefix.length))
					}
				}
			}

			let next = 0
			const drain = async (): Promise<void> => {
				for (let item = evict[next++]; item !== undefined; item = evict[next++]) {
					await this.deleteIfUnchanged(item[0], item[1]).catch(() => undefined)
				}
			}

			await Promise.all(Array.from({ length: Math.min(SWEEP_CONCURRENCY, evict.length) }, drain))
		} catch (error) {
			log.warn(this.options.scope, "record sweep failed", error)
		} finally {
			this.sweeping = false
		}
	}

	// Logout: memory only; the kv rows go with the kv wipe itself, and hold keeps any late write out of it.
	public reset(): void {
		this.epoch++
		this.memory.clear()
		this.absent.clear()
		this.writesSinceSweep = 0
		this.sweptThisSession = false
	}
}
