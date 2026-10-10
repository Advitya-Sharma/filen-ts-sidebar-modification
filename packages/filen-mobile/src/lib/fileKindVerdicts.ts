import { boundedVerdicts, isStoredVerdict, VERDICT_MAX_RECORDS, type ContentVerdict, type StoredVerdict } from "@filen/shared"

// What reads of a file's content already said about its kind (resolveFileKind's verdict): a `.ts` that is a
// transport stream, an `.mp4` holding audio alone. Written only by reads that happen anyway (a player that
// loaded the file, the `.ts` preview gate's one read) and kept only where the verdict changes what the file
// resolves to (recordContentVerdict in previewType.ts). Per file uuid, which a content change rotates; the
// newest VERDICT_MAX_RECORDS kept. Read synchronously from memory; boot attaches the kv
// (fileKindVerdictStorage.ts), which fills it once and keeps every later write. Components follow a file's
// verdict through useFileKindVerdict (hooks/useFileKindVerdict.ts).

export type FileKindVerdictStorage = {
	load: () => Promise<(readonly [string, unknown])[]>
	save: (uuid: string, value: StoredVerdict) => Promise<void>
	remove: (uuid: string) => Promise<void>
	onError: (message: string, error: unknown) => void
}

export type FileKindVerdictReader = {
	get: (uuid: string) => ContentVerdict | undefined
}

// The files a change touched, or null for any file (a reset).
type AnyListener = (changed: readonly string[] | null) => void

export type FileKindVerdictWatch = {
	subscribe: (listener: () => void) => () => void
	getSnapshot: () => FileKindVerdictReader
}

class FileKindVerdicts {
	// Most recently written last.
	private readonly verdicts = new Map<string, StoredVerdict>()
	// Per file, so a change re-renders only the rows showing that file.
	private readonly fileListeners = new Map<string, Set<() => void>>()
	private readonly anyListeners = new Set<AnyListener>()
	private storage: FileKindVerdictStorage | null = null
	// Set at logout and never cleared: the bundle reloads, and nothing may be written for the outgoing
	// account after the kv wipe.
	private locked = false
	// Settled once the kv load has: a forget before then has to reach the row it loads.
	private loaded = false
	// What forget() dropped before the kv load landed, per file: the loaded row holding it is dropped too.
	private readonly forgottenBeforeLoad = new Map<string, ContentVerdict>()
	// A new object per change, so a reader held by a component tells it the store changed.
	private reader: FileKindVerdictReader = this.newReader()

	private newReader(): FileKindVerdictReader {
		return {
			get: uuid => this.verdicts.get(uuid)?.v
		}
	}

	// `uuids` null: any file (a reset).
	private changed(uuids: readonly string[] | null): void {
		this.reader = this.newReader()

		if (uuids === null) {
			for (const listeners of this.fileListeners.values()) {
				for (const listener of listeners) {
					listener()
				}
			}
		} else {
			for (const uuid of uuids) {
				const listeners = this.fileListeners.get(uuid)

				if (listeners) {
					for (const listener of listeners) {
						listener()
					}
				}
			}
		}

		for (const listener of this.anyListeners) {
			listener(uuids)
		}
	}

	private persist(write: (storage: FileKindVerdictStorage) => Promise<void>): void {
		const storage = this.storage

		if (storage === null) {
			return
		}

		write(storage).catch((error: unknown) => {
			storage.onError("verdict not saved", error)
		})
	}

	// Fills memory from `target` once, merged with what this session already wrote, and keeps every later write.
	public attach(target: FileKindVerdictStorage): void {
		if (this.storage !== null || this.locked) {
			return
		}

		this.storage = target

		target
			.load()
			.then(entries => {
				if (this.storage !== target) {
					return
				}

				const stored: (readonly [string, StoredVerdict])[] = []
				const forgotten = this.forgottenBeforeLoad

				this.loaded = true

				for (const [uuid, value] of entries) {
					if (isStoredVerdict(value) && forgotten.get(uuid) !== value.v) {
						stored.push([uuid, value])
					} else if (!this.verdicts.has(uuid)) {
						// A row this session saved since stays.
						this.persist(kv => kv.remove(uuid))
					}
				}

				const { keep, drop } = boundedVerdicts(stored, [...this.verdicts.entries()], VERDICT_MAX_RECORDS)
				const before = new Map(this.verdicts)
				const changed: string[] = []

				this.verdicts.clear()

				for (const [uuid, value] of keep) {
					this.verdicts.set(uuid, value)

					if (before.get(uuid)?.v !== value.v) {
						changed.push(uuid)
					}

					before.delete(uuid)
				}

				for (const uuid of drop) {
					this.persist(kv => kv.remove(uuid))
				}

				// What this session wrote and the bound evicted.
				for (const uuid of before.keys()) {
					changed.push(uuid)
				}

				forgotten.clear()

				if (changed.length > 0) {
					this.changed(changed)
				}
			})
			.catch((error: unknown) => {
				this.loaded = true
				this.forgottenBeforeLoad.clear()

				target.onError("verdicts not read", error)
			})
	}

	public get(uuid: string): ContentVerdict | undefined {
		return this.verdicts.get(uuid)?.v
	}

	public record(uuid: string, verdict: ContentVerdict): void {
		if (this.locked || this.verdicts.get(uuid)?.v === verdict) {
			return
		}

		const value: StoredVerdict = {
			v: verdict,
			at: Date.now()
		}

		this.forgottenBeforeLoad.delete(uuid)

		this.verdicts.delete(uuid)
		this.verdicts.set(uuid, value)
		this.persist(kv => kv.save(uuid, value))

		const oldest = this.verdicts.size > VERDICT_MAX_RECORDS ? this.verdicts.keys().next().value : undefined

		if (oldest !== undefined) {
			this.verdicts.delete(oldest)
			this.persist(kv => kv.remove(oldest))
		}

		this.changed(oldest === undefined ? [uuid] : [uuid, oldest])
	}

	// Drops the file's verdict where it is `verdict`: a later read contradicted it with what the name already says.
	public forget(uuid: string, verdict: ContentVerdict): void {
		if (this.locked) {
			return
		}

		if (!this.loaded) {
			this.forgottenBeforeLoad.delete(uuid)
			this.forgottenBeforeLoad.set(uuid, verdict)

			const oldest = this.forgottenBeforeLoad.size > VERDICT_MAX_RECORDS ? this.forgottenBeforeLoad.keys().next().value : undefined

			if (oldest !== undefined) {
				this.forgottenBeforeLoad.delete(oldest)
			}
		}

		if (this.verdicts.get(uuid)?.v !== verdict) {
			return
		}

		this.verdicts.delete(uuid)
		this.persist(kv => kv.remove(uuid))
		this.changed([uuid])
	}

	public subscribeFile(uuid: string, listener: () => void): () => void {
		let listeners = this.fileListeners.get(uuid)

		if (!listeners) {
			listeners = new Set()

			this.fileListeners.set(uuid, listeners)
		}

		listeners.add(listener)

		return () => {
			const current = this.fileListeners.get(uuid)

			if (!current) {
				return
			}

			current.delete(listener)

			if (current.size === 0) {
				this.fileListeners.delete(uuid)
			}
		}
	}

	public subscribeAny(listener: AnyListener): () => void {
		this.anyListeners.add(listener)

		return () => {
			this.anyListeners.delete(listener)
		}
	}

	public getReader(): FileKindVerdictReader {
		return this.reader
	}

	// A useSyncExternalStore source whose reader is replaced only after a change `affects` says touches the
	// caller (a reset always does): a list filtered by its files' kinds filters again only for its own files.
	public watch(affects: (changed: readonly string[]) => boolean): FileKindVerdictWatch {
		let reader = this.reader

		return {
			subscribe: listener => {
				// A change while unsubscribed is not known to be another file's.
				if (reader !== this.reader) {
					reader = this.reader
				}

				return this.subscribeAny(changed => {
					if (changed === null || affects(changed)) {
						reader = this.reader

						listener()
					}
				})
			},
			getSnapshot: () => reader
		}
	}

	public size(): number {
		return this.verdicts.size
	}

	// Logout: the kv rows go with the kv wipe; this forgets memory, detaches the kv and takes no further writes.
	public clearForLogout(): void {
		this.locked = true
		this.storage = null
		this.verdicts.clear()
		this.forgottenBeforeLoad.clear()
		this.changed(null)
	}
}

const fileKindVerdicts = new FileKindVerdicts()

export { FileKindVerdicts }

export default fileKindVerdicts
