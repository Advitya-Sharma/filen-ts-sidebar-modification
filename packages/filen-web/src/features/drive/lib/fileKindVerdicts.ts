import { boundedVerdicts, VERDICT_MAX_RECORDS, type ContentVerdict, type StoredVerdict } from "@filen/shared"
import { createListenerSet } from "@/lib/listenerSet"
import { log } from "@/lib/log"

// What reads of a file's content already said about its kind (resolveFileKind's verdict): a `.ts` that is a
// transport stream, a `.webm` holding audio alone. Written only by reads that happen anyway (a preview's or
// thumbnail's probe), never by a read made to classify, and kept only where it changes what the file
// resolves to (contentVerdicts.ts). Per file uuid (a content change rotates it), the newest
// VERDICT_MAX_RECORDS kept. Read synchronously from memory; the page attaches kv at boot
// (fileKindVerdictStorage.ts), which fills it once and keeps every write. The SDK worker classifies with
// this module too and never attaches: it holds no verdicts. Components read a file's verdict through
// useFileKindVerdict (hooks/useFileKindVerdict.ts) so they follow it.

export interface FileKindVerdictStorage {
	load: () => Promise<(readonly [string, StoredVerdict])[]>
	save: (uuid: string, value: StoredVerdict) => Promise<void>
	remove: (uuid: string) => Promise<void>
}

// Most recently written last.
const verdicts = new Map<string, StoredVerdict>()
// Each change names the files it concerns; null for any (a load, a reset).
const changes = createListenerSet<readonly string[] | null>("file-kind-verdicts")
let storage: FileKindVerdictStorage | null = null
let epoch = 0

function changed(uuids: readonly string[] | null): void {
	changes.emit(uuids)
}

function persist(write: (storage: FileKindVerdictStorage) => Promise<void>): void {
	if (storage !== null) {
		write(storage).catch((error: unknown) => {
			log.warn("file-kind", "verdict not saved", error)
		})
	}
}

// Fills memory from `target` once, merged with what this session already wrote, and keeps every later write.
export function attachFileKindVerdictStorage(target: FileKindVerdictStorage): void {
	if (storage !== null) {
		return
	}

	storage = target

	const started = epoch

	target
		.load()
		.then(entries => {
			if (started !== epoch) {
				return
			}

			const { keep, drop } = boundedVerdicts(entries, [...verdicts.entries()], VERDICT_MAX_RECORDS)

			verdicts.clear()

			for (const [uuid, value] of keep) {
				verdicts.set(uuid, value)
			}

			for (const uuid of drop) {
				persist(kv => kv.remove(uuid))
			}

			if (keep.length > 0) {
				changed(null)
			}
		})
		.catch((error: unknown) => {
			log.warn("file-kind", "verdicts not read", error)
		})
}

export function fileKindVerdict(uuid: string): ContentVerdict | undefined {
	return verdicts.get(uuid)?.v
}

export function recordFileKindVerdict(uuid: string, verdict: ContentVerdict): void {
	if (verdicts.get(uuid)?.v === verdict) {
		return
	}

	const value: StoredVerdict = { v: verdict, at: Date.now() }

	verdicts.delete(uuid)
	verdicts.set(uuid, value)
	persist(kv => kv.save(uuid, value))

	const oldest = verdicts.size > VERDICT_MAX_RECORDS ? verdicts.keys().next().value : undefined

	if (oldest !== undefined) {
		verdicts.delete(oldest)
		persist(kv => kv.remove(oldest))
	}

	changed(oldest === undefined ? [uuid] : [uuid, oldest])
}

export function subscribeFileKindVerdicts(listener: (uuids: readonly string[] | null) => void): () => void {
	return changes.subscribe(listener)
}

// The files holding `verdict`, at most VERDICT_MAX_RECORDS.
export function filesWithVerdict(verdict: ContentVerdict): string[] {
	const uuids: string[] = []

	for (const [uuid, value] of verdicts) {
		if (value.v === verdict) {
			uuids.push(uuid)
		}
	}

	return uuids
}

// Logout: the kv rows go with the kv wipe; this forgets what memory holds. Storage stays attached.
export function resetFileKindVerdicts(): void {
	epoch++
	verdicts.clear()
	changed(null)
}
