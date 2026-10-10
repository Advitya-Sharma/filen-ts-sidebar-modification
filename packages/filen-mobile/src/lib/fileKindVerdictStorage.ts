import sqlite from "@/lib/sqlite"
import { forEachKvRowByPrefix, KV_SMALL_ROW_PAGE_SIZE } from "@/lib/kvScan"
import { deserialize } from "@/lib/serializer"
import logger from "@/lib/logger"
import { VERDICT_KV_PREFIX } from "@filen/shared"
import fileKindVerdicts from "@/lib/fileKindVerdicts"

// The file-kind verdicts' kv rows, one per file uuid under a versioned prefix. A row that does not decode is
// handed to the store as such, which drops it.

// Boot, once the kv is open: the verdicts are in memory by the first listing's paint, or follow it as one change.
export function connectFileKindVerdictStorage(): void {
	fileKindVerdicts.attach({
		load: async () => {
			const entries: (readonly [string, unknown])[] = []
			const db = await sqlite.openDb()

			await forEachKvRowByPrefix(
				db,
				VERDICT_KV_PREFIX,
				(key, value) => {
					let decoded: unknown = null

					try {
						decoded = deserialize(value)
					} catch {
						// Dropped by the store as malformed.
					}

					entries.push([key.slice(VERDICT_KV_PREFIX.length), decoded])
				},
				KV_SMALL_ROW_PAGE_SIZE
			)

			return entries
		},
		save: async (uuid, value) => {
			await sqlite.kvAsync.set(`${VERDICT_KV_PREFIX}${uuid}`, value)
		},
		remove: async uuid => {
			await sqlite.kvAsync.remove(`${VERDICT_KV_PREFIX}${uuid}`)
		},
		onError: (message, error) => {
			logger.warn("fileKind", message, { error })
		}
	})
}
