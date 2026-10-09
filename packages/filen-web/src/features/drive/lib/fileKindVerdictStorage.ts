import { type } from "arktype"
import { kvDelete, kvEntriesJson, kvSetJson } from "@/lib/storage/adapter"
import { attachFileKindVerdictStorage } from "@/features/drive/lib/fileKindVerdicts"
import { VERDICT_KV_PREFIX } from "@/features/drive/lib/fileKindVerdicts.logic"

// The file-kind verdicts' kv rows, one per file uuid under a versioned prefix.

const verdictSchema = type({ v: "'transportStream' | 'notTransportStream' | 'video' | 'audio'", at: "number" })

// Boot, once storage is open: the verdicts are in memory by the first listing's paint, or follow it as one change.
export function connectFileKindVerdictStorage(): void {
	attachFileKindVerdictStorage({
		load: async () =>
			(await kvEntriesJson(VERDICT_KV_PREFIX, verdictSchema)).map(
				([key, value]) => [key.slice(VERDICT_KV_PREFIX.length), value] as const
			),
		save: (uuid, value) => kvSetJson(`${VERDICT_KV_PREFIX}${uuid}`, value),
		remove: uuid => kvDelete(`${VERDICT_KV_PREFIX}${uuid}`)
	})
}
