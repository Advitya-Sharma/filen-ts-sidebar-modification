import type { ContentVerdict } from "./fileKind"

// How both apps keep the content verdicts resolveFileKind reads (fileKind.ts): one kv row per file uuid under a
// versioned prefix, the newest VERDICT_MAX_RECORDS.

export const VERDICT_KV_PREFIX = "files.v1.kind."
export const VERDICT_MAX_RECORDS = 4096

export interface StoredVerdict {
	v: ContentVerdict
	at: number
}

const VERDICTS: ReadonlySet<unknown> = new Set<ContentVerdict>(["transportStream", "notTransportStream", "video", "audio"])

// Whether a kv row read back is a stored verdict.
export function isStoredVerdict(value: unknown): value is StoredVerdict {
	if (typeof value !== "object" || value === null) {
		return false
	}

	const { v, at } = value as { v?: unknown; at?: unknown }

	return VERDICTS.has(v) && typeof at === "number" && Number.isFinite(at)
}

// The verdicts to keep from what kv holds and what this session already wrote (which wins), newest `max`,
// oldest first; and the stored ones to delete.
export function boundedVerdicts<T extends { at: number }>(
	stored: readonly (readonly [string, T])[],
	session: readonly (readonly [string, T])[],
	max: number
): { keep: (readonly [string, T])[]; drop: string[] } {
	const merged = new Map<string, T>(stored)

	for (const [uuid, value] of session) {
		merged.set(uuid, value)
	}

	const sorted = [...merged.entries()].sort((a, b) => a[1].at - b[1].at)
	const cut = Math.max(0, sorted.length - max)

	return { keep: sorted.slice(cut), drop: sorted.slice(0, cut).map(([uuid]) => uuid) }
}
