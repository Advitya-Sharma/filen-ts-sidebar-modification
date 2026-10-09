import type { ContentVerdict } from "@filen/shared"
import type { VideoProbe } from "@/lib/media/probe/types"

export const VERDICT_KV_PREFIX = "files.v1.kind."
export const VERDICT_MAX_RECORDS = 4096

export interface StoredVerdict {
	v: ContentVerdict
	at: number
}

// The verdicts to keep from what kv holds and what this session already wrote (which wins), newest
// `max`, oldest first; and the stored ones to delete.
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

// What a container probe of a file says about its kind: a transport stream or not (for a `.ts`), else
// whether it holds a picture. Null when it says nothing (a probe that found no tracks).
export function probeVerdict(
	probe: Pick<VideoProbe, "container" | "video" | "audio" | "partial"> | null,
	transportStreamName: boolean
): ContentVerdict | null {
	if (transportStreamName) {
		return probe?.container === "ts" ? "transportStream" : "notTransportStream"
	}

	if (probe === null) {
		return null
	}

	if (probe.video !== null) {
		return "video"
	}

	// Its picture track may lie past what a partial walk read.
	return !probe.partial && probe.audio.length > 0 ? "audio" : null
}
