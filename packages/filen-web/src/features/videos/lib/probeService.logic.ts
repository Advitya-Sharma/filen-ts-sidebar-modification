import { type } from "arktype"
import { MEDIA_CODECS, type ProbeTier, type VideoProbe } from "@/lib/media/probe/types"

// The probe scheduler's rules and the persisted probe record. Plain and DOM-free.

// Where a probe request queues, lowest first: what the user acted on (play, a detail page), a card that
// stayed visible, a card near the viewport once scrolling stopped, then idle prefetch.
export const PROBE_RANK = {
	explicit: 0,
	visible: 1,
	nearViewport: 2,
	idle: 3
} as const

// How long a card stays visible before it asks for its probe.
export const CARD_PROBE_DWELL_MS = 400

export const PROBE_CONCURRENCY = 2

// A probe nobody wants any more keeps running this long, for a card scrolled straight back.
export const PROBE_DROP_GRACE_MS = 300

// An unwanted probe this many blocks from the end of its current read budget finishes (its record is
// persisted); any further from it, it is cancelled.
export const PROBE_FINISH_WITHIN_BLOCKS = 1

// The third failed download for a file ends its probing for the session, until the connection returns.
export const PROBE_FAILURE_LIMIT = 3

// Probes started per session before only explicit requests may start more.
export const PROBE_SESSION_SOFT_CAP = 200

export const PROBE_KV_PREFIX = "videos.v1.probe."
export const PROBE_MAX_RECORDS = 20_000

const probeVideoSchema = type({
	codec: type.enumerated(...MEDIA_CODECS).or("null"),
	fourcc: "string",
	codecString: "string | null",
	w: "number",
	h: "number",
	par: "number",
	rotation: "number",
	bitDepth: "number | null",
	transfer: "'sdr' | 'pq' | 'hlg'",
	"dv?": { profile: "number", compat: "number" }
})

const videoProbeSchema = type({
	v: "1",
	container: "'mp4' | 'mov' | 'mkv' | 'webm'",
	durationSec: "number | null",
	index: "'moov' | 'cues' | 'mfra' | 'none' | null",
	video: probeVideoSchema.or("null"),
	audio: type({
		id: "number",
		codec: type.enumerated(...MEDIA_CODECS).or("null"),
		codecString: "string | null",
		ch: "number | null",
		rate: "number | null",
		lang: "string | null",
		name: "string | null",
		def: "boolean",
		forced: "boolean",
		commentary: "boolean"
	}).array(),
	subs: type({
		id: "number",
		codecId: "string",
		lang: "string | null",
		name: "string | null",
		def: "boolean",
		forced: "boolean",
		hi: "boolean"
	}).array(),
	chapters: type({ t: "number", "end?": "number", title: "string | null", "skipType?": "number" }).array(),
	cover: type({ offset: "number", size: "number", mime: "string" }).or("null"),
	partial: "boolean"
})

// One file's probe, merged across the tiers read so far, keyed by its uuid (a content change rotates
// it). `probe` null is a file that is no container the probe reads. Playability is never stored: it
// depends on the browser asking, decided at play time.
export const probeRecordSchema = type({
	probe: videoProbeSchema.or("null"),
	tiers: { card: "boolean", full: "boolean", cover: "boolean" },
	at: "number"
})

export type ProbeRecord = typeof probeRecordSchema.infer

// Whether `record` already answers a request for `tier`: the full tier holds the card's fields and the cover.
export function recordSatisfies(record: ProbeRecord, tier: ProbeTier): boolean {
	return record.probe === null || record.tiers[tier] || record.tiers.full
}

// `existing` with one tier's result folded in. A full read replaces everything; a cover read holds the
// card's fields too, so it answers both; a card keeps a cover already found.
export function mergeProbe(existing: ProbeRecord | null, tier: ProbeTier, probe: VideoProbe | null, at: number): ProbeRecord {
	const tiers = { card: false, full: false, cover: false, ...existing?.tiers }

	if (probe === null) {
		return { probe: null, tiers: { card: true, full: true, cover: true }, at }
	}

	if (tier === "full") {
		return { probe, tiers: { card: true, full: true, cover: true }, at }
	}

	if (tier === "cover") {
		return { probe, tiers: { ...tiers, card: true, cover: true }, at }
	}

	return { probe: { ...probe, cover: existing?.probe?.cover ?? null }, tiers: { ...tiers, card: true }, at }
}

// Whether a request at `rank` may start a probe, with `started` probes already started this session.
export function mayStartProbe(rank: number, started: number): boolean {
	return rank <= PROBE_RANK.explicit || started < PROBE_SESSION_SOFT_CAP
}
