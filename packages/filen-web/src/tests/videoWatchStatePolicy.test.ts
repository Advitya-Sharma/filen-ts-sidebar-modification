import { describe, expect, it } from "vitest"
import type { FileVersion } from "@filen/sdk-rs"
import {
	BACKOFF_MAX_MS,
	BACKOFF_MIN_MS,
	MAX_UPLOADS_PER_HOUR,
	PENDING_MAX_BYTES,
	SEMANTIC_RESERVED_UPLOADS,
	SEMANTIC_DEBOUNCE_MS,
	UPLOAD_FLOOR_MS,
	URGENT_UPLOAD_FLOOR_MS,
	decodePending,
	earliestUploadAt,
	emptyWatchMeta,
	encodePending,
	gcTargets,
	journalDue,
	mergeFlushTime,
	nextBackoffMs,
	playbackFlushReason,
	pruneTargets,
	isGoneError,
	withOwnUpload,
	type OwnUpload,
	type PendingEntry
} from "@/features/videos/lib/watchState.logic"
import { driveFile, hexUuid } from "@/tests/support/videoDrive"

const NOW = 1_800_000_000_000
const DAY = 86_400_000

describe("journal policy", () => {
	it("journals the first report and every state change", () => {
		expect(journalDue({ now: NOW, pos: 0, last: null, stateChanged: false })).toBe(true)
		expect(journalDue({ now: NOW, pos: 1, last: { pos: 0, at: NOW }, stateChanged: true })).toBe(true)
	})

	it("journals at most every 15 s and only once the position moved 10 s", () => {
		const last = { pos: 100, at: NOW }

		expect(journalDue({ now: NOW + 14_999, pos: 130, last, stateChanged: false })).toBe(false)
		expect(journalDue({ now: NOW + 15_000, pos: 109, last, stateChanged: false })).toBe(false)
		expect(journalDue({ now: NOW + 15_000, pos: 110, last, stateChanged: false })).toBe(true)
		expect(journalDue({ now: NOW + 60_000, pos: 85, last, stateChanged: false })).toBe(true)
	})
})

describe("upload triggers", () => {
	const synced = { pos: 100, at: NOW }

	it("uploads on pause once the position moved 30 s and a minute passed since the last upload", () => {
		expect(playbackFlushReason({ now: NOW, pos: 130, playing: false, paused: true, synced, lastUploadAt: NOW - 60_000 })).toBe("pause")
		expect(playbackFlushReason({ now: NOW, pos: 129, playing: false, paused: true, synced, lastUploadAt: NOW - 60_000 })).toBeNull()
		expect(playbackFlushReason({ now: NOW, pos: 200, playing: false, paused: true, synced, lastUploadAt: NOW - 59_999 })).toBeNull()
	})

	it("uploads every 5 min of playback once the position moved 30 s", () => {
		const at = NOW + 5 * 60_000

		expect(playbackFlushReason({ now: at, pos: 400, playing: true, paused: false, synced, lastUploadAt: NOW })).toBe("periodic")
		expect(playbackFlushReason({ now: at - 1, pos: 400, playing: true, paused: false, synced, lastUploadAt: NOW })).toBeNull()
		expect(playbackFlushReason({ now: at, pos: 129, playing: true, paused: false, synced, lastUploadAt: NOW })).toBeNull()
		expect(playbackFlushReason({ now: at, pos: 400, playing: false, paused: false, synced, lastUploadAt: NOW })).toBeNull()
	})

	it("keeps 60 s between uploads, 10 s for semantic and hidden triggers", () => {
		const meta = { ...emptyWatchMeta("a"), up: [NOW] }

		expect(earliestUploadAt({ reason: "periodic", now: NOW + 1, meta })).toBe(NOW + UPLOAD_FLOOR_MS)
		expect(earliestUploadAt({ reason: "pause", now: NOW + 1, meta })).toBe(NOW + UPLOAD_FLOOR_MS)
		expect(earliestUploadAt({ reason: "semantic", now: NOW + 1, meta })).toBe(NOW + URGENT_UPLOAD_FLOOR_MS)
		expect(earliestUploadAt({ reason: "hidden", now: NOW + 1, meta })).toBe(NOW + URGENT_UPLOAD_FLOOR_MS)
		expect(earliestUploadAt({ reason: "semantic", now: NOW + 20_000, meta })).toBe(NOW + 20_000)
	})

	it("waits out a failure's backoff", () => {
		const meta = { ...emptyWatchMeta("a"), bu: NOW + 30_000 }

		expect(earliestUploadAt({ reason: "semantic", now: NOW, meta })).toBe(NOW + 30_000)
	})

	it("leaves the last uploads of the hour to semantic changes", () => {
		const up = Array.from({ length: MAX_UPLOADS_PER_HOUR - SEMANTIC_RESERVED_UPLOADS }, (_, i) => NOW - 3_000_000 + i * 60_000)
		const meta = { ...emptyWatchMeta("a"), up }

		expect(earliestUploadAt({ reason: "semantic", now: NOW, meta })).toBe(NOW)
		expect(earliestUploadAt({ reason: "hidden", now: NOW, meta })).toBe(NOW - 3_000_000 + 3_600_000)
		expect(earliestUploadAt({ reason: "periodic", now: NOW, meta })).toBe(NOW - 3_000_000 + 3_600_000)
	})

	it("never starts a 31st upload within an hour", () => {
		const up = Array.from({ length: MAX_UPLOADS_PER_HOUR }, (_, i) => NOW - 3_000_000 + i * 60_000)
		const meta = { ...emptyWatchMeta("a"), up }

		expect(earliestUploadAt({ reason: "semantic", now: NOW, meta })).toBe(NOW - 3_000_000 + 3_600_000)
	})

	it("debounces semantic events by 2 s and otherwise keeps the earlier timer", () => {
		expect(mergeFlushTime({ existing: null, candidate: NOW, debounce: true, now: NOW })).toBe(NOW + SEMANTIC_DEBOUNCE_MS)
		expect(mergeFlushTime({ existing: NOW + 1_000, candidate: NOW + 30_000, debounce: false, now: NOW })).toBe(NOW + 1_000)
		expect(mergeFlushTime({ existing: NOW + 1_000, candidate: NOW + 1_500, debounce: true, now: NOW + 500 })).toBe(NOW + 2_500)
	})

	it("backs off from 30 s to 30 min", () => {
		const steps = [nextBackoffMs(0)]

		for (let i = 0; i < 8; i++) {
			steps.push(nextBackoffMs(steps.at(-1) ?? 0))
		}

		expect(steps[0]).toBe(BACKOFF_MIN_MS)
		expect(steps[1]).toBe(60_000)
		expect(steps.at(-1)).toBe(BACKOFF_MAX_MS)
	})
})

describe("version cleanup", () => {
	function own(...uuids: string[]): OwnUpload[] {
		return uuids.map(uuid => [uuid, "L"])
	}

	const head = (uuid: string) => ({ uuid, stableUuid: "L" })

	it("deletes this browser's superseded uploads but the newest of them, never the head", () => {
		expect(pruneTargets({ own: own("a", "b", "c", "d"), head: head("d") })).toEqual(["a", "b"])
		// A foreign save is the head: the newest own upload is the kept copy.
		expect(pruneTargets({ own: own("a", "b", "c"), head: head("x") })).toEqual(["a", "b"])
		// A restore made an old own upload the head again.
		expect(pruneTargets({ own: own("a", "b", "c", "d"), head: head("b") })).toEqual(["a", "c"])
		expect(pruneTargets({ own: own("a", "b"), head: head("b") })).toEqual([])
	})

	it("never touches uploads of another lineage (the file renamed, moved or trashed, a new one started)", () => {
		const uploads: OwnUpload[] = [
			["o1", "OLD"],
			["o2", "OLD"],
			["o3", "OLD"],
			["n1", "NEW"],
			["n2", "NEW"]
		]

		expect(pruneTargets({ own: uploads, head: { uuid: "n2", stableUuid: "NEW" } })).toEqual([])
		expect(pruneTargets({ own: [...uploads, ["n3", "NEW"]], head: { uuid: "n3", stableUuid: "NEW" } })).toEqual(["n1"])
	})

	it("remembers a bounded list of own uploads, newest last", () => {
		let uploads: OwnUpload[] = []

		for (let i = 0; i < 100; i++) {
			uploads = withOwnUpload(uploads, [`u${String(i)}`, "L"])
		}

		expect(uploads).toHaveLength(64)
		expect(uploads.at(-1)?.[0]).toBe("u99")
	})

	function version(ageDays: number): FileVersion {
		const file = driveFile(hexUuid(), "watch-state.v1.json", { timestamp: NOW - ageDays * DAY })

		return {
			bucket: file.bucket,
			region: file.region,
			chunks: file.chunks,
			size: file.size,
			metadata: file.meta,
			timestamp: file.timestamp,
			uuid: file.uuid,
			stableUUID: file.uuid
		}
	}

	it("collects versions older than a week beyond the newest two archived ones", () => {
		const head = version(0)
		const recent = [version(1), version(8)]
		const old = [version(9), version(30)]
		const fresh = version(3)
		const versions = [head, ...recent, fresh, ...old]

		const targets = gcTargets({ versions, headUuid: head.uuid, protect: new Set(), now: NOW })

		expect(targets.map(item => item.uuid).sort()).toEqual([...old.map(item => item.uuid), recent[1]?.uuid].sort())
		expect(targets.map(item => item.uuid)).not.toContain(head.uuid)
		expect(gcTargets({ versions, headUuid: head.uuid, protect: new Set([old[0]?.uuid ?? ""]), now: NOW })).toHaveLength(2)
	})
})

describe("pending progress", () => {
	function entry(stable: string, t: number): PendingEntry {
		return [stable, hexUuid(), 1000, 50, 600, NOW, t, "dev.tab", -1]
	}

	it("keeps the newer entry per file and replaces another account's", () => {
		const previous = decodePending(encodePending({ previous: null, account: "a", entries: [entry("s1", 5), entry("s2", 5)] }))
		const merged = decodePending(encodePending({ previous, account: "a", entries: [entry("s1", 9), entry("s2", 1)] }))

		expect(merged?.e.find(item => item[0] === "s1")?.[6]).toBe(9)
		expect(merged?.e.find(item => item[0] === "s2")?.[6]).toBe(5)
		expect(decodePending(encodePending({ previous, account: "b", entries: [entry("s3", 1)] }))?.e).toHaveLength(1)
	})

	it("stays under 4 KiB, dropping the oldest entries", () => {
		const entries = Array.from({ length: 200 }, (_, i) => entry(hexUuid(), i))
		const text = encodePending({ previous: null, account: "a", entries })
		const decoded = decodePending(text)

		expect(text.length).toBeLessThanOrEqual(PENDING_MAX_BYTES)
		expect(decoded?.e[0]?.[6]).toBe(199)
	})

	it("drops malformed entries and rejects malformed text", () => {
		expect(decodePending("{")).toBeNull()
		expect(decodePending(JSON.stringify({ a: "a", e: [["s", "u", "x"]] }))?.e).toEqual([])
	})
})

describe("downloads that found nothing", () => {
	it("recognizes a deleted file or version by the SDK's error kinds, nothing else", () => {
		expect(isGoneError({ kind: "FileChunkNotFound" })).toBe(true)
		expect(isGoneError({ kind: "FileNotFound" })).toBe(true)
		expect(isGoneError({ kind: "Reqwest" })).toBe(false)
		expect(isGoneError(new Error("offline"))).toBe(false)
	})
})
