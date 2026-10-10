import { describe, expect, it } from "vitest"
import { applyWalk, emptyRootState, fromRecords, projectWalk, type RootState } from "@/workers/videoIndex/rootIndex.logic"
import {
	SNAPSHOT_FILE_BYTES,
	SNAPSHOT_HEADER_BYTES,
	crc32,
	decodeSnapshot,
	encodeSnapshot,
	nextSlot,
	pickSnapshot,
	readSnapshotHeader
} from "@/workers/videoIndex/snapshot.logic"
import { FakeDrive, hexUuid } from "@/tests/support/videoDrive"

const ROOT = hexUuid()

function indexedRoot(): RootState {
	const drive = new FakeDrive()
	const movie = drive.addDir(ROOT, "Heat (1995)")
	const show = drive.addDir(ROOT, "Show")
	const season = drive.addDir(show, "Season 1")

	drive.addFile(movie, "Heat (1995).mkv", { timestamp: 1_000 })
	drive.addFile(movie, "Heat (1995).en.srt")
	drive.addFile(movie, "poster.jpg")
	drive.addFile(movie, "notes.txt")
	drive.addFile(season, "Show S01E01 Pilot – ünïcödé.mkv", { key: "abcdefghijklmnopqrstuvwxyz012345", version: 2 })
	drive.addFile(season, "Show S01E02.mkv", { uuid: "LEGACY-UUID-0000-0000-000000000000" })
	drive.addFile(season, "Show S01E03.mkv", { key: "short-key", version: 1 })

	const state = emptyRootState(ROOT, "Library ✓")

	applyWalk({
		state,
		walk: projectWalk(ROOT, drive.walk(ROOT)),
		startedAt: 5_000,
		now: 6_000
	})
	state.watermark = 123_456_789_012n
	state.generation = 7n

	return state
}

function roundTrip(state: RootState): RootState {
	const decoded = decodeSnapshot(encodeSnapshot(state), state.uuid)

	if (decoded === null) {
		throw new Error("snapshot did not decode")
	}

	return fromRecords(decoded)
}

describe("library index snapshot", () => {
	it("round-trips every field, unicode names, legacy uuids and keys of every format", () => {
		const state = indexedRoot()
		const restored = roundTrip(state)

		expect(restored.name).toBe("Library ✓")
		expect(restored.watermark).toBe(123_456_789_012n)
		expect(restored.generation).toBe(7n)
		expect(restored.completeAt).toBe(5_000)
		expect(restored.walkedEntries).toBe(state.walkedEntries)
		expect(restored.dirs).toEqual(state.dirs)
		expect(restored.files).toEqual(state.files)
		expect(restored.byUuid).toEqual(state.byUuid)
		// notes.txt is not video-related, so it is not indexed.
		expect([...restored.files.values()].map(file => file.name)).not.toContain("notes.txt")
	})

	it("keeps an unknown watermark unknown", () => {
		const state = indexedRoot()

		state.watermark = null

		expect(roundTrip(state).watermark).toBeNull()
	})

	it("stays near 180 bytes per indexed file", () => {
		const state = indexedRoot()
		const bytes = encodeSnapshot(state)
		const perFile = (bytes.length - SNAPSHOT_HEADER_BYTES) / state.files.size

		expect(SNAPSHOT_FILE_BYTES).toBeLessThan(160)
		expect(perFile).toBeLessThan(260)
	})

	it("rejects a flipped byte, a truncated file, another root's snapshot and foreign bytes", () => {
		const bytes = encodeSnapshot(indexedRoot())
		const flipped = bytes.slice()

		flipped[bytes.length - 3] = (flipped[bytes.length - 3] ?? 0) ^ 0xff

		expect(decodeSnapshot(flipped, ROOT)).toBeNull()
		expect(decodeSnapshot(bytes.slice(0, bytes.length - 1), ROOT)).toBeNull()
		expect(decodeSnapshot(bytes, hexUuid())).toBeNull()
		expect(
			decodeSnapshot(new TextEncoder().encode("not a snapshot at all, but long enough to hold a header ............"), ROOT)
		).toBeNull()
	})

	it("computes the IEEE CRC-32", () => {
		expect(crc32(new TextEncoder().encode("123456789"))).toBe(0xcbf43926)
		expect(crc32(new Uint8Array(0))).toBe(0)
	})

	it("matches a byte-at-a-time CRC-32 at every length around the eight-byte steps", () => {
		function reference(bytes: Uint8Array): number {
			let crc = 0xffffffff

			for (const byte of bytes) {
				crc ^= byte

				for (let k = 0; k < 8; k++) {
					crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1
				}
			}

			return (crc ^ 0xffffffff) >>> 0
		}

		const bytes = Uint8Array.from({ length: 70 }, (_, i) => (i * 167 + 13) & 0xff)

		for (let length = 0; length <= bytes.length; length++) {
			expect(crc32(bytes.subarray(0, length))).toBe(reference(bytes.subarray(0, length)))
		}

		expect(crc32(bytes.subarray(3, 61))).toBe(reference(bytes.subarray(3, 61)))
	})
})

describe("snapshot slots", () => {
	function withGeneration(generation: bigint): Uint8Array {
		const state = indexedRoot()

		state.generation = generation

		return encodeSnapshot(state)
	}

	it("loads the slot with the highest generation", () => {
		const loaded = pickSnapshot([withGeneration(4n), withGeneration(5n)], ROOT)

		expect(loaded?.header.generation).toBe(5n)
	})

	it("falls back to the other slot when the newest fails its CRC", () => {
		const newest = withGeneration(9n)

		newest[newest.length - 1] = (newest[newest.length - 1] ?? 0) ^ 1

		expect(pickSnapshot([newest, withGeneration(8n)], ROOT)?.header.generation).toBe(8n)
	})

	it("has nothing to load when neither slot is intact", () => {
		const broken = withGeneration(2n)

		broken[SNAPSHOT_HEADER_BYTES + 1] = (broken[SNAPSHOT_HEADER_BYTES + 1] ?? 0) ^ 1

		expect(pickSnapshot([broken, null], ROOT)).toBeNull()
		expect(pickSnapshot([null, null], ROOT)).toBeNull()
	})

	it("reads the generation from the header alone", () => {
		expect(readSnapshotHeader(withGeneration(42n).slice(0, SNAPSHOT_HEADER_BYTES))?.generation).toBe(42n)
	})

	it("overwrites the older slot with a generation past both", () => {
		expect(nextSlot([null, null])).toEqual({ slot: 0, generation: 1n })
		expect(nextSlot([3n, null])).toEqual({ slot: 1, generation: 4n })
		expect(nextSlot([3n, 4n])).toEqual({ slot: 0, generation: 5n })
		expect(nextSlot([6n, 5n])).toEqual({ slot: 1, generation: 7n })
	})

	it("never overwrites the slot known intact while it holds that generation, whatever the other's header says", () => {
		// A torn write left generation 4's header over a broken body; generation 3 is what loaded.
		expect(nextSlot([3n, 4n], { slot: 0, generation: 3n })).toEqual({ slot: 1, generation: 5n })
		expect(nextSlot([5n, 4n], { slot: 1, generation: 4n })).toEqual({ slot: 0, generation: 6n })
		// Another writer replaced it since: back to the older slot.
		expect(nextSlot([5n, 4n], { slot: 0, generation: 3n })).toEqual({ slot: 1, generation: 6n })
	})
})

describe("snapshot interning", () => {
	it("keeps every bucket right past 65,535 distinct strings", () => {
		const state = emptyRootState(ROOT, "Big")

		for (let i = 0; i < 70_000; i++) {
			const uuid = hexUuid()

			state.files.set(uuid, {
				uuid,
				stableUuid: uuid,
				parent: ROOT,
				name: `Clip ${String(i)}.mkv`,
				size: 1,
				metaSize: 1,
				chunks: 1,
				timestamp: i,
				created: Number.NaN,
				modified: i,
				dateAdded: i,
				tombstoneAt: 0,
				tombstoneWalks: 0,
				key: "c".repeat(64),
				keyVersion: 3,
				mime: "video/mp4",
				region: "de-1",
				bucket: `bucket-${String(i)}`,
				canMakeThumbnail: false,
				favorited: false
			})
		}

		const restored = roundTrip(state)
		const last = [...restored.files.values()].at(-1)

		expect(last?.bucket).toBe("bucket-69999")
		expect([...restored.files.values()].every((file, index) => file.bucket === `bucket-${String(index)}`)).toBe(true)
	})
})
