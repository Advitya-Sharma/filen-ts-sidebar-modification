import { describe, expect, it } from "vitest"
import type { DriveEvent, File } from "@filen/sdk-rs"
import {
	TOMBSTONE_MS,
	applyDriveEvent,
	applyWalk,
	emptyRootState,
	projectWalk,
	type PatchResult,
	type RootState
} from "@/workers/videoIndex/rootIndex.logic"
import { FakeDrive, driveDir, driveFile, hexUuid } from "@/tests/support/videoDrive"

const ROOT = hexUuid()
const NOW = 1_800_000_000_000

function walked(drive: FakeDrive, state: RootState, startedAt: number, complete = true): void {
	const projection = projectWalk(ROOT, drive.walk(ROOT))

	applyWalk({
		state,
		walk: {
			...projection,
			complete: projection.complete && complete
		},
		startedAt,
		now: startedAt
	})
}

function setup(): { drive: FakeDrive; state: RootState; movies: string; heat: File } {
	const drive = new FakeDrive()
	const movies = drive.addDir(ROOT, "Movies")
	const heat = drive.addFile(movies, "Heat (1995).mkv", { timestamp: 1_000 })
	const state = emptyRootState(ROOT, "Root")

	walked(drive, state, NOW)

	return { drive, state, movies, heat }
}

function apply(state: RootState, inner: DriveEvent): PatchResult {
	return applyDriveEvent({
		state,
		inner,
		now: NOW + 1
	})
}

function live(state: RootState): string[] {
	return [...state.files.values()].filter(file => file.tombstoneAt === 0).map(file => file.name)
}

describe("dateAdded", () => {
	it("is the file's timestamp on a root's first walk, never earlier than the previous complete walk after", () => {
		const { drive, state, movies } = setup()

		expect(state.files.size).toBe(1)
		expect([...state.files.values()][0]?.dateAdded).toBe(1_000)

		// Moved or copied in with an old timestamp: it joined when the library could first have seen it.
		drive.addFile(movies, "Old.mkv", { timestamp: 2_000 })
		walked(drive, state, NOW + 10_000)

		expect([...state.files.values()].find(file => file.name === "Old.mkv")?.dateAdded).toBe(NOW)
	})

	it("never changes for a known stable uuid: not on a new version, a rename, a move or a version restore", () => {
		const { state, movies, heat } = setup()
		const stableUuid = heat.stableUUID ?? heat.uuid
		const version = driveFile(movies, "Heat (1995).mkv", { stableUuid, timestamp: NOW + 500 })

		apply(state, { type: "fileArchived", uuid: heat.uuid, stableUUID: stableUuid, newUUID: version.uuid })
		apply(state, { type: "fileNew", file: version })
		apply(state, { type: "fileMove", file: { ...version, parent: ROOT } })
		apply(state, { type: "fileArchiveRestored", currentUuid: version.uuid, file: heat })

		expect(state.files.get(stableUuid)?.dateAdded).toBe(1_000)
		expect(state.files.get(stableUuid)?.uuid).toBe(heat.uuid)
	})

	it("comes back with a file that returns from the trash", () => {
		const { state, heat } = setup()

		apply(state, { type: "fileTrash", uuid: heat.uuid, stableUUID: heat.uuid, newUUID: undefined })
		expect(live(state)).toEqual([])

		apply(state, { type: "fileRestore", file: { ...heat, timestamp: BigInt(NOW + 9_000) } })
		expect(state.files.get(heat.uuid)?.dateAdded).toBe(1_000)
	})
})

describe("socket patches", () => {
	it("adds a new video in the root and ignores one elsewhere or a non-video file", () => {
		const { state, movies } = setup()

		expect(apply(state, { type: "fileNew", file: driveFile(movies, "Ronin (1998).mkv") }).changed).toBe(true)
		expect(apply(state, { type: "fileNew", file: driveFile(hexUuid(), "Elsewhere.mkv") }).changed).toBe(false)
		expect(apply(state, { type: "fileNew", file: driveFile(movies, "notes.txt") }).changed).toBe(false)
		expect(live(state).sort()).toEqual(["Heat (1995).mkv", "Ronin (1998).mkv"])
	})

	it("ignores a trash that retires a version (newUUID), tombstones a real one", () => {
		const { state, heat } = setup()

		expect(apply(state, { type: "fileTrash", uuid: heat.uuid, stableUUID: hexUuid(), newUUID: hexUuid() }).changed).toBe(false)
		expect(apply(state, { type: "fileTrash", uuid: heat.uuid, stableUUID: heat.uuid, newUUID: undefined }).changed).toBe(true)
		expect(state.files.get(heat.uuid)?.tombstoneAt).toBe(NOW + 1)
	})

	it("removes on a permanent delete only when the event names the indexed head (version deletes carry the stable uuid too)", () => {
		const { state, heat } = setup()

		// A version-only delete: its uuid is an archived version's, the stable uuid is the lineage's.
		expect(apply(state, { type: "fileDeletedPermanent", uuid: hexUuid(), stableUUID: heat.uuid }).changed).toBe(false)
		expect(live(state)).toEqual(["Heat (1995).mkv"])

		expect(apply(state, { type: "fileDeletedPermanent", uuid: heat.uuid, stableUUID: heat.uuid }).changed).toBe(true)
		expect(state.files.size).toBe(0)
	})

	it("drops a lineage replaced by a file moved over it (fileArchived without newUUID)", () => {
		const { state, heat } = setup()

		apply(state, { type: "fileArchived", uuid: heat.uuid, stableUUID: heat.uuid, newUUID: undefined })

		expect(state.files.size).toBe(0)
	})

	it("treats the files of a permanently deleted directory as gone", () => {
		const { state, movies } = setup()

		apply(state, { type: "folderDeletedPermanent", uuid: movies as File["uuid"] })

		expect(live(state)).toEqual([])
		expect(state.dirs.size).toBe(0)
	})

	it("renames a known file in place and looks up an unknown one renamed to a video name", () => {
		const { state, heat } = setup()
		const renamed = apply(state, {
			type: "fileMetadataChanged",
			uuid: heat.uuid,
			metadata: {
				type: "decoded",
				data: {
					...(heat.meta.type === "decoded" ? heat.meta.data : { mime: "", modified: 0n, size: 0n, key: "", version: 3 }),
					name: "Heat.1995.2160p.mkv"
				}
			}
		})

		expect(renamed.changed).toBe(true)
		expect(live(state)).toEqual(["Heat.1995.2160p.mkv"])

		const unknown = hexUuid()
		const lookup = apply(state, {
			type: "fileMetadataChanged",
			uuid: unknown,
			metadata: { type: "decoded", data: { name: "Movie.mkv", mime: "", modified: 0n, size: 1n, key: "k", version: 3 } }
		})

		expect(lookup.followUp).toEqual({ type: "getFile", uuid: unknown })
		expect(
			apply(state, {
				type: "fileMetadataChanged",
				uuid: hexUuid(),
				metadata: { type: "decoded", data: { name: "report.pdf", mime: "", modified: 0n, size: 1n, key: "k", version: 3 } }
			}).followUp
		).toEqual({ type: "none" })
	})

	it("re-lists a directory moved or restored into the root, and tombstones what moved out", () => {
		const { state, movies } = setup()
		const incoming = driveDir(hexUuid(), "Incoming")

		expect(apply(state, { type: "folderMove", dir: { ...incoming, parent: ROOT } }).followUp).toEqual({
			type: "scopedWalk",
			dirUuid: incoming.uuid
		})
		expect(
			apply(state, { type: "folderRestore", dir: { ...driveDir(ROOT, "Back"), parent: movies as File["uuid"] } }).followUp.type
		).toBe("scopedWalk")

		apply(state, { type: "folderMove", dir: { ...driveDir(hexUuid(), "Movies", movies as File["uuid"]) } })

		expect(live(state)).toEqual([])
		expect(state.files.size).toBe(1)
	})

	it("flags the root itself going, coming back and moving, and a whole-drive delete", () => {
		const { state } = setup()

		expect(apply(state, { type: "folderTrash", parent: hexUuid(), uuid: ROOT }).followUp).toEqual({ type: "rootGone" })
		expect(apply(state, { type: "folderRestore", dir: driveDir(hexUuid(), "Root", ROOT) }).followUp).toEqual({ type: "rootRestored" })
		expect(apply(state, { type: "folderMove", dir: driveDir(hexUuid(), "Root", ROOT) }).followUp).toEqual({ type: "rootMoved" })
		expect(apply(state, { type: "deleteAll" }).followUp).toEqual({ type: "fullWalk" })
	})

	it("empties the index at once when the whole drive is deleted", () => {
		const { state } = setup()

		expect(apply(state, { type: "deleteAll" }).changed).toBe(true)
		expect(state.files.size + state.dirs.size).toBe(0)
	})

	it("persists a favorite without reclassifying", () => {
		const { state, heat } = setup()
		const result = apply(state, { type: "itemFavorite", item: { type: "file", ...heat, favorited: true } })

		expect(result).toMatchObject({ changed: false, dirty: true })
		expect(state.files.get(heat.uuid)?.favorited).toBe(true)
	})
})

describe("tombstones", () => {
	it("a complete walk tombstones what it misses, revives what returns, and purges after a second walk and 30 days", () => {
		const { drive, state, heat } = setup()

		drive.files.delete(heat.uuid)
		walked(drive, state, NOW + 1_000)

		expect(state.files.get(heat.uuid)).toMatchObject({ tombstoneAt: NOW + 1_000, tombstoneWalks: 1 })

		// A second complete walk inside 30 days keeps it.
		walked(drive, state, NOW + 2_000)
		expect(state.files.get(heat.uuid)?.tombstoneWalks).toBe(2)

		walked(drive, state, NOW + 1_000 + TOMBSTONE_MS)
		expect(state.files.has(heat.uuid)).toBe(false)
	})

	it("a tombstoned file that comes back keeps its dateAdded", () => {
		const { drive, state, heat } = setup()

		drive.files.delete(heat.uuid)
		walked(drive, state, NOW + 1_000)
		drive.files.set(heat.uuid, heat)
		walked(drive, state, NOW + 2_000)

		expect(state.files.get(heat.uuid)).toMatchObject({ tombstoneAt: 0, dateAdded: 1_000 })
	})

	it("a partial walk only inserts and updates", () => {
		const { drive, state, heat, movies } = setup()

		drive.files.delete(heat.uuid)
		drive.addFile(movies, "Ronin (1998).mkv")
		walked(drive, state, NOW + 1_000, false)

		expect(live(state).sort()).toEqual(["Heat (1995).mkv", "Ronin (1998).mkv"])
		expect(state.completeAt).toBe(NOW)
	})

	it("an orphaned entry makes a walk partial", () => {
		const drive = new FakeDrive()

		drive.addFile(hexUuid(), "Lost.mkv")

		const raw = drive.walk(ROOT)

		raw.files.push(driveFile(hexUuid(), "Orphan.mkv"))

		expect(projectWalk(ROOT, raw).complete).toBe(false)
	})
})
