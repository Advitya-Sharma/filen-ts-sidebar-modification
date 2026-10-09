import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { DirsAndFiles } from "@filen/sdk-rs"
import { CLASSIFIER_VERSION } from "@filen/shared/videos"
import { VideoIndexEngine, type IndexRootConfig, type SnapshotStore } from "@/workers/videoIndex/engine"
import { SNAPSHOT_HEADER_BYTES, crc32, readSnapshotHeader } from "@/workers/videoIndex/snapshot.logic"
import { WALK_FLOOR_MS } from "@/workers/videoIndex/engine"
import type { VideoIndexUpdate, VideoTitleView } from "@/workers/videoIndex/view.logic"
import { FakeDrive, driveDir, driveEvent, driveFile, hexUuid } from "@/tests/support/videoDrive"

let drive: FakeDrive
let latest: bigint
let snapshots: Map<string, [Uint8Array | null, Uint8Array | null]>
let walkGate: Promise<void>
let calls: { walks: number; lastEventIds: number; getFile: number }
let titles: Map<string, VideoTitleView>
let walkFails: number
let walkResult: (raw: DirsAndFiles) => DirsAndFiles
let writeFailures: number
let driveRootUuid: string | null
let lastIdFails: number
let goneRoots: Set<string>

function ancestorsOf(uuid: string): string[] {
	const out: string[] = []

	for (let parent = drive.dirs.get(uuid)?.parent; parent !== undefined; parent = drive.dirs.get(parent)?.parent) {
		out.push(parent)
	}

	return out
}

const store: SnapshotStore = {
	read: rootUuid => Promise.resolve(snapshots.get(rootUuid) ?? [null, null]),
	readGenerations: rootUuid => {
		const [a, b] = snapshots.get(rootUuid) ?? [null, null]

		return Promise.resolve([
			a === null ? null : (readSnapshotHeader(a)?.generation ?? null),
			b === null ? null : (readSnapshotHeader(b)?.generation ?? null)
		])
	},
	write: (rootUuid, slot, bytes) => {
		if (writeFailures > 0) {
			writeFailures--

			return Promise.reject(new Error("quota"))
		}

		const slots = snapshots.get(rootUuid) ?? [null, null]

		slots[slot] = bytes
		snapshots.set(rootUuid, slots)

		return Promise.resolve()
	},
	remove: rootUuid => {
		snapshots.delete(rootUuid)

		return Promise.resolve()
	}
}

function engineFor(): VideoIndexEngine {
	const engine = new VideoIndexEngine({
		walkRoot: async uuid => {
			calls.walks++

			const raw: DirsAndFiles = drive.walk(uuid)
			const meta = drive.dirs.get(uuid)?.meta

			await walkGate

			if (goneRoots.has(uuid)) {
				return { type: "gone" }
			}

			if (walkFails > 0) {
				walkFails--

				throw new Error("network")
			}

			return { type: "ok", name: meta?.type === "decoded" ? meta.data.name : "", raw: walkResult(raw), ancestors: ancestorsOf(uuid) }
		},
		walkMark: () => 0,
		walkDir: uuid => Promise.resolve(drive.walk(uuid)),
		lastDriveEventId: () => {
			calls.lastEventIds++

			if (lastIdFails > 0) {
				lastIdFails--

				return Promise.reject(new Error("offline"))
			}

			return Promise.resolve(latest)
		},
		driveRoot: () => driveRootUuid,
		getFile: uuid => {
			calls.getFile++

			return Promise.resolve(drive.files.get(uuid))
		},
		store,
		now: () => Date.now(),
		yieldSlice: () => Promise.resolve()
	})

	engine.onSocketEvent({ type: "authSuccess" })
	engine.subscribe((update: VideoIndexUpdate) => {
		if (update.type !== "titles") {
			return
		}

		if (update.reset) {
			titles.clear()
		}

		for (const title of update.upserts) {
			titles.set(title.id, title)
		}

		for (const id of update.removed) {
			titles.delete(id)
		}
	})

	return engine
}

async function settle(): Promise<void> {
	await vi.advanceTimersByTimeAsync(6_000)
}

function roots(...entries: [string, IndexRootConfig["type"]][]): IndexRootConfig[] {
	return entries.map(([uuid, type]) => ({ uuid, type, enabled: true }))
}

beforeEach(() => {
	vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] })
	drive = new FakeDrive()
	latest = 50n
	snapshots = new Map()
	walkGate = Promise.resolve()
	calls = { walks: 0, lastEventIds: 0, getFile: 0 }
	titles = new Map()
	walkFails = 0
	walkResult = raw => raw
	writeFailures = 0
	driveRootUuid = null
	lastIdFails = 0
	goneRoots = new Set()
})

afterEach(() => {
	vi.useRealTimers()
})

describe("library index classification", () => {
	it("merges a title found in two roots by title and year, and keeps a yearless one per root", async () => {
		const a = drive.addDir(hexUuid(), "Films")
		const b = drive.addDir(hexUuid(), "More films")

		drive.addFile(drive.addDir(a, "Heat (1995)"), "Heat (1995).mkv")
		drive.addFile(drive.addDir(b, "Heat (1995)"), "Heat (1995) - 2160p.mkv")
		drive.addFile(drive.addDir(a, "Untitled"), "Untitled.mkv")
		drive.addFile(drive.addDir(b, "Untitled"), "Untitled.mkv")

		const engine = engineFor()

		await engine.configure(roots([a, "auto"], [b, "auto"]))
		await engine.check(latest)
		await settle()

		const heat = [...titles.values()].filter(title => title.title === "Heat")

		expect(heat).toHaveLength(1)
		expect(heat[0]?.roots.sort()).toEqual([a, b].sort())
		expect([...titles.values()].filter(title => title.title === "Untitled")).toHaveLength(2)
	})

	it("honours a root's type override", async () => {
		const root = drive.addDir(hexUuid(), "Stuff")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

		const engine = engineFor()

		await engine.configure(roots([root, "other"]))
		await engine.check(latest)
		await settle()

		expect([...titles.values()].map(title => title.kind)).toEqual(["other"])
	})

	it("serves a series' episodes on request, keyless", async () => {
		const root = drive.addDir(hexUuid(), "Shows")
		const season = drive.addDir(drive.addDir(root, "Show"), "Season 1")

		drive.addFile(season, "Show S01E01.mkv")
		drive.addFile(season, "Show S01E02.mkv")
		drive.addFile(season, "Show S01E02.en.srt")

		const engine = engineFor()

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()

		const show = [...titles.values()].find(title => title.kind === "series")
		const detail = show === undefined ? null : engine.titleDetail(show.id)

		expect(detail?.kind).toBe("series")

		if (detail?.kind !== "series") {
			return
		}

		const episodes = detail.seasons[0]?.episodes ?? []

		expect(episodes.map(episode => episode.episode)).toEqual([1, 2])
		expect(episodes[1]?.versions[0]?.parts[0]?.subtitles.map(subtitle => subtitle.language)).toEqual(["en"])
		expect(JSON.stringify(detail)).not.toContain('"key"')
	})
})

describe("library index art", () => {
	it("names a title's art candidates by keyless refs", async () => {
		const root = drive.addDir(hexUuid(), "Films")
		const movie = drive.addDir(root, "Heat (1995)")
		const poster = drive.addFile(movie, "poster.jpg")

		drive.addFile(movie, "Heat (1995).mkv")

		const engine = engineFor()

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()

		const art = engine.titleArt([...titles.keys()][0] ?? "", "p")

		expect(art?.candidates[0]).toEqual({ source: "sidecar", file: 0 })
		expect(art?.files[0]).toEqual({ stableUuid: poster.uuid, uuid: poster.uuid, name: "poster.jpg", size: Number(poster.size) })
		expect(JSON.stringify(art)).not.toContain('"key"')
	})
})

describe("library index videos", () => {
	it("lists every version and episode for the home rows and search, keyless and without requests", async () => {
		const root = drive.addDir(hexUuid(), "Library")
		const movie = drive.addDir(root, "Heat (1995)")
		const season = drive.addDir(drive.addDir(root, "Lost"), "Season 1")

		drive.addFile(movie, "Heat (1995) - 1080p.mkv")
		drive.addFile(movie, "Heat (1995) - 2160p.mkv")

		const pilot = drive.addFile(season, "Lost S01E01 - Pilot.mkv")

		drive.addFile(season, "Lost S01E02.mkv")
		drive.addFile(drive.addDir(root, "Holiday Lisbon"), "VID_0001.mp4")

		const engine = engineFor()

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()

		const before = { ...calls }
		const items = engine.homeItems()
		const heat = [...titles.values()].find(title => title.kind === "movie")
		const lost = [...titles.values()].find(title => title.kind === "series")

		expect(items.filter(item => item.itemClass === "movie").map(item => [item.titleId, item.group, item.folder])).toEqual([
			[heat?.id, heat?.id, "Heat (1995)"],
			[heat?.id, heat?.id, "Heat (1995)"]
		])
		expect(
			items.filter(item => item.itemClass === "episode").map(item => [item.titleId, item.season, item.episode, item.episodeTitle])
		).toEqual([
			[lost?.id, 1, 1, "Pilot"],
			[lost?.id, 1, 2, ""]
		])
		expect(items.find(item => item.itemClass === "other")?.folder).toBe("Holiday Lisbon")
		expect(engine.homeItems()).toEqual(items)
		expect(JSON.stringify(items)).not.toContain('"key"')

		const still = engine.itemArt(lost?.id ?? "", pilot.uuid, "s")

		expect(still?.files.map(file => file.uuid)).toEqual([pilot.uuid])
		expect(calls).toEqual(before)
	})
})

describe("library index video list", () => {
	it("sends the videos only while asked, then each classification's changes, including ones no title shows", async () => {
		const root = drive.addDir(hexUuid(), "Shows")
		const season = drive.addDir(drive.addDir(root, "Lost"), "Season 1")

		drive.addFile(season, "Lost S01E01.mkv")

		const second = drive.addFile(season, "Lost S01E02.mkv")
		const engine = engineFor()
		const home: Extract<VideoIndexUpdate, { type: "home" }>[] = []
		let titleUpdates = 0

		engine.subscribe(update => {
			if (update.type === "home") {
				home.push(update)
			} else if (update.type === "titles") {
				titleUpdates++
			}
		})

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()

		expect(home).toEqual([])

		engine.setHomeEnabled(true)

		expect(home).toHaveLength(1)
		expect(home[0]).toMatchObject({ reset: true, removed: [] })
		expect(home[0]?.upserts).toHaveLength(2)

		// A new version of the second episode: its file rotates, the series' title view stays the same.
		const version = driveFile(season, "Lost S01E02.mkv", { stableUuid: second.uuid, size: 5 })
		const titlesBefore = titleUpdates

		engine.onSocketEvent(
			driveEvent(++latest, { type: "fileArchived", uuid: second.uuid, stableUUID: second.uuid, newUUID: version.uuid })
		)
		engine.onSocketEvent(driveEvent(++latest, { type: "fileNew", file: version }))
		await settle()

		expect(titleUpdates).toBe(titlesBefore)
		expect(home).toHaveLength(2)
		expect(home[1]).toMatchObject({ reset: false, removed: [], upserts: [{ stableUuid: second.uuid, uuid: version.uuid, size: 5 }] })

		engine.setHomeEnabled(false)
		engine.onSocketEvent(driveEvent(++latest, { type: "fileNew", file: driveFile(season, "Lost S01E03.mkv") }))
		await settle()

		expect(home).toHaveLength(2)
	})
})

describe("library index lifecycle", () => {
	it("applies events that land during a walk over its result", async () => {
		const root = drive.addDir(hexUuid(), "Films")
		const heat = drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")
		let release = (): void => undefined

		walkGate = new Promise(resolve => {
			release = resolve
		})

		const engine = engineFor()

		await engine.configure(roots([root, "auto"]))

		const checking = engine.check(latest)

		// The walk read the tree before this trash; the event arrives while it is still out.
		engine.onSocketEvent(driveEvent(++latest, { type: "fileTrash", uuid: heat.uuid, stableUUID: heat.uuid, newUUID: undefined }))
		release()
		await checking
		await settle()

		expect(titles.size).toBe(0)
		expect(engine.rootStatuses()[0]).toMatchObject({ current: true, state: "ready" })
	})

	it("checks once after a reconnect and walks only when something was missed", async () => {
		const root = drive.addDir(hexUuid(), "Films")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

		const engine = engineFor()

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()
		calls = { walks: 0, lastEventIds: 0, getFile: 0 }

		engine.onSocketEvent({ type: "reconnecting" })
		engine.onSocketEvent({ type: "authSuccess" })
		await settle()

		expect(calls).toEqual({ walks: 0, lastEventIds: 1, getFile: 0 })
		expect(engine.rootStatuses()[0]?.current).toBe(true)
	})

	it("needs no walk when the events delivered since the snapshot loaded cover the gap", async () => {
		const root = drive.addDir(hexUuid(), "Films")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

		const first = engineFor()

		first.setWriter(true)
		await first.configure(roots([root, "auto"]))
		await first.check(latest)
		await first.flush()

		const second = engineFor()

		await second.configure(roots([root, "auto"]))
		calls = { walks: 0, lastEventIds: 0, getFile: 0 }

		const ronin = drive.addFile(root, "Ronin (1998).mkv")

		second.onSocketEvent(driveEvent(++latest, { type: "trashEmpty" }))
		second.onSocketEvent(driveEvent(++latest, { type: "fileNew", file: ronin }))
		await second.check(latest)
		await settle()

		expect(calls.walks).toBe(0)
		expect(second.rootStatuses()[0]?.current).toBe(true)
		expect([...titles.values()].map(title => title.title).sort()).toEqual(["Heat", "Ronin"])
	})

	it("rebuilds a playable file for a live indexed video only", async () => {
		const root = drive.addDir(hexUuid(), "Films")
		const heat = drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv", { key: "b".repeat(64) })
		const engine = engineFor()

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)

		const [file] = engine.playableFiles([heat.uuid])

		expect(file).toMatchObject({ uuid: heat.uuid, stableUUID: heat.uuid, parent: heat.parent, size: heat.size, chunks: heat.chunks })
		expect(file?.meta).toEqual(heat.meta)

		engine.onSocketEvent(driveEvent(++latest, { type: "fileTrash", uuid: heat.uuid, stableUUID: heat.uuid, newUUID: undefined }))

		expect(engine.playableFiles([heat.uuid])).toEqual([])
	})

	it("writes snapshots only as the leader, alternating slots with rising generations", async () => {
		const root = drive.addDir(hexUuid(), "Films")
		const engine = engineFor()

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await vi.advanceTimersByTimeAsync(120_000)

		expect(snapshots.has(root)).toBe(false)

		engine.setWriter(true)
		await vi.advanceTimersByTimeAsync(120_000)

		expect(snapshots.get(root)?.map(slot => (slot === null ? null : readSnapshotHeader(slot)?.generation))).toEqual([1n, null])

		engine.onSocketEvent(driveEvent(++latest, { type: "trashEmpty" }))
		await engine.flush()

		expect(snapshots.get(root)?.map(slot => (slot === null ? null : readSnapshotHeader(slot)?.generation))).toEqual([1n, 2n])
	})

	it("keeps a root that went missing, flagged", async () => {
		const root = drive.addDir(hexUuid(), "Films")
		const engine = engineFor()

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()

		engine.onSocketEvent(
			driveEvent(++latest, { type: "folderTrash", parent: hexUuid(), uuid: root as `${string}-${string}-${string}-${string}` })
		)
		await settle()

		expect(engine.rootStatuses()[0]?.state).toBe("missing")
		expect(titles.size).toBe(1)
	})
})

// A root with one movie, indexed and current.
async function indexedRoot(engine: VideoIndexEngine): Promise<string> {
	const root = drive.addDir(hexUuid(), "Films")

	drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

	await engine.configure(roots([root, "auto"]))
	await engine.check(latest)
	await settle()
	calls = { walks: 0, lastEventIds: 0, getFile: 0 }

	return root
}

describe("library index recovery", () => {
	it("walks a root whose snapshot was written by another classifier, whatever the floor", async () => {
		const first = engineFor()

		first.setWriter(true)

		const root = await indexedRoot(first)

		await first.flush()

		const slot = snapshots.get(root)?.[0]

		if (slot === null || slot === undefined) {
			throw new Error("no snapshot written")
		}

		// The classifier version sits outside the CRC; the kinds signature inside it.
		new DataView(slot.buffer).setUint16(6, 0, true)

		const second = engineFor()

		await second.configure(roots([root, "auto"]))
		await second.check(latest)

		expect(calls.walks).toBe(1)

		new DataView(slot.buffer).setUint16(6, CLASSIFIER_VERSION, true)
		new DataView(slot.buffer).setUint32(92, 1, true)
		new DataView(slot.buffer).setUint32(12, crc32(slot.subarray(16)), true)
		snapshots.set(root, [slot, null])

		const third = engineFor()

		await third.configure(roots([root, "auto"]))
		await third.check(latest)

		expect(calls.walks).toBe(2)
		expect(slot.length).toBeGreaterThan(SNAPSHOT_HEADER_BYTES)
	})

	it("a whole-drive delete empties the index at once and walks whatever the floor", async () => {
		const engine = engineFor()
		const root = await indexedRoot(engine)

		expect(titles.size).toBe(1)

		// The walk it starts doesn't come back during the test: the index is empty before it does.
		walkGate = new Promise(() => undefined)
		drive.dirs.set(root, driveDir(hexUuid(), "Films", root as `${string}-${string}-${string}-${string}`))
		engine.onSocketEvent(driveEvent(++latest, { type: "deleteAll" }))
		await settle()

		expect(calls.walks).toBe(1)
		expect(titles.size).toBe(0)
	})

	it("more directories moved in than are re-listed one by one walk the root, whatever the floor", async () => {
		const engine = engineFor()
		const root = await indexedRoot(engine)

		for (let i = 0; i < 9; i++) {
			const dir = driveDir(root, `Incoming ${String(i)}`)

			drive.dirs.set(dir.uuid, dir)
			engine.onSocketEvent(driveEvent(++latest, { type: "folderMove", dir }))
		}

		await settle()

		expect(calls.walks).toBe(1)
	})

	it("doesn't take the last event id from before a drop for a watermark", async () => {
		const engine = engineFor()
		const root = drive.addDir(hexUuid(), "Films")

		engine.onSocketEvent(driveEvent(60n, { type: "trashEmpty" }))
		engine.onSocketEvent({ type: "reconnecting" })
		engine.onSocketEvent({ type: "authSuccess" })
		await engine.configure(roots([root, "auto"]))
		await engine.rescan(root)

		expect(calls.lastEventIds).toBe(1)
	})

	it("writes a watermark-only change with the next write, or when the page goes away, not on its own", async () => {
		const engine = engineFor()

		engine.setWriter(true)

		const root = await indexedRoot(engine)

		await vi.advanceTimersByTimeAsync(120_000)

		const generations = (): (bigint | null | undefined)[] =>
			snapshots.get(root)?.map(slot => (slot === null ? null : readSnapshotHeader(slot)?.generation)) ?? []

		expect(generations()).toEqual([1n, null])

		engine.onSocketEvent(driveEvent(++latest, { type: "trashEmpty" }))
		await vi.advanceTimersByTimeAsync(120_000)

		expect(generations()).toEqual([1n, null])

		await engine.flush()

		expect(generations()).toEqual([1n, 2n])
	})

	it("stops writing after repeated failures instead of retrying for the session", async () => {
		const engine = engineFor()

		writeFailures = 10
		engine.setWriter(true)

		const root = await indexedRoot(engine)

		for (let i = 0; i < 6; i++) {
			drive.addFile(root, `Clip ${String(i)} (2001).mkv`)
			engine.onSocketEvent(
				driveEvent(++latest, { type: "fileNew", file: [...drive.files.values()].at(-1) ?? driveFile(root, "x.mkv") })
			)
			await vi.advanceTimersByTimeAsync(120_000)
		}

		await engine.flush()

		expect(writeFailures).toBe(7)
	})

	it("looks up a renamed file only when no root knows it", async () => {
		const engine = engineFor()
		const a = drive.addDir(hexUuid(), "Films")
		const b = drive.addDir(hexUuid(), "Shows")
		const known = drive.addFile(a, "heat.part")

		drive.addFile(a, "Ronin (1998).mkv")
		await engine.configure(roots([a, "auto"], [b, "auto"]))
		await engine.check(latest)
		await settle()

		const rename = (uuid: string, name: string) => {
			engine.onSocketEvent(
				driveEvent(++latest, {
					type: "fileMetadataChanged",
					uuid: uuid as `${string}-${string}-${string}-${string}`,
					metadata: { type: "decoded", data: { name, mime: "", modified: 0n, size: 2_000_000_000n, key: "k", version: 3 } }
				})
			)
		}
		const ronin = [...drive.files.values()].find(file => file.meta.type === "decoded" && file.meta.data.name.startsWith("Ronin"))

		calls.getFile = 0
		rename(ronin?.uuid ?? "", "Ronin.1998.mkv")
		await settle()

		expect(calls.getFile).toBe(0)

		rename(known.uuid, "Heat (1995).mkv")
		await settle()

		expect(calls.getFile).toBe(1)
	})

	it("a trashed directory above the root takes the root away", async () => {
		const engine = engineFor()
		const parent = drive.addDir(hexUuid(), "Media")
		const root = drive.addDir(parent, "Films")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")
		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()

		engine.onSocketEvent(
			driveEvent(++latest, { type: "folderTrash", parent: hexUuid(), uuid: parent as `${string}-${string}-${string}-${string}` })
		)

		expect(engine.rootStatuses()[0]?.state).toBe("missing")
	})

	it("walks a root again once the floor passes when a walk came back partial", async () => {
		const engine = engineFor()

		walkResult = raw => ({ ...raw, files: [...raw.files, driveFile(hexUuid(), "Orphan.mkv")] })

		const root = await indexedRoot(engine)

		walkResult = raw => raw

		expect(engine.rootStatuses().find(status => status.uuid === root)?.current).toBe(false)

		await vi.advanceTimersByTimeAsync(WALK_FLOOR_MS + 1_000)

		expect(calls.walks).toBe(1)
		expect(engine.rootStatuses()[0]?.current).toBe(true)
	})
})

describe("library index scan stop", () => {
	function gateWalks(): () => void {
		let release = (): void => undefined

		walkGate = new Promise(resolve => {
			release = resolve
		})

		return () => {
			release()
		}
	}

	async function firstScanStopped(engine: VideoIndexEngine): Promise<{ root: string; release: () => void; checking: Promise<void> }> {
		const root = drive.addDir(hexUuid(), "Films")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

		const release = gateWalks()

		await engine.configure(roots([root, "auto"]))

		const checking = engine.check(latest)

		await vi.advanceTimersByTimeAsync(0)
		expect(engine.rootStatuses()[0]?.walking).toBe(true)
		engine.stopWalk(root)

		return { root, release, checking }
	}

	it("flags a stopped first scan incomplete at once, then drops its result when it lands", async () => {
		const engine = engineFor()
		const { release, checking } = await firstScanStopped(engine)

		expect(engine.rootStatuses()[0]).toMatchObject({ walking: false, incomplete: true, state: "error" })

		release()
		await checking
		await settle()

		expect(titles.size).toBe(0)
		expect(engine.rootStatuses()[0]).toMatchObject({ walking: false, incomplete: true, state: "error", videos: 0 })
	})

	it("walks a stopped root again only when a rescan asks, not on the next check or configuration", async () => {
		const engine = engineFor()
		const { root, release, checking } = await firstScanStopped(engine)

		release()
		await checking
		await settle()

		expect(engine.needsCheck()).toBe(false)
		await engine.configure(roots([root, "movies"]))
		await engine.check(latest)
		await vi.advanceTimersByTimeAsync(WALK_FLOOR_MS + 1_000)
		expect(calls.walks).toBe(1)

		await engine.rescan(root)
		await settle()

		expect(calls.walks).toBe(2)
		expect(engine.rootStatuses()[0]).toMatchObject({ state: "ready", incomplete: false })
	})

	it("keeps the previous index when a rescan is stopped", async () => {
		const engine = engineFor()
		const root = await indexedRoot(engine)

		drive.addFile(drive.addDir(root, "Alien (1979)"), "Alien (1979).mkv")

		const release = gateWalks()
		const rescanning = engine.rescan(root)

		engine.stopWalk(root)

		expect(engine.rootStatuses()[0]).toMatchObject({ walking: false, state: "ready", incomplete: true })

		release()
		await rescanning
		await settle()

		expect([...titles.values()].map(title => title.title)).toEqual(["Heat"])
		expect(engine.rootStatuses()[0]).toMatchObject({ state: "ready", incomplete: true })
	})

	it("a root that was current stays current after a stopped rescan, with the events delivered meanwhile", async () => {
		const engine = engineFor()
		const root = await indexedRoot(engine)

		expect(engine.rootStatuses()[0]?.current).toBe(true)

		const release = gateWalks()
		const rescanning = engine.rescan(root)

		engine.stopWalk(root)

		const alien = drive.addFile(drive.addDir(root, "Alien (1979)"), "Alien (1979).mkv")

		engine.onSocketEvent(driveEvent(++latest, { type: "folderSubCreated", dir: drive.dirs.get(alien.parent) ?? driveDir(root, "x") }))
		engine.onSocketEvent(driveEvent(++latest, { type: "fileNew", file: alien }))
		release()
		await rescanning
		await settle()

		expect(engine.rootStatuses()[0]?.current).toBe(true)
		expect([...titles.values()].map(title => title.title).sort()).toEqual(["Alien", "Heat"])
	})

	it("shows a root trashed while its stopped walk was out as missing", async () => {
		const engine = engineFor()
		const root = await indexedRoot(engine)
		let release = (): void => undefined

		walkGate = new Promise(resolve => {
			release = resolve
		})
		goneRoots.add(root)

		const rescanning = engine.rescan(root)

		engine.stopWalk(root)
		release()
		await rescanning

		expect(engine.rootStatuses()[0]?.state).toBe("missing")
	})

	it("takes a stopped walk back when a rescan is asked for while it is still out", async () => {
		const engine = engineFor()
		const root = await indexedRoot(engine)

		drive.addFile(drive.addDir(root, "Alien (1979)"), "Alien (1979).mkv")

		const release = gateWalks()
		const rescanning = engine.rescan(root)

		engine.stopWalk(root)

		const again = engine.rescan(root)

		expect(engine.rootStatuses()[0]?.walking).toBe(true)

		release()
		await Promise.all([rescanning, again])
		await settle()

		expect(calls.walks).toBe(1)
		expect([...titles.values()].map(title => title.title).sort()).toEqual(["Alien", "Heat"])
		expect(engine.rootStatuses()[0]?.incomplete).toBe(false)
	})

	it("reads a scan as under way until its result is classified, so it ends with the new count", async () => {
		const root = drive.addDir(hexUuid(), "Films")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")
		drive.addFile(drive.addDir(root, "Alien (1979)"), "Alien (1979).mkv")

		const engine = engineFor()
		const ended: number[] = []
		let walking = false

		engine.subscribe(update => {
			const status = update.type === "roots" ? update.roots[0] : undefined

			if (status === undefined) {
				return
			}

			if (walking && !status.walking) {
				ended.push(status.videos)
			}

			walking = status.walking
		})

		await engine.configure(roots([root, "auto"]))
		await engine.check(latest)
		await settle()

		expect(ended).toEqual([2])
	})

	it("flags a failed walk incomplete until a complete one lands", async () => {
		const engine = engineFor()
		const root = await indexedRoot(engine)

		walkFails = 1
		await engine.rescan(root)

		expect(engine.rootStatuses()[0]).toMatchObject({ state: "ready", incomplete: true })

		await engine.rescan(root)

		expect(engine.rootStatuses()[0]?.incomplete).toBe(false)
	})
})

describe("library index checks", () => {
	it("walks a new root even when the latest event id can't be read, and reports a walk that fails too", async () => {
		const root = drive.addDir(hexUuid(), "Films")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")

		const engine = engineFor()

		engine.onSocketEvent({ type: "reconnecting" })
		lastIdFails = 1
		await engine.configure(roots([root, "auto"]))
		await engine.check()
		await settle()

		expect(engine.rootStatuses()[0]).toMatchObject({ state: "ready", incomplete: false })
		expect([...titles.values()].map(title => title.title)).toEqual(["Heat"])

		const other = drive.addDir(hexUuid(), "More")

		lastIdFails = 2
		await engine.configure(roots([root, "auto"], [other, "auto"]))
		await engine.check()
		await settle()

		expect(engine.rootStatuses().find(status => status.uuid === other)).toMatchObject({
			state: "error",
			incomplete: true,
			walking: false
		})
	})

	it("reads a root whose only stored walk was partial as incomplete after a reload", async () => {
		const first = engineFor()

		first.setWriter(true)
		walkResult = raw => ({ ...raw, files: [...raw.files, driveFile(hexUuid(), "Orphan.mkv")] })

		const root = drive.addDir(hexUuid(), "Films")

		drive.addFile(drive.addDir(root, "Heat (1995)"), "Heat (1995).mkv")
		await first.configure(roots([root, "auto"]))
		await first.check(latest)
		await settle()
		await first.flush()

		expect(first.rootStatuses()[0]?.incomplete).toBe(true)

		const second = engineFor()

		await second.configure(roots([root, "auto"]))

		expect(second.rootStatuses()[0]).toMatchObject({ state: "ready", incomplete: true })
	})

	it("leaves Filen's own directory out of a drive-root library, walked or created later", async () => {
		const driveRoot = hexUuid()
		const dotFilen = drive.addDir(driveRoot, ".filen")

		driveRootUuid = driveRoot
		drive.addFile(drive.addDir(dotFilen, "Chat Uploads"), "clip.mp4")
		drive.addFile(drive.addDir(driveRoot, "Heat (1995)"), "Heat (1995).mkv")

		const engine = engineFor()

		await engine.configure(roots([driveRoot, "auto"]))
		await engine.check(latest)
		await settle()

		expect([...titles.values()].map(title => title.title)).toEqual(["Heat"])

		const again = driveDir(driveRoot, ".Filen")

		engine.onSocketEvent(driveEvent(++latest, { type: "folderSubCreated", dir: again }))
		engine.onSocketEvent(driveEvent(++latest, { type: "fileNew", file: driveFile(again.uuid, "clip 2.mp4") }))
		await settle()

		expect([...titles.values()].map(title => title.title)).toEqual(["Heat"])
	})
})
