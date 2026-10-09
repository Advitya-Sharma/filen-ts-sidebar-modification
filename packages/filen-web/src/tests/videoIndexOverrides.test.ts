import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { DirsAndFiles } from "@filen/sdk-rs"

// Corrections in the library index: a correction re-applies over the last classification instead of classifying the
// library again (only a split changes what the classifier reads), one set while the first classification runs isn't
// lost, and a change only a title's episodes see still refreshes that title.

const { classified } = vi.hoisted(() => ({ classified: { count: 0, slow: false } }))

vi.mock("@filen/shared/videos", async importOriginal => {
	const actual = await importOriginal<typeof import("@filen/shared/videos")>()

	return {
		...actual,
		// Counted; when `slow`, the first step outlasts a slice so the engine yields (and a test can hold it there).
		classifyVideoLibrarySteps: function* (input: Parameters<typeof actual.classifyVideoLibrarySteps>[0]) {
			classified.count++

			if (classified.slow) {
				const start = performance.now()

				while (performance.now() - start < 12) {
					// Busy on purpose: a slice is measured in wall time.
				}

				yield
			}

			return yield* actual.classifyVideoLibrarySteps(input)
		}
	}
})

import { splitOverrideKey } from "@filen/shared/videos"
import { VideoIndexEngine, type SnapshotStore } from "@/workers/videoIndex/engine"
import type { VideoIndexUpdate, VideoTitleView } from "@/workers/videoIndex/view.logic"
import { FakeDrive, hexUuid } from "@/tests/support/videoDrive"

let drive: FakeDrive
let titles: Map<string, VideoTitleView>
let upserted: string[]

let snapshots: Map<string, [Uint8Array | null, Uint8Array | null]>

const store: SnapshotStore = {
	read: root => Promise.resolve(snapshots.get(root) ?? [null, null]),
	readGenerations: () => Promise.resolve([null, null]),
	write: (root, slot, bytes) => {
		const slots = snapshots.get(root) ?? [null, null]

		slots[slot] = bytes
		snapshots.set(root, slots)

		return Promise.resolve()
	},
	remove: root => {
		snapshots.delete(root)

		return Promise.resolve()
	}
}

function engineFor(yieldSlice: () => Promise<void> = () => Promise.resolve()): VideoIndexEngine {
	const engine = new VideoIndexEngine({
		walkRoot: uuid => {
			const raw: DirsAndFiles = drive.walk(uuid)
			const meta = drive.dirs.get(uuid)?.meta

			return Promise.resolve({ type: "ok", name: meta?.type === "decoded" ? meta.data.name : "", raw, ancestors: [] })
		},
		walkMark: () => 0,
		walkDir: uuid => Promise.resolve(drive.walk(uuid)),
		lastDriveEventId: () => Promise.resolve(50n),
		driveRoot: () => null,
		getFile: uuid => Promise.resolve(drive.files.get(uuid)),
		store,
		now: () => Date.now(),
		yieldSlice
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
			upserted.push(title.id)
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

beforeEach(() => {
	vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] })
	drive = new FakeDrive()
	titles = new Map()
	upserted = []
	snapshots = new Map()
	classified.count = 0
	classified.slow = false
})

afterEach(() => {
	vi.useRealTimers()
})

describe("corrections in the library index", () => {
	it("re-applies a correction without classifying again; only a split does", async () => {
		const a = drive.addDir(hexUuid(), "Shows A")
		const b = drive.addDir(hexUuid(), "Shows B")

		drive.addFile(drive.addDir(drive.addDir(a, "Show (2020)"), "Season 1"), "Show (2020) S01E01.mkv")
		drive.addFile(drive.addDir(drive.addDir(b, "Show (2020)"), "Season 1"), "Show (2020) S01E02.mkv")

		const engine = engineFor()

		await engine.configure([
			{ uuid: a, type: "shows", enabled: true },
			{ uuid: b, type: "shows", enabled: true }
		])
		await engine.check(50n)
		await settle()

		const series = [...titles.values()]
		const before = classified.count

		expect(series).toHaveLength(1)

		engine.setOverrides({ [`d:${series[0]?.id ?? ""}`]: { title: "Renamed" } })
		await settle()

		expect([...titles.values()].map(title => title.title)).toEqual(["Renamed"])
		expect(classified.count).toBe(before)

		engine.setOverrides({
			[`d:${series[0]?.id ?? ""}`]: { title: "Renamed" },
			[splitOverrideKey(series[0]?.mergeKey ?? "")]: { split: true }
		})
		await settle()

		expect(titles.size).toBe(2)
		expect(classified.count).toBe(before + 1)
	})

	it("keeps a correction set while the first classification is still running", async () => {
		const a = drive.addDir(hexUuid(), "Shows A")
		const b = drive.addDir(hexUuid(), "Shows B")

		drive.addFile(drive.addDir(drive.addDir(a, "Show (2020)"), "Season 1"), "Show (2020) S01E01.mkv")
		drive.addFile(drive.addDir(drive.addDir(b, "Show (2020)"), "Season 1"), "Show (2020) S01E02.mkv")

		const roots = [
			{ uuid: a, type: "shows" as const, enabled: true },
			{ uuid: b, type: "shows" as const, enabled: true }
		]
		const first = engineFor()

		first.setWriter(true)
		await first.configure(roots)
		await first.check(50n)
		await first.flush()

		// A new session loads the snapshots and classifies them; nothing walks or patches afterwards.
		let release: () => void = () => undefined
		const gate = new Promise<void>(resolve => {
			release = resolve
		})
		const second = engineFor(() => gate)

		titles.clear()
		classified.slow = true

		const configured = second.configure(roots)

		await vi.advanceTimersByTimeAsync(0)
		second.setOverrides({ [splitOverrideKey("show|2020")]: { split: true } })
		classified.slow = false
		release()
		await configured
		await settle()

		expect([...titles.values()].filter(title => title.kind === "series")).toHaveLength(2)
	})

	it("refreshes a series when only an episode's correction changes it", async () => {
		const root = drive.addDir(hexUuid(), "Shows")
		const season = drive.addDir(drive.addDir(root, "Show"), "Season 1")
		const file = drive.addFile(season, "Show S01E01.mkv")
		const engine = engineFor()

		await engine.configure([{ uuid: root, type: "shows", enabled: true }])
		await engine.check(50n)
		await settle()

		const id = [...titles.values()][0]?.id ?? ""
		const hash = titles.get(id)?.detailHash

		upserted = []
		engine.setOverrides({ [`s:${file.stableUUID ?? file.uuid}`]: { title: "Renamed episode" } })
		await settle()

		expect(upserted).toEqual([id])
		expect(titles.get(id)?.detailHash).not.toBe(hash)
		expect(titles.get(id)?.title).toBe("Show")
	})
})

describe("transport-stream verdicts in the library index", () => {
	it("leaves out a .ts a content read found to be TypeScript, classifying again without a walk", async () => {
		const root = drive.addDir(hexUuid(), "Movies")

		drive.addFile(root, "Heat (1995).ts", { size: 40 * 1024 * 1024 })

		const code = drive.addFile(root, "Big (1988).ts", { size: 40 * 1024 * 1024 })
		const engine = engineFor()

		await engine.configure([{ uuid: root, type: "movies", enabled: true }])
		await engine.check(50n)
		await settle()

		expect([...titles.values()].map(title => title.title).sort()).toEqual(["Big", "Heat"])

		const walk = vi.spyOn(drive, "walk")

		engine.setNotTransportStreams([code.uuid])
		engine.setNotTransportStreams([code.uuid])
		await settle()

		expect([...titles.values()].map(title => title.title)).toEqual(["Heat"])
		expect(walk).not.toHaveBeenCalled()
	})
})
