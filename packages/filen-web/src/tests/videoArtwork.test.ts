import { describe, expect, it, vi } from "vitest"
import type { AnyFile } from "@filen/sdk-rs"
import { classifyVideoLibrary, type LibraryScanFile, type VideoLibrary, type VideoProbe } from "@filen/shared/videos"

vi.mock("@/lib/sdk/client", () => ({ sdkApi: {} }))
vi.mock("@/lib/storage/adapter", () => ({
	kvGetJson: vi.fn(),
	kvSetJson: vi.fn(),
	kvKeys: vi.fn(),
	kvEntriesJson: vi.fn(),
	kvDelete: vi.fn()
}))
vi.mock("@/features/videos/workers/frame.worker.ts?worker", () => ({ default: vi.fn() }))

import { ArtLoader, type ArtFile, type ArtLoaderDeps } from "@/features/videos/lib/artwork"
import {
	ART_BOXES,
	ART_SOURCE_MAX_BYTES,
	STILL_BOX_2X,
	artBox,
	artCandidates,
	artRecordSchema,
	titleColour,
	type ArtCandidate,
	type ArtRecord
} from "@/features/videos/lib/artwork.logic"
import { mergeProbe } from "@/features/videos/lib/probeService.logic"
import type { SdkThumbnailResult } from "@/workers/sdk.worker"
import { RecordStore } from "@/features/videos/lib/recordStore.logic"
import { memoryKv } from "@/tests/support/memoryKv"

// Movies/Up (2009) with directory and stem posters, fanart and a -thumb; Movies/Bare (2010) with nothing;
// Shows/Show with its own art and a season poster.
const FILES = [
	["d1", "Up (2009).mkv"],
	["d1", "poster.jpg"],
	["d1", "Up (2009)-poster.jpg"],
	["d1", "fanart.jpg"],
	["d1", "Up (2009)-thumb.jpg"],
	["s1", "poster.jpg"],
	["s1", "fanart.jpg"],
	["s1s1", "Show S01E01.mkv"],
	["s1s1", "Show S01E01-thumb.jpg"],
	["s1s1", "poster.jpg"],
	["s1s1", "Show S01E02.mkv"],
	["d2", "Bare (2010).mp4"]
] as const

const SCAN: LibraryScanFile[] = FILES.map(([parent, name], i) => ({ uuid: `f${String(i)}`, parent, name, size: 10, timestamp: 1 }))

function library(): VideoLibrary {
	return classifyVideoLibrary({
		roots: [
			{ uuid: "rm", name: "Movies", type: "movies" },
			{ uuid: "rs", name: "Shows", type: "shows" }
		],
		dirs: [
			{ uuid: "d1", parent: "rm", name: "Up (2009)" },
			{ uuid: "d2", parent: "rm", name: "Bare (2010)" },
			{ uuid: "s1", parent: "rs", name: "Show" },
			{ uuid: "s1s1", parent: "s1", name: "Season 01" }
		],
		files: SCAN
	})
}

function index(name: string, parent: string): number {
	return FILES.findIndex(([dir, file]) => dir === parent && file === name)
}

describe("artCandidates", () => {
	const lib = library()
	const [up, bare] = lib.movies
	const [show] = lib.series
	const season = show?.seasons[0]
	const episode = season?.episodes[0]

	if (up === undefined || bare === undefined || show === undefined || season === undefined || episode === undefined) {
		throw new Error("fixture library did not classify")
	}

	const upVideo = index("Up (2009).mkv", "d1")
	const episodeVideo = index("Show S01E01.mkv", "s1s1")

	it("orders a movie poster: NFO, stem poster, embedded cover, drive thumbnail, typographic, then a frame", () => {
		expect(artCandidates(lib, { kind: "movie", movie: up, nfo: { poster: 99 } }, "p")).toEqual<ArtCandidate[]>([
			{ source: "sidecar", file: 99 },
			{ source: "sidecar", file: index("Up (2009)-poster.jpg", "d1") },
			{ source: "embedded", file: upVideo },
			{ source: "driveThumb", file: upVideo },
			{ source: "typographic" },
			{ source: "frame", file: upVideo }
		])
	})

	it("falls straight to the video's own sources for a movie with no sidecars", () => {
		const video = index("Bare (2010).mp4", "d2")

		expect(artCandidates(lib, { kind: "movie", movie: bare }, "p").map(candidate => candidate.source)).toEqual([
			"embedded",
			"driveThumb",
			"typographic",
			"frame"
		])
		expect(artCandidates(lib, { kind: "movie", movie: bare }, "b")).toEqual<ArtCandidate[]>([
			{ source: "frame", file: video },
			{ source: "gradient" }
		])
	})

	it("orders a movie backdrop and still", () => {
		expect(artCandidates(lib, { kind: "movie", movie: up }, "b")).toEqual<ArtCandidate[]>([
			{ source: "sidecar", file: index("fanart.jpg", "d1") },
			// A Matroska file's landscape cover, once its probe locates one.
			{ source: "embedded", file: upVideo },
			{ source: "frame", file: upVideo },
			{ source: "gradient" }
		])
		expect(artCandidates(lib, { kind: "movie", movie: up }, "s")).toEqual<ArtCandidate[]>([
			{ source: "sidecar", file: index("Up (2009)-thumb.jpg", "d1") },
			{ source: "sidecar", file: index("fanart.jpg", "d1") },
			{ source: "driveThumb", file: upVideo },
			{ source: "frame", file: upVideo },
			{ source: "gradient" }
		])
	})

	it("orders series and season posters", () => {
		expect(artCandidates(lib, { kind: "series", series: show }, "p")).toEqual<ArtCandidate[]>([
			{ source: "sidecar", file: index("poster.jpg", "s1") },
			{ source: "sidecar", file: index("poster.jpg", "s1s1") },
			{ source: "frame", file: episodeVideo },
			{ source: "typographic" }
		])
		expect(artCandidates(lib, { kind: "season", series: show, season }, "p")).toEqual<ArtCandidate[]>([
			{ source: "sidecar", file: index("poster.jpg", "s1s1") },
			{ source: "sidecar", file: index("poster.jpg", "s1") },
			{ source: "frame", file: episodeVideo },
			{ source: "typographic" }
		])
		expect(artCandidates(lib, { kind: "series", series: show }, "b")).toEqual<ArtCandidate[]>([
			{ source: "sidecar", file: index("fanart.jpg", "s1") },
			{ source: "frame", file: episodeVideo },
			{ source: "gradient" }
		])
	})

	it("orders an episode still: -thumb, NFO thumb, embedded, drive thumbnail, frame", () => {
		expect(artCandidates(lib, { kind: "episode", episode, nfo: { thumb: 77 } }, "s")).toEqual<ArtCandidate[]>([
			{ source: "sidecar", file: index("Show S01E01-thumb.jpg", "s1s1") },
			{ source: "sidecar", file: 77 },
			{ source: "embedded", file: episodeVideo },
			{ source: "driveThumb", file: episodeVideo },
			{ source: "frame", file: episodeVideo },
			{ source: "gradient" }
		])
	})

	it("sizes slots, with a larger still on 2x screens", () => {
		expect(artBox("p", 2)).toEqual(ART_BOXES.p)
		expect(artBox("s", 1)).toEqual({ width: 480, height: 270 })
		expect(artBox("s", 2)).toEqual(STILL_BOX_2X)
	})

	it("colours a typographic card the same for the same title", () => {
		expect(titleColour("Up")).toEqual(titleColour("Up"))
		expect(titleColour("Up")).not.toEqual(titleColour("Down"))
		expect(Math.max(...titleColour("Up"))).toBeLessThan(160)
	})
})

function probe(cover: VideoProbe["cover"]): VideoProbe {
	return {
		v: 1,
		container: "mkv",
		durationSec: 60,
		index: "cues",
		video: {
			codec: "h264",
			fourcc: "avc1",
			codecString: "avc1.64001f",
			w: 1280,
			h: 720,
			par: 1,
			rotation: 0,
			bitDepth: 8,
			transfer: "sdr"
		},
		audio: [],
		subs: [],
		chapters: [],
		cover,
		coverLand: null,
		partial: false
	}
}

function loader(overrides: Partial<ArtLoaderDeps> = {}) {
	const { kv } = memoryKv()
	const records = new RecordStore<ArtRecord>({
		prefix: "videos.v1.art.",
		schema: artRecordSchema,
		maxRecords: 100,
		memoryRecords: 100,
		writesPerSweep: 1000,
		scope: "test",
		kv
	})
	const stored = new Map<string, Uint8Array>()
	const deps: ArtLoaderDeps = {
		makeArt: vi.fn((): Promise<SdkThumbnailResult | { type: "notImage" }> =>
			Promise.resolve({
				type: "thumbnail" as const,
				bytes: new Uint8Array([1, 2]),
				width: 342,
				height: 500,
				fromEmbeddedPreview: false
			})
		),
		cancel: vi.fn(),
		storeArt: vi.fn((key: string, bytes: Uint8Array) => {
			stored.set(key, bytes)

			return Promise.resolve()
		}),
		readArt: vi.fn((key: string) => {
			const bytes = stored.get(key)

			return Promise.resolve(bytes === undefined ? null : new Blob([bytes as Uint8Array<ArrayBuffer>]))
		}),
		readDriveThumb: vi.fn(() => Promise.resolve(null)),
		sampleImage: vi.fn(() => Promise.resolve({ luma: 100, dominant: [10, 20, 30] as [number, number, number] })),
		measureImage: vi.fn(() =>
			Promise.resolve({ width: 384, height: 216, sample: { luma: 100, dominant: [1, 2, 3] as [number, number, number] } })
		),
		probes: { request: vi.fn(() => Promise.resolve(mergeProbe(null, "full", probe(null), 1))) },
		frames: { generate: vi.fn(() => Promise.resolve(null)) },
		records,
		reads: { run: (_rank, read) => read() },
		devicePixelRatio: () => 1,
		now: () => 5,
		...overrides
	}

	return { loader: new ArtLoader(deps), deps, records, stored }
}

const VIDEO: ArtFile = { uuid: "video", size: 1_000_000, file: { uuid: "video" } as unknown as AnyFile }
const POSTER: ArtFile = { uuid: "poster", size: 50_000, file: { uuid: "poster" } as unknown as AnyFile }

function files(...entries: ArtFile[]): (file: number) => ArtFile | null {
	return file => entries[file] ?? null
}

describe("ArtLoader", () => {
	it("makes a sidecar poster once, stores it with its record, and serves it from the store after", async () => {
		const { loader: art, deps, records, stored } = loader()
		const request = {
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }] satisfies ArtCandidate[],
			slot: "p" as const,
			title: "Up",
			resolveFile: files(POSTER)
		}
		const first = await art.load(request)

		expect(first).toMatchObject({ type: "image", source: "sidecar", width: 342, height: 500, fit: "cover", dominant: [10, 20, 30] })
		expect(deps.makeArt).toHaveBeenCalledWith(POSTER.file, { offset: 0, size: POSTER.size }, ART_BOXES.p, expect.any(String))
		expect(stored.has("poster.p")).toBe(true)
		expect(records.peek("poster.p")).toEqual({ slot: "p", source: "sidecar", fit: "cover", w: 342, h: 500, dom: [10, 20, 30], at: 5 })

		const second = await art.load(request)

		expect(second).toMatchObject({ type: "image", source: "sidecar" })
		expect(deps.makeArt).toHaveBeenCalledTimes(1)
	})

	it("takes a file resolved asynchronously, as the library index hands them out", async () => {
		const { loader: art, deps } = loader()
		const result = await art.load({
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
			slot: "p",
			title: "Up",
			resolveFile: file => Promise.resolve(files(POSTER)(file))
		})

		expect(result).toMatchObject({ type: "image", source: "sidecar" })
		expect(deps.makeArt).toHaveBeenCalledWith(POSTER.file, { offset: 0, size: POSTER.size }, ART_BOXES.p, expect.any(String))
	})

	it("skips a sidecar past the size gate without reading it", async () => {
		const { loader: art, deps } = loader()
		const result = await art.load({
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
			slot: "p",
			title: "Up",
			resolveFile: files({ ...POSTER, size: ART_SOURCE_MAX_BYTES + 1 })
		})

		expect(result).toEqual({ type: "typographic", title: "Up", colour: titleColour("Up") })
		expect(deps.makeArt).not.toHaveBeenCalled()
	})

	it("reuses the drive's thumbnail without any download, unless it is black", async () => {
		const blob = new Blob([new Uint8Array([9])])
		const { loader: art, deps } = loader({ readDriveThumb: vi.fn(() => Promise.resolve(blob)) })
		const request = {
			candidates: [{ source: "driveThumb", file: 0 }, { source: "gradient" }] satisfies ArtCandidate[],
			slot: "s" as const,
			title: "Up",
			resolveFile: files(VIDEO)
		}

		expect(await art.load(request)).toEqual({
			type: "image",
			blob,
			source: "driveThumb",
			width: 384,
			height: 216,
			fit: "cover",
			dominant: [1, 2, 3]
		})
		expect(deps.probes.request).not.toHaveBeenCalled()
		expect(deps.makeArt).not.toHaveBeenCalled()
		expect(deps.storeArt).not.toHaveBeenCalled()

		const dark = loader({
			readDriveThumb: vi.fn(() => Promise.resolve(blob)),
			measureImage: vi.fn(() =>
				Promise.resolve({ width: 384, height: 216, sample: { luma: 3, dominant: [0, 0, 0] as [number, number, number] } })
			)
		})

		expect((await dark.loader.load(request))?.type).toBe("gradient")
	})

	it("asks for the embedded cover only through a cover probe, and remembers a cover that is no image", async () => {
		const withCover = mergeProbe(null, "full", probe({ offset: 10, size: 20, mime: "image/*" }), 1)
		const {
			loader: art,
			deps,
			records
		} = loader({
			probes: { request: vi.fn(() => Promise.resolve(withCover)) },
			makeArt: vi.fn(() => Promise.resolve({ type: "notImage" as const }))
		})
		const request = {
			candidates: [{ source: "embedded", file: 0 }, { source: "typographic" }] satisfies ArtCandidate[],
			slot: "p" as const,
			title: "Up",
			resolveFile: files(VIDEO)
		}

		expect((await art.load(request))?.type).toBe("typographic")
		expect(deps.probes.request).toHaveBeenCalledWith({ uuid: "video", file: VIDEO.file }, "cover", expect.anything())
		expect(deps.makeArt).toHaveBeenCalledTimes(1)

		await art.load(request)
		expect(deps.makeArt).toHaveBeenCalledTimes(1)
		expect(deps.makeArt).toHaveBeenCalledWith(VIDEO.file, { offset: 10, size: 20, mime: "image/*" }, ART_BOXES.p, expect.any(String))
		// The verdict outlives the session.
		expect(records.peek("video.p")).toMatchObject({ source: "embedded", miss: true })
	})

	it("passes over the embedded cover when the probe found none", async () => {
		const { loader: art, deps } = loader()

		await art.load({
			candidates: [{ source: "embedded", file: 0 }, { source: "gradient" }],
			slot: "p",
			title: "Up",
			resolveFile: files(VIDEO)
		})

		expect(deps.makeArt).not.toHaveBeenCalled()
	})

	it("stores a generated frame under the video's uuid", async () => {
		const frame = { bytes: new Uint8Array([4, 5]), width: 480, height: 270, dominant: [7, 8, 9] as [number, number, number] }
		const { loader: art, deps, records } = loader({ frames: { generate: vi.fn(() => Promise.resolve(frame)) } })
		const result = await art.load({
			candidates: [{ source: "frame", file: 0 }, { source: "gradient" }],
			slot: "s",
			title: "Up",
			resolveFile: files(VIDEO)
		})

		expect(result).toMatchObject({ type: "image", source: "frame", width: 480, height: 270, fit: "cover", dominant: [7, 8, 9] })
		expect(deps.probes.request).toHaveBeenCalledWith({ uuid: "video", file: VIDEO.file }, "card", expect.anything())
		expect(records.peek("video.sf")?.source).toBe("frame")
	})

	it("leaves a failing source alone after three tries", async () => {
		const { loader: art, deps } = loader({ makeArt: vi.fn(() => Promise.reject(new Error("download failed"))) })
		const request = {
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }] satisfies ArtCandidate[],
			slot: "p" as const,
			title: "Up",
			resolveFile: files(POSTER)
		}

		for (let i = 0; i < 4; i++) {
			expect((await art.load(request))?.type).toBe("typographic")
		}

		expect(deps.makeArt).toHaveBeenCalledTimes(3)
	})

	it("keeps a frame the player captured in its own entry, standing in for a generated one", async () => {
		const { loader: art, deps, records, stored } = loader()

		stored.set("video.s", new Uint8Array([1]))
		records.put("video.s", { slot: "s", source: "embedded", fit: "cover", w: 480, h: 270, dom: [1, 1, 1], at: 1 })
		stored.set("video.sc", new Uint8Array([2]))
		records.put("video.sc", { slot: "s", source: "capture", fit: "cover", w: 480, h: 270, dom: [2, 2, 2], at: 2 })

		const result = await art.load({ candidates: [{ source: "frame", file: 0 }], slot: "s", title: "Up", resolveFile: files(VIDEO) })

		expect(result).toMatchObject({ type: "image", source: "frame", dominant: [2, 2, 2] })
		expect(deps.frames.generate).not.toHaveBeenCalled()
		expect(records.peek("video.s")?.source).toBe("embedded")
	})

	it("never starts a decode for a requester that gave up while it queued for a read", async () => {
		const queued = Promise.withResolvers<undefined>()
		const entered = Promise.withResolvers<undefined>()
		const { loader: art, deps } = loader({
			reads: {
				run: async (_rank, read) => {
					entered.resolve(undefined)
					await queued.promise

					return read()
				}
			}
		})
		const controller = new AbortController()
		const loading = art.load({
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
			slot: "p",
			title: "Up",
			resolveFile: files(POSTER),
			signal: controller.signal
		})

		await entered.promise
		controller.abort()
		queued.resolve(undefined)

		expect(await loading).toBeNull()
		expect(deps.makeArt).not.toHaveBeenCalled()
		expect(deps.cancel).toHaveBeenCalledTimes(1)
	})

	it("makes at most two images at once, and shares one between requests for the same image", async () => {
		const pending: (() => void)[] = []
		const { loader: art, deps } = loader({
			makeArt: vi.fn(
				() =>
					new Promise<SdkThumbnailResult>(resolve => {
						pending.push(() => {
							resolve({ type: "thumbnail", bytes: new Uint8Array([1]), width: 342, height: 513, fromEmbeddedPreview: false })
						})
					})
			)
		})
		const posters = ["a", "b", "c"].map(uuid => ({ ...POSTER, uuid }))
		const loads = [...posters, posters[0] ?? POSTER].map(poster =>
			art.load({ candidates: [{ source: "sidecar", file: 0 }], slot: "p", title: "Up", resolveFile: files(poster) })
		)

		await vi.waitFor(() => {
			expect(deps.makeArt).toHaveBeenCalledTimes(2)
		})
		pending.shift()?.()
		await vi.waitFor(() => {
			expect(deps.makeArt).toHaveBeenCalledTimes(3)
		})

		for (const finish of pending) {
			finish()
		}

		const [first, , , again] = await Promise.all(loads)

		expect(again).toBe(first)
		expect(deps.makeArt).toHaveBeenCalledTimes(3)
	})

	it("answers null once its signal has aborted", async () => {
		const { loader: art } = loader()
		const controller = new AbortController()

		controller.abort()

		expect(
			await art.load({
				candidates: [{ source: "typographic" }],
				slot: "p",
				title: "Up",
				resolveFile: files(),
				signal: controller.signal
			})
		).toBeNull()
	})
})

describe("ArtLoader.loadLocal", () => {
	const uuids =
		(...entries: ArtFile[]) =>
		(file: number) =>
			entries[file]?.uuid ?? null
	const POSTER_CANDIDATES = [
		{ source: "sidecar", file: 0 },
		{ source: "embedded", file: 1 },
		{ source: "driveThumb", file: 1 },
		{ source: "typographic" },
		{ source: "frame", file: 1 }
	] satisfies ArtCandidate[]

	it("asks for a download at the first candidate that needs one, before any lower-precedence local image", async () => {
		const { loader: art, deps } = loader({
			readDriveThumb: vi.fn(() => Promise.resolve(new Blob([new Uint8Array([1])])))
		})
		const resolveFile = vi.fn(files(POSTER, VIDEO))

		expect(
			await art.loadLocal({ candidates: POSTER_CANDIDATES, slot: "p", title: "Up", resolveFile, fileUuid: uuids(POSTER, VIDEO) })
		).toBe("network")
		expect(deps.readDriveThumb).not.toHaveBeenCalled()
		expect(resolveFile).not.toHaveBeenCalled()
		expect(deps.makeArt).not.toHaveBeenCalled()
	})

	it("serves a stored sidecar ahead of the drive thumbnail, and passes over settled misses to the local thumbnail", async () => {
		const { loader: art, deps } = loader({
			readDriveThumb: vi.fn(() => Promise.resolve(new Blob([new Uint8Array([1])])))
		})
		const resolveFile = files(POSTER, VIDEO)

		await art.load({ candidates: POSTER_CANDIDATES, slot: "p", title: "Up", resolveFile })

		expect(
			await art.loadLocal({ candidates: POSTER_CANDIDATES, slot: "p", title: "Up", resolveFile, fileUuid: uuids(POSTER, VIDEO) })
		).toMatchObject({
			type: "image",
			source: "sidecar"
		})

		const missing = loader({
			readDriveThumb: vi.fn(() => Promise.resolve(new Blob([new Uint8Array([1])]))),
			makeArt: vi.fn(() => Promise.resolve({ type: "notImage" as const }))
		})

		await missing.loader.load({
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
			slot: "p",
			title: "Up",
			resolveFile
		})

		expect(
			await missing.loader.loadLocal({
				candidates: [{ source: "sidecar", file: 0 }, { source: "driveThumb", file: 1 }, { source: "typographic" }],
				slot: "p",
				title: "Up",
				resolveFile,
				fileUuid: uuids(POSTER, VIDEO)
			})
		).toMatchObject({ type: "image", source: "driveThumb" })
		expect(deps.makeArt).toHaveBeenCalledTimes(1)
	})

	it("ends at a synthetic card when nothing ahead of it needs a download", async () => {
		const { loader: art } = loader()

		expect(
			await art.loadLocal({
				candidates: [{ source: "driveThumb", file: 0 }, { source: "typographic" }, { source: "frame", file: 0 }],
				slot: "p",
				title: "Up",
				resolveFile: files(VIDEO),
				fileUuid: uuids(VIDEO)
			})
		).toMatchObject({ type: "typographic" })
	})
})
