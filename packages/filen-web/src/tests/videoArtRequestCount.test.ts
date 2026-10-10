import { beforeEach, describe, expect, it, vi } from "vitest"
import type { AnyFile } from "@filen/sdk-rs"
import {
	classifyVideoLibrary,
	type LibraryScanFile,
	BLOCK_BYTES,
	type ReadRange,
	CARD_BLOCK_READS,
	probeSource,
	probeVideo,
	type ProbeTier
} from "@filen/shared/videos"

// Request counts for the Videos probe and artwork paths: what browsing costs in SDK calls and block reads.

const { sdkApi, kvRows, driveThumbs } = vi.hoisted(() => ({
	sdkApi: {
		readVideoProbe: vi.fn(),
		makeSdkThumbnail: vi.fn(),
		makeEmbeddedCoverArt: vi.fn(),
		openRangePort: vi.fn(),
		storeArt: vi.fn(),
		cancelPreviewDownload: vi.fn(),
		cancelPreviewDownloadIfEarly: vi.fn()
	},
	kvRows: new Map<string, unknown>(),
	driveThumbs: new Map<string, Blob>()
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi }))
vi.mock("@/features/videos/workers/frame.worker.ts?worker", () => ({ default: vi.fn() }))
vi.mock("@/features/drive/lib/thumbCache", () => ({ readThumbnailBlob: (uuid: string) => Promise.resolve(driveThumbs.get(uuid) ?? null) }))
vi.mock("@/lib/storage/adapter", async () => {
	const { decodeValue } = await import("@/lib/storage/decode")

	return {
		kvGetJson: (key: string, schema: Parameters<typeof decodeValue>[1]) =>
			Promise.resolve(kvRows.has(key) ? decodeValue(kvRows.get(key), schema, "kv", key) : null),
		kvSetJson: (key: string, value: unknown) => {
			kvRows.set(key, value)

			return Promise.resolve()
		},
		kvKeys: (prefix: string) => Promise.resolve([...kvRows.keys()].filter(key => key.startsWith(prefix))),
		kvEntriesJson: () => Promise.resolve([]),
		kvDelete: (key: string) => {
			kvRows.delete(key)

			return Promise.resolve()
		}
	}
})

import { probeService } from "@/features/videos/lib/probeService"
import { artLoader } from "@/features/videos/lib/artwork"
import { artCandidates } from "@/features/videos/lib/artwork.logic"
import { mergeProbe } from "@/features/videos/lib/probeService.logic"
import {
	FTYP,
	box,
	concat,
	ebmlHeader,
	element,
	f64,
	hdlr,
	mdhd,
	mvhd,
	segment,
	stringElement,
	tkhd,
	trak,
	uintElement,
	videoEntry
} from "@/tests/support/containerBytes"

function totalCalls(): number {
	return Object.values(sdkApi).reduce((sum, fn) => sum + fn.mock.calls.length, 0)
}

// A tail-moov MP4 (the header after a multi-block mdat) and a Matroska file, as real muxers lay them out.
function tailMoovMp4(): Uint8Array {
	const moov = box(
		"moov",
		mvhd({ timescale: 1000, duration: 60_000, nextTrackId: 2 }),
		trak(tkhd({ id: 1 }), [mdhd({}), hdlr("vide")], videoEntry("avc1", {}))
	)

	return concat(FTYP, box("mdat", new Uint8Array(3 * BLOCK_BYTES)), moov)
}

function matroska(): Uint8Array {
	return concat(
		ebmlHeader(),
		segment(
			element(0x1549a966, uintElement(0x2ad7b1, 1_000_000), element(0x4489, f64(60_000))),
			element(0x1654ae6b, element(0xae, uintElement(0xd7, 1), uintElement(0x83, 1), stringElement(0x86, "V_VP9"))),
			element(0x1f43b675, new Uint8Array(2 * BLOCK_BYTES))
		)
	)
}

const FILES = new Map<string, Uint8Array>()
const blockReads = new Map<string, number>()

// The worker op, run here over the bytes with the real probe, counting block downloads.
sdkApi.readVideoProbe.mockImplementation(async (_file: AnyFile, uuid: string, tier: ProbeTier) => {
	const bytes = FILES.get(uuid) ?? new Uint8Array()
	const readRange: ReadRange = (start, end) => {
		blockReads.set(uuid, (blockReads.get(uuid) ?? 0) + 1)

		return Promise.resolve(bytes.slice(start, end))
	}

	return probeVideo(probeSource(bytes.length, readRange, tier), tier)
})

function target(uuid: string): { uuid: string; file: AnyFile } {
	return { uuid, file: { uuid } as unknown as AnyFile }
}

beforeEach(() => {
	for (const fn of Object.values(sdkApi)) {
		fn.mockClear()
	}

	blockReads.clear()
})

describe("Videos browsing request counts", () => {
	it("reads nothing to classify a library and decide every title's artwork", () => {
		const files: LibraryScanFile[] = [
			["d1", "Up (2009).mkv"],
			["d1", "poster.jpg"],
			["s1", "poster.jpg"],
			["s1s1", "Show S01E01.mkv"],
			["s1s1", "Show S01E01-thumb.jpg"]
		].map(([parent, name], i) => ({ uuid: `f${String(i)}`, parent: parent ?? "", name: name ?? "", size: 10, timestamp: 1 }))
		const library = classifyVideoLibrary({
			roots: [
				{ uuid: "rm", name: "Movies", type: "movies" },
				{ uuid: "rs", name: "Shows", type: "shows" }
			],
			dirs: [
				{ uuid: "d1", parent: "rm", name: "Up (2009)" },
				{ uuid: "s1", parent: "rs", name: "Show" },
				{ uuid: "s1s1", parent: "s1", name: "Season 01" }
			],
			files
		})

		for (const slot of ["p", "b", "s"] as const) {
			for (const movie of library.movies) {
				artCandidates(library, { kind: "movie", movie }, slot)
			}

			for (const series of library.series) {
				artCandidates(library, { kind: "series", series }, slot)

				for (const season of series.seasons) {
					artCandidates(library, { kind: "season", series, season }, slot)
				}
			}
		}

		expect(totalCalls()).toBe(0)
	})

	it("probes a card within its block budget, once per file", async () => {
		FILES.set("mp4", tailMoovMp4())
		FILES.set("mkv", matroska())

		const mp4 = await probeService.request(target("mp4"), "card")
		const mkv = await probeService.request(target("mkv"), "card")

		expect(mp4?.probe?.container).toBe("mp4")
		expect(mkv?.probe?.container).toBe("mkv")
		expect(blockReads.get("mp4")).toBeLessThanOrEqual(CARD_BLOCK_READS)
		expect(blockReads.get("mkv")).toBe(1)
		expect(sdkApi.readVideoProbe).toHaveBeenCalledTimes(2)

		await probeService.request(target("mp4"), "card")
		await probeService.request(target("mkv"), "card")

		expect(sdkApi.readVideoProbe).toHaveBeenCalledTimes(2)
	})

	it("reuses the drive's thumbnail for a still without a single SDK call", async () => {
		kvRows.set("videos.v1.probe.episode", mergeProbe(null, "full", null, 1))
		driveThumbs.set("episode", new Blob([new Uint8Array([1, 2, 3])]))
		vi.stubGlobal(
			"createImageBitmap",
			vi.fn(() => Promise.resolve({ width: 384, height: 216, close: vi.fn() }))
		)
		vi.stubGlobal(
			"OffscreenCanvas",
			class {
				getContext() {
					return { drawImage: vi.fn(), getImageData: () => ({ data: new Uint8ClampedArray(64 * 36 * 4).fill(120) }) }
				}
			}
		)
		vi.stubGlobal("window", { devicePixelRatio: 1 })

		const result = await artLoader.load({
			candidates: [
				{ source: "embedded", file: 0 },
				{ source: "driveThumb", file: 0 },
				{ source: "frame", file: 0 },
				{ source: "gradient" }
			],
			slot: "s",
			title: "Show",
			resolveFile: () => ({ uuid: "episode", size: 50 * BLOCK_BYTES, file: { uuid: "episode" } as unknown as AnyFile })
		})

		expect(result).toMatchObject({ type: "image", source: "driveThumb", width: 384, height: 216 })
		expect(totalCalls()).toBe(0)

		vi.unstubAllGlobals()
	})
})
