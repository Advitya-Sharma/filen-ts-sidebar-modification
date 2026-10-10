import { beforeEach, describe, expect, it, vi } from "vitest"
import { CueBuilder, MAX_CONCURRENT_CUES, type VideoProbe } from "@filen/shared/videos"
import type { VideoTitleView } from "@/workers/videoIndex/view.logic"
import type { PlayerSubtitleTrack } from "@/features/videos/lib/player/subtitles.logic"

const mocks = vi.hoisted(() => ({
	download: vi.fn(),
	playableFiles: vi.fn(),
	openRangePort: vi.fn(() => Promise.resolve()),
	cancel: vi.fn(),
	register: vi.fn((_file: unknown, _name: string, _type: string, id: string) => Promise.resolve(`/sw/${id}`)),
	release: vi.fn(),
	warmCancel: vi.fn(),
	warmHandOver: vi.fn()
}))

vi.mock("@/lib/sdk/client", () => ({
	sdkApi: { downloadFileBytes: mocks.download, openRangePort: mocks.openRangePort, cancelPreviewDownload: mocks.cancel }
}))
vi.mock("@/features/videos/lib/index", () => ({ videoLibrary: { playableFiles: mocks.playableFiles } }))
vi.mock("@/lib/storage/adapter", () => ({ kvGetJson: () => Promise.resolve(null), kvSetJson: () => Promise.resolve() }))
vi.mock("@/features/preview/lib/previewStream", () => ({
	previewStreamUrl: mocks.register,
	releasePreviewStream: mocks.release,
	holdPreviewStreamAlive: () => () => undefined,
	onPreviewStreamGone: () => () => undefined,
	previewStreamStatus: () => Promise.resolve("present")
}))
vi.mock("@/lib/media/mseController", () => ({
	warmRemuxFile: () => Promise.resolve({ cancel: mocks.warmCancel, handOver: mocks.warmHandOver })
}))

import { nextMovieInDirectory } from "@/features/videos/lib/player/postPlay.logic"
import { MAX_KEPT_DECODED, MAX_KEPT_SUBTITLE_FILES, SidecarReader } from "@/features/videos/lib/player/sidecarReader"
import { SubtitleTrackPool, addedCueIndices } from "@/lib/media/subtitles/subtitleTracks"
import { cancelPrewarm, directSource, prewarmPlayback, remuxSource } from "@/lib/media/playback/sources"

function movie(id: string, sortName: string, roots: string[], extra: Partial<VideoTitleView> = {}): VideoTitleView {
	return {
		id,
		kind: "movie",
		title: sortName,
		sortName,
		year: 0,
		mergeKey: "",
		roots,
		dateAdded: 0,
		poster: null,
		backdrop: null,
		logo: null,
		primary: { stableUuid: `${id}-s`, uuid: `${id}-u`, name: "", size: 1 },
		container: "mp4",
		height: 0,
		unsupported: false,
		disc: false,
		versionCount: 1,
		extraCount: 0,
		seasonCount: 0,
		episodeCount: 0,
		...extra
	}
}

describe("movie post-play", () => {
	const titles = [
		movie("m10", "Movie 10", ["r1"]),
		movie("m2", "Movie 2", ["r1"]),
		movie("m1", "Movie 1", ["r1"]),
		movie("x", "Alien", ["r2"]),
		movie("bad", "Movie 3", ["r1"], { unsupported: true }),
		movie("s", "Show", ["r1"], { kind: "series" })
	]

	it("offers the next unwatched movie of the same library directory in natural order, wrapping", () => {
		expect(nextMovieInDirectory({ titles, currentId: "m2", played: () => false })?.id).toBe("m10")
		expect(nextMovieInDirectory({ titles, currentId: "m10", played: () => false })?.id).toBe("m1")
		expect(nextMovieInDirectory({ titles, currentId: "m1", played: title => title.id === "m2" })?.id).toBe("m10")
		expect(nextMovieInDirectory({ titles, currentId: "m1", played: title => title.id !== "m1" })).toBeNull()
		expect(nextMovieInDirectory({ titles, currentId: "x", played: () => false })).toBeNull()
	})
})

// A text track as browsers have it: a disabled track exposes no cues.
class FakeTrack {
	public mode: TextTrackMode = "disabled"
	private readonly list: unknown[] = []

	public get cues(): unknown[] | null {
		return this.mode === "disabled" ? null : this.list
	}

	public added = 0

	public addCue(cue: unknown): void {
		this.added++
		this.list.push(cue)
	}

	public removeCue(cue: unknown): void {
		this.list.splice(this.list.indexOf(cue), 1)
	}

	public get stored(): number {
		return this.list.length
	}

	public addEventListener = (): undefined => undefined
	public removeEventListener = (): undefined => undefined
}

function cueStore(cues: readonly [number, string][]) {
	const builder = new CueBuilder()

	for (const [start, text] of cues) {
		builder.add(start, start + 1, text)
	}

	return builder.finish()
}

describe("SubtitleTrackPool", () => {
	const tracks: FakeTrack[] = []
	const video = {
		addTextTrack: () => {
			const track = new FakeTrack()

			tracks.push(track)

			return track
		}
	}

	beforeEach(() => {
		tracks.length = 0
		vi.stubGlobal(
			"VTTCue",
			class {
				public line: number | "auto" = "auto"
				public snapToLines = true
				public align = "center"
				public id = ""
				public startTime: number
				public endTime: number
				public text: string
				public constructor(startTime: number, endTime: number, text: string) {
					this.startTime = startTime
					this.endTime = endTime
					this.text = text
				}
			}
		)
	})

	it("empties a track before disabling it, so reusing it never piles cues up", () => {
		const builder = new CueBuilder()

		builder.add(1, 2, "a")
		builder.add(3, 4, "b")

		const store = builder.finish()
		const pool = new SubtitleTrackPool(video as unknown as HTMLVideoElement)

		pool.show(store, 0)
		pool.show(store, 0.5)
		pool.show(store, 1)
		pool.show(null, 0)

		expect(tracks.every(track => track.stored === 0)).toBe(true)
		pool.show(store, 0)
		expect(tracks.reduce((sum, track) => sum + track.stored, 0)).toBe(2)
		pool.dispose()
		expect(tracks.every(track => track.stored === 0 && track.mode === "disabled")).toBe(true)
	})

	it("adds only the new cues when the cues just grew, and refills on a delay change or other cues", () => {
		const pool = new SubtitleTrackPool(video as unknown as HTMLVideoElement)
		const first = cueStore([
			[1, "a"],
			[5, "c"]
		])
		const grown = cueStore([
			[1, "a"],
			[3, "b"],
			[5, "c"],
			[7, "d"]
		])

		pool.show(first, 0)
		pool.show(grown, 0)

		expect(tracks).toHaveLength(1)
		expect(tracks[0]?.added).toBe(4)
		expect(tracks[0]?.stored).toBe(4)

		pool.show(grown, 0)
		expect(tracks[0]?.added).toBe(4)

		pool.show(grown, 0.5)
		expect(tracks.reduce((sum, track) => sum + track.stored, 0)).toBe(4)
		expect(tracks.reduce((sum, track) => sum + track.added, 0)).toBe(8)

		pool.show(cueStore([[2, "x"]]), 0.5)
		expect(tracks.reduce((sum, track) => sum + track.stored, 0)).toBe(1)
		pool.dispose()
	})

	it("never hands the browser more overlapping cues than show at once, also as embedded cues grow", () => {
		const pool = new SubtitleTrackPool(video as unknown as HTMLVideoElement)
		const overlapping = (from: number, count: number) => {
			const builder = new CueBuilder()

			for (let i = from; i < from + count; i++) {
				builder.add(i / 100, 10_800, `c${String(i)}`)
			}

			return builder.finish()
		}
		const current = () => (tracks.find(track => track.stored > 0)?.cues ?? []) as { startTime: number; endTime: number; text: string }[]
		const activeAt = (time: number) => current().filter(cue => cue.startTime <= time && cue.endTime > time)
		const mostAtOnce = () => Math.max(...current().map(cue => activeAt(cue.startTime).length))

		pool.show(overlapping(1, 1000), 0)
		expect(mostAtOnce()).toBe(MAX_CONCURRENT_CUES)
		expect(activeAt(30).map(cue => cue.text)).toEqual(
			Array.from({ length: MAX_CONCURRENT_CUES }, (_, i) => `c${String(1001 - MAX_CONCURRENT_CUES + i)}`)
		)

		// An earlier cue read later shortens the ones after it: the track is refilled rather than left overlapping.
		pool.show(overlapping(0, 1001), 0)
		expect(tracks.filter(track => track.stored > 0)).toHaveLength(1)
		expect(mostAtOnce()).toBe(MAX_CONCURRENT_CUES)
		expect(activeAt(5).map(cue => cue.text)).toEqual(
			Array.from({ length: MAX_CONCURRENT_CUES }, (_, i) => `c${String(501 - MAX_CONCURRENT_CUES + i)}`)
		)
		pool.dispose()
	})

	it("tells the added cues of a store that holds every earlier cue in order, and nothing else", () => {
		const base = cueStore([
			[1, "a"],
			[5, "c"]
		])

		expect(
			addedCueIndices(
				base,
				cueStore([
					[0, "z"],
					[1, "a"],
					[5, "c"]
				])
			)
		).toEqual([0])
		expect(addedCueIndices(base, base)).toEqual([])
		expect(
			addedCueIndices(
				base,
				cueStore([
					[1, "a"],
					[5, "changed"]
				])
			)
		).toBeNull()
		expect(addedCueIndices(base, cueStore([[1, "a"]]))).toBeNull()
	})
})

describe("SidecarReader", () => {
	const track: PlayerSubtitleTrack = {
		id: "s:ru",
		lang: "rus",
		title: "",
		forced: false,
		hi: false,
		isDefault: false,
		external: true,
		file: { stableUuid: "ru", uuid: "ru-u", name: "a.ru.srt", size: 100 },
		format: "srt",
		unavailable: null
	}

	beforeEach(() => {
		mocks.download.mockReset()
		mocks.playableFiles.mockReset()
	})

	it("downloads a subtitle file once and decodes it again for each encoding picked", async () => {
		mocks.playableFiles.mockResolvedValue([{ uuid: "ru-u", size: 100n }])
		mocks.download.mockResolvedValue(
			new Uint8Array([
				49, 10, 48, 48, 58, 48, 48, 58, 48, 49, 44, 48, 48, 48, 32, 45, 45, 62, 32, 48, 48, 58, 48, 48, 58, 48, 50, 44, 48, 48, 48,
				10, 0xcf, 0xf0, 0xe8, 10
			])
		)

		const reader = new SidecarReader()
		const auto = await reader.subtitle(track, null)
		const latin = await reader.subtitle(track, "windows-1252")
		const again = await reader.subtitle(track, null)

		expect(mocks.download).toHaveBeenCalledTimes(1)
		expect(auto.status === "ok" ? auto.cues.texts[0] : null).toBe("При")
		expect(latin.status === "ok" ? latin.cues.texts[0] : null).toBe("Ïðè")
		expect(again).toBe(auto)
	})

	it("keeps only the latest subtitle files across a binge, and no .edl bytes", async () => {
		const srt = new TextEncoder().encode("1\n00:00:01,000 --> 00:00:02,000\nHi\n")
		const episode = (index: number): PlayerSubtitleTrack => ({
			...track,
			id: `s:e${String(index)}`,
			file: { stableUuid: `e${String(index)}`, uuid: `e${String(index)}-u`, name: "a.srt", size: 100 }
		})

		mocks.playableFiles.mockResolvedValue([{ uuid: "x", size: 100n }])
		mocks.download.mockResolvedValue(srt)

		const reader = new SidecarReader()
		const internals = reader as unknown as { bytes: Map<string, unknown>; decoded: Map<string, unknown>; edls: Map<string, unknown> }

		for (let index = 0; index < 30; index++) {
			await reader.subtitle(episode(index), null)
			await reader.edl({ stableUuid: `edl${String(index)}`, uuid: `edl${String(index)}-u`, name: "a.edl", size: 10 })
		}

		expect(internals.bytes.size).toBe(MAX_KEPT_SUBTITLE_FILES)
		expect(internals.decoded.size).toBe(MAX_KEPT_DECODED)
		expect(internals.edls.size).toBeLessThanOrEqual(4)

		// The item shown now is still at hand: another encoding is decoded from the kept bytes.
		mocks.download.mockClear()
		await reader.subtitle(episode(29), "windows-1252")
		expect(mocks.download).not.toHaveBeenCalled()
	})

	it("refuses by the file's size on the drive now, not the indexed one", async () => {
		mocks.playableFiles.mockResolvedValue([{ uuid: "ru-u", size: 9n * 1024n * 1024n }])

		const result = await new SidecarReader().subtitle(track, null)

		expect(result.status).toBe("tooLarge")
		expect(mocks.download).not.toHaveBeenCalled()
	})
})

const FILE = { uuid: "next-u", size: 1000n } as const

function probeOf(container: VideoProbe["container"]): VideoProbe {
	return {
		v: 1,
		container,
		durationSec: 60,
		index: "cues",
		video: null,
		audio: [],
		subs: [],
		chapters: [],
		cover: null,
		coverLand: null,
		partial: false
	}
}

describe("prewarm", () => {
	beforeEach(() => {
		cancelPrewarm()
		mocks.warmCancel.mockReset()
		mocks.warmHandOver.mockReset()
		mocks.release.mockReset()
		mocks.cancel.mockReset()
	})

	it("opens a Matroska file in the remux worker once, and the next session takes it over", async () => {
		const file = FILE as never

		prewarmPlayback({ file, name: "next.mkv", probe: probeOf("mkv"), idrOnlyStart: false })
		prewarmPlayback({ file, name: "next.mkv", probe: probeOf("mkv"), idrOnlyStart: false })
		await Promise.resolve()

		expect(mocks.openRangePort).toHaveBeenCalledTimes(1)

		const taken = remuxSource(file)

		expect(taken.warmKey).toBeTypeOf("string")
		await new Promise(resolve => setTimeout(resolve, 0))
		expect(mocks.warmHandOver).toHaveBeenCalledTimes(1)
		expect(mocks.warmCancel).not.toHaveBeenCalled()
		expect(remuxSource(file).warmKey).toBeUndefined()
	})

	it("lets the session that takes a warm wait until the warm has reached the remux worker", async () => {
		const file = FILE as never

		prewarmPlayback({ file, name: "next.mkv", probe: probeOf("mkv"), idrOnlyStart: false })

		const taken = remuxSource(file)

		expect(mocks.warmHandOver).not.toHaveBeenCalled()
		expect(taken.warmPort).toBeInstanceOf(Promise)
		await taken.warmPort
		await taken.warmed
		expect(mocks.warmHandOver).toHaveBeenCalledTimes(1)
		expect(remuxSource(file).warmed).toBeUndefined()
		expect(remuxSource(file).warmPort).toBeUndefined()
	})

	it("lets a warmed file go when nobody takes it", async () => {
		prewarmPlayback({ file: FILE as never, name: "next.mkv", probe: probeOf("mkv"), idrOnlyStart: false })
		cancelPrewarm()
		await new Promise(resolve => setTimeout(resolve, 0))

		expect(mocks.cancel).toHaveBeenCalledTimes(1)
		expect(mocks.warmCancel).toHaveBeenCalledTimes(1)
	})

	it("registers a directly played file's stream up front and hands the registration over", () => {
		prewarmPlayback({ file: FILE as never, name: "next.mp4", probe: probeOf("mp4"), idrOnlyStart: false })

		const id = mocks.register.mock.calls.at(-1)?.[3]
		const source = directSource({ file: FILE as never, name: "next.mp4", contentType: "video/mp4" })

		expect(source.presetId).toBe(id)
		cancelPrewarm()
		expect(mocks.release).not.toHaveBeenCalled()
	})

	it("streams a public link's file through the anon registration, never taking a prewarmed one", async () => {
		prewarmPlayback({ file: FILE as never, name: "next.mp4", probe: probeOf("mp4"), idrOnlyStart: false })

		const source = directSource({ file: FILE as never, name: "next.mp4", contentType: "video/mp4", access: "anon" })

		expect(source.presetId).toBeUndefined()
		await source.register("anon-id")
		expect(mocks.register).toHaveBeenLastCalledWith(FILE, "next.mp4", "video/mp4", "anon-id", { access: "anon" })
		cancelPrewarm()
	})
})
