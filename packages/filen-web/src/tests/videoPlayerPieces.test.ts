import { beforeEach, describe, expect, it, vi } from "vitest"
import { CueBuilder } from "@filen/shared/videos"
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
import { SidecarReader } from "@/features/videos/lib/player/sidecarReader"
import { SubtitleTrackPool } from "@/lib/media/subtitles/subtitleTracks"
import { cancelPrewarm, directSource, prewarmPlayback, remuxSource } from "@/lib/media/playback/sources"
import type { VideoProbe } from "@/lib/media/probe/types"

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

	public addCue(cue: unknown): void {
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

describe("SubtitleTrackPool", () => {
	it("empties a track before disabling it, so reusing it never piles cues up", () => {
		const tracks: FakeTrack[] = []
		const video = {
			addTextTrack: () => {
				const track = new FakeTrack()

				tracks.push(track)

				return track
			}
		}

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
