import { describe, expect, it, vi } from "vitest"
import { NO_FILE, type VideoLibrary } from "@filen/shared/videos"
import type { TFunction } from "i18next"
import type { VideoProbe } from "@/lib/media/probe/types"

const { sdkApi, library } = vi.hoisted(() => ({
	sdkApi: { downloadFileBytes: vi.fn() },
	library: { playableFiles: vi.fn() }
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi }))
vi.mock("@/features/videos/lib/index", () => ({ videoLibrary: library }))
vi.mock("@/lib/storage/adapter", () => ({
	kvGetJson: () => Promise.resolve(null),
	kvSetJson: () => Promise.resolve(),
	kvKeys: () => Promise.resolve([]),
	kvEntriesJson: () => Promise.resolve([]),
	kvDelete: () => Promise.resolve()
}))

import { defaultAudioIndex } from "@/features/videos/lib/player/audioChoice.logic"
import { resolveSubtitleChoice, type PlayerSubtitleTrack } from "@/features/videos/lib/player/subtitles.logic"
import { NFO_MAX_BYTES, readNfo } from "@/features/videos/lib/nfoReader"
import { countedSeasons, detailHeaderMode, seasonLayout, timeLeftText } from "@/features/videos/lib/titleDetail.logic"
import { artCandidates, probeLandscapeCover } from "@/features/videos/lib/artwork.logic"
import type { VideoSeasonView } from "@/workers/videoIndex/view.logic"
import { title } from "@/tests/support/videoHome"

function audio(codec: string, lang: string, def: boolean): VideoProbe["audio"][number] {
	return { id: 1, codec: codec as never, codecString: null, ch: 6, rate: 48000, lang, name: null, def, forced: false, commentary: false }
}

function track(id: string, lang: string): PlayerSubtitleTrack {
	return { id, lang, title: "", forced: false, hi: false, isDefault: false, external: true, file: null, format: "srt", unavailable: null }
}

describe("title page review fixes", () => {
	it("starts on the audio track the player would (U3): a natively playable track over a default one needing conversion", () => {
		const probe = { audio: [audio("ac3", "eng", true), audio("aac", "eng", false)] }

		expect(defaultAudioIndex({ probe, rememberedLang: null, userLanguages: ["en"] })).toBe(1)
		expect(
			defaultAudioIndex({
				probe: { audio: [audio("aac", "eng", true), audio("aac", "deu", false)] },
				rememberedLang: "deu",
				userLanguages: ["en"]
			})
		).toBe(1)
	})

	it("falls back to Smart for a subtitle id the item doesn't have (U6)", () => {
		const tracks = [track("s:en", "eng"), track("s:de", "deu")]
		const base = { tracks, audioLang: "deu", userLanguages: ["en"], remembered: null }

		expect(resolveSubtitleChoice({ ...base, choice: { kind: "track", id: "s:gone" } })?.id).toBe(
			resolveSubtitleChoice({ ...base, choice: { kind: "auto" } })?.id
		)
		expect(resolveSubtitleChoice({ ...base, choice: { kind: "track", id: "s:de" } })?.id).toBe("s:de")
		expect(resolveSubtitleChoice({ ...base, choice: { kind: "off" } })).toBeNull()
	})

	it("refuses an NFO over 256 KiB by the index's size, without asking for the file (U8)", async () => {
		await expect(readNfo({ stableUuid: "big-s", uuid: "big-u", name: "movie.nfo", size: NFO_MAX_BYTES + 1 })).resolves.toBeNull()
		expect(library.playableFiles).not.toHaveBeenCalled()
		expect(sdkApi.downloadFileBytes).not.toHaveBeenCalled()
		expect(NFO_MAX_BYTES).toBe(256 * 1024)
	})

	it("counts and marks a specials-only series by its specials (U10)", () => {
		const specials: VideoSeasonView = { number: 0, dirUuid: "", poster: null, backdrop: null, nfo: null, episodes: [] }
		const withEpisodes = {
			...specials,
			episodes: [{ season: 0, episode: 1, episodeEnd: 1, date: 0, title: "", dateAdded: 0, versions: [] }]
		}
		const layout = seasonLayout([withEpisodes])

		expect(layout.regular).toHaveLength(0)
		expect(countedSeasons(layout)).toEqual([withEpisodes])
		expect(countedSeasons(seasonLayout([withEpisodes, { ...withEpisodes, number: 1 }])).map(season => season.number)).toEqual([1])
	})

	it("words time left the same on every title page (U11)", () => {
		const t = ((key: string, values: Record<string, unknown>) =>
			key === "detailTimeLeft"
				? `${String(values["time"])} left`
				: key === "videosHomeRuntimeHours"
					? `${String(values["hours"])}h ${String(values["minutes"])}m`
					: `${String(values["minutes"])}m`) as unknown as TFunction<"videos">

		expect(timeLeftText(21 * 60, t)).toBe("21m left")
		expect(timeLeftText(91 * 60, t)).toBe("1h 31m left")
	})

	it("lets a Matroska landscape cover drive the backdrop once the probe records it (C1)", () => {
		const library: VideoLibrary = {
			classifierVersion: 1,
			videos: [
				{ file: 0, container: "mkv", unsupported: false, height: 0, subtitles: [], nfo: NO_FILE, thumb: NO_FILE, edl: NO_FILE }
			],
			movies: [
				{
					id: "m",
					title: "M",
					sortName: "M",
					year: 0,
					mergeKey: "",
					roots: [0],
					nfo: NO_FILE,
					art: { poster: NO_FILE, backdrop: NO_FILE, logo: NO_FILE },
					versions: [{ label: "", parts: [0] }],
					extras: [],
					dateAdded: 0,
					disc: false
				}
			],
			series: [],
			others: []
		}
		const movie = library.movies[0]

		if (movie === undefined) {
			throw new Error("fixture")
		}

		expect(artCandidates(library, { kind: "movie", movie }, "b").map(candidate => candidate.source)).toEqual([
			"embedded",
			"frame",
			"gradient"
		])

		const probe = { cover: null } as unknown as VideoProbe

		expect(probeLandscapeCover(probe)).toBeNull()
		expect(probeLandscapeCover({ ...probe, coverLand: { offset: 10, size: 20, mime: "image/jpeg" } } as VideoProbe)).toEqual({
			offset: 10,
			size: 20,
			mime: "image/jpeg"
		})
		expect(detailHeaderMode(title("m"), true)).toBe("backdrop")
		expect(detailHeaderMode(title("m"), false)).toBe("band")
	})
})
