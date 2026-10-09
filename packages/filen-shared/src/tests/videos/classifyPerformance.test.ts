import { describe, expect, it } from "vitest"
import {
	classifyVideoLibrary,
	classifyVideoLibrarySteps,
	type LibraryScanDir,
	type LibraryScanFile,
	type LibraryScanInput
} from "@filen/shared/videos"

const TOTAL_FILES = 50_000

// A large, realistic library: movie directories with sidecars and versions, shows with seasons and subtitles,
// phone clips and unrelated files.
function syntheticLibrary(): LibraryScanInput {
	const dirs: LibraryScanDir[] = []
	const files: LibraryScanFile[] = []
	let nextId = 0
	const dir = (parent: string, name: string): string => {
		const uuid = `d${nextId++}`

		dirs.push({
			uuid,
			parent,
			name
		})

		return uuid
	}
	const file = (parent: string, name: string): void => {
		files.push({
			uuid: `f${nextId++}`,
			parent,
			name,
			size: 1_500_000_000,
			timestamp: 1_700_000_000_000 + files.length
		})
	}
	const movies = dir("root-movies", "Movies")

	for (let m = 0; m < 3000; m++) {
		const title = `Movie Title Number ${m} (${1950 + (m % 70)})`
		const folder = dir(movies, title)

		file(folder, `${title}.mkv`)
		file(folder, `${title}.en.srt`)
		file(folder, `${title}.de.forced.srt`)
		file(folder, "poster.jpg")
		file(folder, "fanart.jpg")
		file(folder, "movie.nfo")

		if (m % 5 === 0) {
			file(folder, `${title} - 2160p.mkv`)
			file(folder, `${title}-trailer.mkv`)
		}
	}

	const shows = dir("root-shows", "Shows")

	for (let s = 0; files.length < 40_000; s++) {
		const show = dir(shows, `Show Name ${s} (${2000 + (s % 20)})`)

		file(show, "tvshow.nfo")
		file(show, "poster.jpg")

		for (let season = 1; season <= 4; season++) {
			const seasonDir = dir(show, `Season ${season}`)

			for (let episode = 1; episode <= 12; episode++) {
				const stem = `Show.Name.${s}.S${String(season).padStart(2, "0")}E${String(episode).padStart(2, "0")}.Episode.Title.1080p.WEB-DL`

				file(seasonDir, `${stem}.mkv`)
				file(seasonDir, `${stem}.en.srt`)
			}
		}
	}

	const phone = dir("root-other", "Camera")

	for (let day = 0; files.length < TOTAL_FILES; day++) {
		const folder = dir(phone, `2023-${String((day % 12) + 1).padStart(2, "0")}-${String((day % 28) + 1).padStart(2, "0")} Trip ${day}`)

		for (let clip = 0; clip < 20 && files.length < TOTAL_FILES; clip++) {
			file(folder, clip % 4 === 0 ? `IMG_${day}${clip}.jpg` : `VID_2023010${clip % 10}_1010${clip}.mp4`)
		}
	}

	return {
		roots: [
			{
				uuid: "root-movies",
				name: "Movies root",
				type: "auto"
			},
			{
				uuid: "root-shows",
				name: "Shows root",
				type: "auto"
			},
			{
				uuid: "root-other",
				name: "Phone",
				type: "auto"
			}
		],
		dirs,
		files
	}
}

describe("classifyVideoLibrary performance", () => {
	it(`classifies ${TOTAL_FILES} names well within budget`, () => {
		const input = syntheticLibrary()

		expect(input.files.length).toBe(TOTAL_FILES)

		// The first run includes regex compilation and JIT warm-up, which a worker pays once per session.
		const coldStarted = performance.now()

		classifyVideoLibrary(input)

		const cold = performance.now() - coldStarted
		const started = performance.now()
		const library = classifyVideoLibrary(input)
		const elapsed = performance.now() - started

		console.info(`classified ${TOTAL_FILES} names in ${elapsed.toFixed(1)} ms (first run ${cold.toFixed(1)} ms)`)

		expect(library.movies).toHaveLength(3000)
		expect(library.series.length).toBeGreaterThan(100)
		expect(library.others.length).toBeGreaterThan(1000)
		// The target is well under 500 ms; the bound is loose so slow CI machines do not flake.
		expect(elapsed).toBeLessThan(2500)
		expect(cold).toBeLessThan(5000)
	})
})

describe("classifyVideoLibrary in one flat directory", () => {
	it("matches title-named sidecars in a directory of 25k movies in about linear time, in short steps", () => {
		const files: LibraryScanFile[] = []

		for (let m = 0; m < 25_000; m++) {
			const title = `Flat Movie ${String(m)} (${String(1950 + (m % 70))})`

			files.push({ uuid: `v${String(m)}`, parent: "root", name: `${title} - 1080p.mkv`, size: 1_500_000_000, timestamp: m })

			if (m % 2 === 0) {
				// A second version, and a subtitle named after the title that both versions claim.
				files.push({ uuid: `w${String(m)}`, parent: "root", name: `${title} - 2160p.mkv`, size: 1_500_000_000, timestamp: m })
				files.push({ uuid: `s${String(m)}`, parent: "root", name: `${title}.en.srt`, size: 50_000, timestamp: m })
			} else {
				// Named after nothing: unmatched, and tested against every title by a quadratic match.
				files.push({
					uuid: `u${String(m)}`,
					parent: "root",
					name: `Extra subtitle ${String(m)}.en.srt`,
					size: 50_000,
					timestamp: m
				})
			}
		}

		const input: LibraryScanInput = { roots: [{ uuid: "root", name: "Movies", type: "auto" }], dirs: [], files }
		const started = performance.now()
		const steps = classifyVideoLibrarySteps(input)
		const durations: number[] = []
		let step = steps.next()

		while (step.done !== true) {
			const stepStarted = performance.now()

			step = steps.next()
			durations.push(performance.now() - stepStarted)
		}

		const longest = Math.max(...durations)

		const elapsed = performance.now() - started

		console.info(`25k flat titles in ${elapsed.toFixed(1)} ms, longest step ${longest.toFixed(1)} ms`)

		expect(step.value.movies).toHaveLength(25_000)
		expect(step.value.videos.reduce((count, video) => count + video.subtitles.length, 0)).toBe(25_000)
		expect(step.value.movies.filter(movie => movie.versions.length === 2)).toHaveLength(12_500)
		// Quadratic matching took ~10 s here; the bounds are loose so slow CI machines do not flake.
		expect(elapsed).toBeLessThan(3_000)
		expect(longest).toBeLessThan(500)
	})
})
