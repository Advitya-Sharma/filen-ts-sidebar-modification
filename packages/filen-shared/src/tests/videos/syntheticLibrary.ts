import type { LibraryScanDir, LibraryScanFile, LibraryScanInput } from "@filen/shared/videos"

// A large, realistic library: movie directories with sidecars and versions, shows with seasons and subtitles,
// phone clips and unrelated files.
export function syntheticLibrary(totalFiles: number): LibraryScanInput {
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

	for (let s = 0; files.length < totalFiles * 0.8; s++) {
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

	for (let day = 0; files.length < totalFiles; day++) {
		const folder = dir(phone, `2023-${String((day % 12) + 1).padStart(2, "0")}-${String((day % 28) + 1).padStart(2, "0")} Trip ${day}`)

		for (let clip = 0; clip < 20 && files.length < totalFiles; clip++) {
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
