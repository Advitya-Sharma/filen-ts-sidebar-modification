import {
	classifyVideoLibrary,
	type LibraryRootType,
	type LibraryScanDir,
	type LibraryScanFile,
	type LibraryScanInput,
	type LibraryScanRoot,
	type LibraryVersion,
	type VideoLibrary
} from "@filen/shared/videos"

export interface FixtureRoot {
	name: string
	type?: LibraryRootType
	// Paths relative to the root, "/"-separated.
	paths: readonly string[]
	// Byte sizes by path; 1 GiB otherwise.
	sizes?: Record<string, number>
}

const GIB = 1024 * 1024 * 1024

export interface FixtureInput extends LibraryScanInput {
	roots: LibraryScanRoot[]
	dirs: LibraryScanDir[]
	files: LibraryScanFile[]
}

// Builds a scan input from path lists. Directory and file uuids are their full paths, so results read naturally.
export function fixtureInput(roots: readonly FixtureRoot[]): FixtureInput {
	const dirs: LibraryScanDir[] = []
	const files: LibraryScanFile[] = []
	const seenDirs = new Set<string>()

	for (const root of roots) {
		for (const path of root.paths) {
			const segments = path.split("/")
			let parent = root.name

			for (let i = 0; i < segments.length - 1; i++) {
				const uuid = `${parent}/${segments[i] ?? ""}`

				if (!seenDirs.has(uuid)) {
					seenDirs.add(uuid)
					dirs.push({
						uuid,
						parent,
						name: segments[i] ?? ""
					})
				}

				parent = uuid
			}

			files.push({
				uuid: `${root.name}/${path}`,
				parent,
				name: segments[segments.length - 1] ?? "",
				size: root.sizes?.[path] ?? GIB,
				timestamp: files.length + 1
			})
		}
	}

	return {
		roots: roots.map(root => ({
			uuid: root.name,
			name: root.name,
			type: root.type ?? "auto"
		})),
		dirs,
		files
	}
}

function pathOf(input: LibraryScanInput, file: number): string {
	return input.files[file]?.uuid ?? "?"
}

function videoPath(library: VideoLibrary, input: LibraryScanInput, video: number): string {
	return pathOf(input, library.videos[video]?.file ?? -1)
}

function versionText(library: VideoLibrary, input: LibraryScanInput, version: LibraryVersion): string {
	const parts = version.parts.map(part => videoPath(library, input, part)).join(" + ")

	return version.label.length > 0 ? `${version.label}: ${parts}` : parts
}

export interface MovieSummary {
	title: string
	year: number
	versions: string[]
	extras: string[]
}

export interface LibrarySummary {
	movies: MovieSummary[]
	series: { title: string; year: number; seasons: Record<number, string[]>; extras: string[] }[]
	others: string[]
}

// Readable projection of a classification for assertions.
export function summarize(library: VideoLibrary, input: LibraryScanInput): LibrarySummary {
	return {
		movies: library.movies.map(movie => ({
			title: movie.title,
			year: movie.year,
			versions: movie.versions.map(version => versionText(library, input, version)),
			extras: movie.extras.map(extra => `${extra.type}: ${videoPath(library, input, extra.video)}`)
		})),
		series: library.series.map(series => {
			const seasons: Record<number, string[]> = {}

			for (const season of series.seasons) {
				seasons[season.number] = season.episodes.map(episode => {
					const number =
						episode.date > 0
							? `D${episode.date}`
							: episode.episodeEnd > episode.episode
								? `E${episode.episode}-${episode.episodeEnd}`
								: `E${episode.episode}`
					const versions = episode.versions.map(version => versionText(library, input, version)).join(" | ")

					return `${number} ${episode.title} = ${versions}`
				})
			}

			return {
				title: series.title,
				year: series.year,
				seasons,
				extras: series.extras.map(extra => `${extra.type}: ${videoPath(library, input, extra.video)}`)
			}
		}),
		others: library.others.map(other => videoPath(library, input, other.video))
	}
}

export function classifyPaths(roots: readonly FixtureRoot[]): { library: VideoLibrary; input: LibraryScanInput; summary: LibrarySummary } {
	const input = fixtureInput(roots)
	const library = classifyVideoLibrary(input)

	return {
		library,
		input,
		summary: summarize(library, input)
	}
}
