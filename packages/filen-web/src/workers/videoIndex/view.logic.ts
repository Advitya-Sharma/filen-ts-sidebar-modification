import {
	NO_FILE,
	type ExtraType,
	type LibraryArt,
	type LibraryEpisode,
	type LibraryExtra,
	type LibraryRootType,
	type LibrarySeries,
	type LibrarySubtitle,
	type LibraryVersion,
	type SidecarSubtitleFormat,
	type VideoLibrary
} from "@filen/shared/videos"
import type { IndexFile } from "@/workers/videoIndex/rootIndex.logic"
import { artCandidates, type ArtCandidate, type ArtSubject } from "@/features/videos/lib/artwork.logic"
import type { ArtSlotCode } from "@/features/videos/lib/artStore.logic"

// What the main thread gets of the library: plain, keyless records. Titles go over in full (diffed); a title's
// versions, seasons and episodes only when asked for. A file is named by its stable uuid; the worker rebuilds
// the playable file from it on demand.

export interface VideoFileRef {
	stableUuid: string
	uuid: string
	name: string
	size: number
}

export interface VideoTitleView {
	id: string
	kind: "movie" | "series" | "other"
	title: string
	sortName: string
	// 0 when unknown.
	year: number
	// Normalized title + year; empty without a year.
	mergeKey: string
	// Root directory uuids the title was found in.
	roots: string[]
	dateAdded: number
	poster: VideoFileRef | null
	backdrop: VideoFileRef | null
	logo: VideoFileRef | null
	// What the play button plays: a movie's primary version, a series' first episode, the other video itself.
	primary: VideoFileRef | null
	container: string
	// Picture height the name advertises, 0 when it says nothing.
	height: number
	// Every version is in a container that is listed but never played.
	unsupported: boolean
	disc: boolean
	versionCount: number
	extraCount: number
	seasonCount: number
	episodeCount: number
}

export interface VideoSubtitleView {
	file: VideoFileRef
	format: SidecarSubtitleFormat
	language: string | null
	forced: boolean
	sdh: boolean
	isDefault: boolean
	title: string
}

export interface VideoPartView {
	file: VideoFileRef
	container: string
	unsupported: boolean
	height: number
	subtitles: VideoSubtitleView[]
	nfo: VideoFileRef | null
	thumb: VideoFileRef | null
	edl: VideoFileRef | null
}

export interface VideoVersionView {
	label: string
	parts: VideoPartView[]
}

export interface VideoExtraView {
	type: ExtraType
	title: string
	part: VideoPartView
}

export interface VideoEpisodeView {
	season: number
	episode: number
	episodeEnd: number
	date: number
	title: string
	dateAdded: number
	versions: VideoVersionView[]
}

export interface VideoSeasonView {
	number: number
	dirUuid: string
	poster: VideoFileRef | null
	backdrop: VideoFileRef | null
	nfo: VideoFileRef | null
	episodes: VideoEpisodeView[]
}

export type VideoTitleDetail =
	| { kind: "movie"; title: VideoTitleView; versions: VideoVersionView[]; extras: VideoExtraView[]; nfo: VideoFileRef | null }
	| { kind: "series"; title: VideoTitleView; seasons: VideoSeasonView[]; extras: VideoExtraView[]; nfo: VideoFileRef | null }
	| { kind: "other"; title: VideoTitleView; part: VideoPartView | null }

export type VideoRootSize = "normal" | "large" | "huge"

export interface VideoRootStatus {
	uuid: string
	name: string
	type: LibraryRootType
	// missing: the directory was trashed, deleted or can't be resolved; its last index is kept.
	state: "loading" | "ready" | "missing" | "error"
	walking: boolean
	// Known current against the drive's event stream.
	current: boolean
	completeAt: number
	walkedEntries: number
	size: VideoRootSize
	// Video files indexed.
	videos: number
	// The directory was moved since the roots were last validated.
	moved: boolean
	// The index doesn't reflect a complete walk (the latest failed, came back partial or was stopped); what it
	// shows is the last complete one, or nothing before the first.
	incomplete: boolean
}

// One playable video of the library, as the home rows, watch filters and search need it: what the watch state is
// keyed by (the stable uuid), the title it belongs to and, for an episode, its place in the series. One per version
// (its first part); extras are left out.
export interface VideoHomeItem {
	titleId: string
	stableUuid: string
	uuid: string
	size: number
	itemClass: "movie" | "episode" | "other"
	// Episodes: 0 is specials (season) or unnumbered (episode).
	season: number
	episode: number
	episodeEnd: number
	// The episode's title from its name; empty for movies, other videos and names without one.
	episodeTitle: string
	// Versions of one movie or episode share it.
	group: string
	label: string
	dateAdded: number
	unsupported: boolean
	// The directory the file sits in.
	folder: string
}

export type VideoIndexUpdate =
	| { type: "titles"; reset: boolean; upserts: VideoTitleView[]; removed: string[] }
	| { type: "roots"; roots: VideoRootStatus[] }
	// The library's videos, keyed by stable uuid: all of them on `reset`, else what changed since the last update.
	| { type: "home"; reset: boolean; upserts: VideoHomeItem[]; removed: string[] }

// The classifier's file indices resolved against the files it was given.
export class LibraryFiles {
	public readonly library: VideoLibrary
	private readonly files: readonly IndexFile[]
	private readonly rootUuids: readonly string[]

	public constructor({
		library,
		files,
		rootUuids
	}: {
		library: VideoLibrary
		files: readonly IndexFile[]
		rootUuids: readonly string[]
	}) {
		this.library = library
		this.files = files
		this.rootUuids = rootUuids
	}

	public file(index: number): IndexFile | undefined {
		return index === NO_FILE ? undefined : this.files[index]
	}

	public ref(index: number): VideoFileRef | null {
		const file = this.file(index)

		return file === undefined
			? null
			: {
					stableUuid: file.stableUuid,
					uuid: file.uuid,
					name: file.name,
					size: file.size
				}
	}

	public roots(indices: readonly number[]): string[] {
		const out: string[] = []

		for (const index of indices) {
			const uuid = this.rootUuids[index]

			if (uuid !== undefined) {
				out.push(uuid)
			}
		}

		return out
	}

	public part(video: number): VideoPartView | null {
		const entry = this.library.videos[video]
		const file = entry === undefined ? null : this.ref(entry.file)

		if (entry === undefined || file === null) {
			return null
		}

		return {
			file,
			container: entry.container,
			unsupported: entry.unsupported,
			height: entry.height,
			subtitles: entry.subtitles.flatMap(subtitle => this.subtitle(subtitle)),
			nfo: this.ref(entry.nfo),
			thumb: this.ref(entry.thumb),
			edl: this.ref(entry.edl)
		}
	}

	private subtitle(subtitle: LibrarySubtitle): VideoSubtitleView[] {
		const file = this.ref(subtitle.file)

		return file === null
			? []
			: [
					{
						file,
						format: subtitle.format,
						language: subtitle.language,
						forced: subtitle.forced,
						sdh: subtitle.sdh,
						isDefault: subtitle.isDefault,
						title: subtitle.title
					}
				]
	}

	public versions(versions: readonly LibraryVersion[]): VideoVersionView[] {
		return versions.map(version => ({
			label: version.label,
			parts: version.parts.flatMap(video => this.part(video) ?? [])
		}))
	}

	public extras(extras: readonly LibraryExtra[]): VideoExtraView[] {
		return extras.flatMap(extra => {
			const part = this.part(extra.video)

			return part === null
				? []
				: [
						{
							type: extra.type,
							title: extra.title,
							part
						}
					]
		})
	}

	public firstPart(versions: readonly LibraryVersion[]): { video: number; file: number } | null {
		const video = versions[0]?.parts[0]
		const entry = video === undefined ? undefined : this.library.videos[video]

		return video === undefined || entry === undefined ? null : { video, file: entry.file }
	}

	public allUnsupported(versions: readonly LibraryVersion[]): boolean {
		return (
			versions.length > 0 && versions.every(version => version.parts.every(video => this.library.videos[video]?.unsupported === true))
		)
	}
}

// The first episode of the first regular season (specials only when there is nothing else).
function firstEpisode(series: LibrarySeries): LibraryEpisode | null {
	const season =
		series.seasons.find(item => item.number > 0 && item.episodes.length > 0) ?? series.seasons.find(item => item.episodes.length > 0)

	return season?.episodes[0] ?? null
}

function artRefs(files: LibraryFiles, art: LibraryArt): Pick<VideoTitleView, "poster" | "backdrop" | "logo"> {
	return {
		poster: files.ref(art.poster),
		backdrop: files.ref(art.backdrop),
		logo: files.ref(art.logo)
	}
}

function playable(files: LibraryFiles, versions: readonly LibraryVersion[]): Pick<VideoTitleView, "primary" | "container" | "height"> {
	const first = files.firstPart(versions)
	const entry = first === null ? undefined : files.library.videos[first.video]

	return {
		primary: first === null ? null : files.ref(first.file),
		container: entry?.container ?? "",
		height: entry?.height ?? 0
	}
}

// Every title of the library, in classifier order.
export function buildTitles(files: LibraryFiles): VideoTitleView[] {
	const library = files.library
	const out: VideoTitleView[] = []

	for (const movie of library.movies) {
		out.push({
			id: movie.id,
			kind: "movie",
			title: movie.title,
			sortName: movie.sortName,
			year: movie.year,
			mergeKey: movie.mergeKey,
			roots: files.roots(movie.roots),
			dateAdded: movie.dateAdded,
			...artRefs(files, movie.art),
			...playable(files, movie.versions),
			unsupported: files.allUnsupported(movie.versions),
			disc: movie.disc,
			versionCount: movie.versions.length,
			extraCount: movie.extras.length,
			seasonCount: 0,
			episodeCount: 0
		})
	}

	for (const series of library.series) {
		const episode = firstEpisode(series)
		let episodeCount = 0

		for (const season of series.seasons) {
			episodeCount += season.episodes.length
		}

		out.push({
			id: series.id,
			kind: "series",
			title: series.title,
			sortName: series.sortName,
			year: series.year,
			mergeKey: series.mergeKey,
			roots: files.roots(series.roots),
			dateAdded: series.dateAdded,
			...artRefs(files, series.art),
			...playable(files, episode?.versions ?? []),
			unsupported: false,
			disc: false,
			versionCount: 1,
			extraCount: series.extras.length,
			seasonCount: series.seasons.length,
			episodeCount
		})
	}

	for (const other of library.others) {
		const entry = library.videos[other.video]

		out.push({
			id: other.id,
			kind: "other",
			title: other.title,
			sortName: other.sortName,
			year: 0,
			mergeKey: "",
			roots: files.roots([other.root]),
			dateAdded: other.dateAdded,
			poster: null,
			backdrop: null,
			logo: null,
			primary: entry === undefined ? null : files.ref(entry.file),
			container: entry?.container ?? "",
			height: entry?.height ?? 0,
			unsupported: entry?.unsupported ?? false,
			disc: false,
			versionCount: 1,
			extraCount: 0,
			seasonCount: 0,
			episodeCount: 0
		})
	}

	return out
}

export function buildTitleDetail(files: LibraryFiles, title: VideoTitleView): VideoTitleDetail | null {
	const library = files.library

	switch (title.kind) {
		case "movie": {
			const movie = library.movies.find(item => item.id === title.id)

			return movie === undefined
				? null
				: {
						kind: "movie",
						title,
						versions: files.versions(movie.versions),
						extras: files.extras(movie.extras),
						nfo: files.ref(movie.nfo)
					}
		}

		case "series": {
			const series = library.series.find(item => item.id === title.id)

			return series === undefined
				? null
				: {
						kind: "series",
						title,
						seasons: series.seasons.map(season => ({
							number: season.number,
							dirUuid: season.dirUuid,
							poster: files.ref(season.art.poster),
							backdrop: files.ref(season.art.backdrop),
							nfo: files.ref(season.nfo),
							episodes: season.episodes.map(episode => ({
								season: episode.season,
								episode: episode.episode,
								episodeEnd: episode.episodeEnd,
								date: episode.date,
								title: episode.title,
								dateAdded: episode.dateAdded,
								versions: files.versions(episode.versions)
							}))
						})),
						extras: files.extras(series.extras),
						nfo: files.ref(series.nfo)
					}
		}

		case "other": {
			const other = library.others.find(item => item.id === title.id)

			return other === undefined
				? null
				: {
						kind: "other",
						title,
						part: files.part(other.video)
					}
		}
	}
}

function pushHomeItems({
	out,
	files,
	versions,
	base,
	folder
}: {
	out: VideoHomeItem[]
	files: LibraryFiles
	versions: readonly LibraryVersion[]
	base: Omit<VideoHomeItem, "stableUuid" | "uuid" | "size" | "label" | "unsupported" | "folder">
	folder: (parent: string) => string
}): void {
	for (const version of versions) {
		const video = version.parts[0] === undefined ? undefined : files.library.videos[version.parts[0]]
		const file = video === undefined ? undefined : files.file(video.file)

		if (video === undefined || file === undefined) {
			continue
		}

		out.push({
			...base,
			stableUuid: file.stableUuid,
			uuid: file.uuid,
			size: file.size,
			label: version.label,
			unsupported: video.unsupported,
			folder: folder(file.parent)
		})
	}
}

// Every playable video of the library, in classifier order; `folder` names a directory by uuid.
export function buildHomeItems(files: LibraryFiles, folder: (parent: string) => string): VideoHomeItem[] {
	const library = files.library
	const out: VideoHomeItem[] = []

	for (const movie of library.movies) {
		pushHomeItems({
			out,
			files,
			versions: movie.versions,
			base: {
				titleId: movie.id,
				itemClass: "movie",
				season: 0,
				episode: 0,
				episodeEnd: 0,
				episodeTitle: "",
				group: movie.id,
				dateAdded: movie.dateAdded
			},
			folder
		})
	}

	for (const series of library.series) {
		series.seasons.forEach((season, seasonIndex) => {
			season.episodes.forEach((episode, episodeIndex) => {
				pushHomeItems({
					out,
					files,
					versions: episode.versions,
					base: {
						titleId: series.id,
						itemClass: "episode",
						season: episode.season,
						episode: episode.episode,
						episodeEnd: episode.episodeEnd,
						episodeTitle: episode.title,
						group: `${series.id}/${String(seasonIndex)}/${String(episodeIndex)}`,
						dateAdded: episode.dateAdded
					},
					folder
				})
			})
		})
	}

	for (const other of library.others) {
		pushHomeItems({
			out,
			files,
			versions: [{ label: "", parts: [other.video] }],
			base: {
				titleId: other.id,
				itemClass: "other",
				season: 0,
				episode: 0,
				episodeEnd: 0,
				episodeTitle: "",
				group: other.id,
				dateAdded: other.dateAdded
			},
			folder
		})
	}

	return out
}

// A slot's art candidates for a title, with the files they name as keyless refs: candidate `file` indexes `files`.
export interface TitleArtCandidates {
	candidates: ArtCandidate[]
	files: VideoFileRef[]
}

function artSubject(files: LibraryFiles, title: VideoTitleView): ArtSubject | null {
	const library = files.library

	switch (title.kind) {
		case "movie": {
			const movie = library.movies.find(item => item.id === title.id)

			return movie === undefined ? null : { kind: "movie", movie }
		}

		case "series": {
			const series = library.series.find(item => item.id === title.id)

			return series === undefined ? null : { kind: "series", series }
		}

		case "other": {
			const other = library.others.find(item => item.id === title.id)

			// Another video's art is a still of itself, as an episode's is.
			return other === undefined
				? null
				: {
						kind: "episode",
						episode: {
							season: 0,
							episode: 0,
							episodeEnd: 0,
							date: 0,
							title: other.title,
							versions: [
								{
									label: "",
									parts: [other.video]
								}
							],
							dateAdded: other.dateAdded
						}
					}
		}
	}
}

export function titleArtCandidates(files: LibraryFiles, title: VideoTitleView, slot: ArtSlotCode): TitleArtCandidates | null {
	const subject = artSubject(files, title)

	return subject === null ? null : subjectArtCandidates(files, subject, slot)
}

function subjectArtCandidates(files: LibraryFiles, subject: ArtSubject, slot: ArtSlotCode): TitleArtCandidates {
	const refs: VideoFileRef[] = []
	const local = new Map<number, number>()
	const candidates: ArtCandidate[] = []

	for (const candidate of artCandidates(files.library, subject, slot)) {
		if (candidate.source === "typographic" || candidate.source === "gradient") {
			candidates.push(candidate)

			continue
		}

		let index = local.get(candidate.file)

		if (index === undefined) {
			const ref = files.ref(candidate.file)

			if (ref === null) {
				continue
			}

			index = refs.length
			refs.push(ref)
			local.set(candidate.file, index)
		}

		candidates.push({
			source: candidate.source,
			file: index
		})
	}

	return {
		candidates,
		files: refs
	}
}

function hasVideoFile(files: LibraryFiles, versions: readonly LibraryVersion[], stableUuid: string): boolean {
	return versions.some(version =>
		version.parts.some(video => {
			const entry = files.library.videos[video]

			return entry !== undefined && files.file(entry.file)?.stableUuid === stableUuid
		})
	)
}

// The art of one video of a title: an episode's own still for a series, the title's art otherwise.
export function itemArtCandidates(
	files: LibraryFiles,
	title: VideoTitleView,
	stableUuid: string,
	slot: ArtSlotCode
): TitleArtCandidates | null {
	if (title.kind !== "series") {
		return titleArtCandidates(files, title, slot)
	}

	const series = files.library.series.find(item => item.id === title.id)

	for (const season of series?.seasons ?? []) {
		const episode = season.episodes.find(item => hasVideoFile(files, item.versions, stableUuid))

		if (episode !== undefined) {
			return subjectArtCandidates(files, { kind: "episode", episode }, slot)
		}
	}

	return titleArtCandidates(files, title, slot)
}

// Records that changed since `previous` (id -> serialized record), and ids that are gone. Updates `previous`.
export function diffRecords<T>(
	previous: Map<string, string>,
	next: readonly T[],
	idOf: (record: T) => string
): { upserts: T[]; removed: string[] } {
	const upserts: T[] = []
	const seen = new Set<string>()

	for (const record of next) {
		const id = idOf(record)
		const signature = JSON.stringify(record)

		seen.add(id)

		if (previous.get(id) !== signature) {
			previous.set(id, signature)
			upserts.push(record)
		}
	}

	const removed: string[] = []

	for (const id of previous.keys()) {
		if (!seen.has(id)) {
			removed.push(id)
		}
	}

	for (const id of removed) {
		previous.delete(id)
	}

	return {
		upserts,
		removed
	}
}

// Titles that changed since `previous` (id -> serialized title), and ids that are gone. Updates `previous`.
export function diffTitles(
	previous: Map<string, string>,
	next: readonly VideoTitleView[]
): { upserts: VideoTitleView[]; removed: string[] } {
	return diffRecords(previous, next, title => title.id)
}
