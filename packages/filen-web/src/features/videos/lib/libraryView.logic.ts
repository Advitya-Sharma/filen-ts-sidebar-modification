import type { ActiveLibraryRoot, LibraryRootType } from "@filen/shared/videos"
import { pathIsUnder } from "@/features/shell/lib/appShell.logic"
import type { VideoRootStatus, VideoTitleView } from "@/workers/videoIndex/view.logic"

// What the Videos shell shows of the library: per-directory phases, counts, the scan line and the home's
// state, all derived from the configuration and the index snapshot already held (no requests).

export type VideoRootPhase = "loading" | "scanning" | "ready" | "empty" | "incomplete" | "missing"

export function videoRootPhase(status: VideoRootStatus | undefined): VideoRootPhase {
	if (status === undefined) {
		return "loading"
	}

	if (status.walking) {
		return "scanning"
	}

	if (status.state === "missing") {
		return "missing"
	}

	if (status.state === "error" || status.incomplete) {
		return "incomplete"
	}

	if (status.state === "loading") {
		return "loading"
	}

	return status.videos === 0 ? "empty" : "ready"
}

export interface VideoLibraryRootView {
	uuid: string
	type: LibraryRootType
	enabled: boolean
	// Absent for a disabled directory (the index doesn't load it) and until the index reports it.
	status: VideoRootStatus | undefined
	phase: VideoRootPhase
}

// The configured directories in the order they were added, each with what the index reports of it.
export function libraryRootViews(roots: readonly ActiveLibraryRoot[], statuses: readonly VideoRootStatus[]): VideoLibraryRootView[] {
	const byUuid = new Map(statuses.map(status => [status.uuid, status]))

	return roots.map(root => {
		const status = root.enabled ? byUuid.get(root.dirUuid) : undefined

		return {
			uuid: root.dirUuid,
			type: root.type,
			enabled: root.enabled,
			status,
			phase: videoRootPhase(status)
		}
	})
}

// A directory whose phase shows work under way: scanned now, or waiting for its first scan.
export function isRootBusy(phase: VideoRootPhase): boolean {
	return phase === "scanning" || phase === "loading"
}

export interface VideoKindCounts {
	movies: number
	shows: number
	other: number
	episodes: number
}

function emptyCounts(): VideoKindCounts {
	return {
		movies: 0,
		shows: 0,
		other: 0,
		episodes: 0
	}
}

function addTitle(counts: VideoKindCounts, title: VideoTitleView): void {
	if (title.kind === "movie") {
		counts.movies++
	} else if (title.kind === "series") {
		counts.shows++
		counts.episodes += title.episodeCount
	} else {
		counts.other++
	}
}

// Titles per kind across the library and per directory, in one pass. A title found in two directories counts
// once in the total and once in each.
export function videoKindCounts(titles: readonly VideoTitleView[]): { total: VideoKindCounts; byRoot: Map<string, VideoKindCounts> } {
	const total = emptyCounts()
	const byRoot = new Map<string, VideoKindCounts>()

	for (const title of titles) {
		addTitle(total, title)

		for (const root of title.roots) {
			let counts = byRoot.get(root)

			if (counts === undefined) {
				counts = emptyCounts()
				byRoot.set(root, counts)
			}

			addTitle(counts, title)
		}
	}

	return {
		total,
		byRoot
	}
}

// What an auto directory's names were classified as, by the kind most of its titles have; null before any.
export function detectedRootKind(counts: VideoKindCounts | undefined): Exclude<LibraryRootType, "auto"> | null {
	if (counts === undefined) {
		return null
	}

	const top = Math.max(counts.movies, counts.shows, counts.other)

	if (top === 0) {
		return null
	}

	return counts.movies === top ? "movies" : counts.shows === top ? "shows" : "other"
}

// The directories the scan line names: those being scanned or waiting for their first scan.
export function busyRoots(roots: readonly VideoLibraryRootView[]): VideoLibraryRootView[] {
	return roots.filter(root => root.enabled && isRootBusy(root.phase))
}

// Directories that were scanning before and have now finished with a complete scan (a stopped or failed one
// isn't "scanned").
export function finishedScans(previouslyScanning: readonly string[], roots: readonly VideoLibraryRootView[]): VideoLibraryRootView[] {
	return roots.filter(root => previouslyScanning.includes(root.uuid) && (root.phase === "ready" || root.phase === "empty"))
}

// What changed in the scan line since it named `previouslyBusy`: scans that completed, and the first newly started.
export function scanTransition(
	previouslyBusy: readonly string[],
	roots: readonly VideoLibraryRootView[]
): { done: VideoLibraryRootView[]; started: VideoLibraryRootView | undefined } {
	return {
		done: finishedScans(previouslyBusy, roots),
		started: busyRoots(roots).find(root => !previouslyBusy.includes(root.uuid))
	}
}

export interface VideoRootNotice {
	uuid: string
	kind: "incomplete" | "missing"
	// An earlier complete scan is what the library shows of it.
	hasPrevious: boolean
}

export type VideosHomeContent =
	// Titles to show (the rows).
	| { kind: "rows" }
	// Nothing to show yet: a scan is under way, or the index hasn't reported.
	| { kind: "pending" }
	// Every scanned directory came back without a video file.
	| { kind: "empty"; roots: string[] }

export type VideosHomeView =
	| { kind: "firstRun" }
	// Directories exist but every one is turned off.
	| { kind: "allOff" }
	| { kind: "library"; notices: VideoRootNotice[]; content: VideosHomeContent }

export function videosHomeView({ roots, titleCount }: { roots: readonly VideoLibraryRootView[]; titleCount: number }): VideosHomeView {
	if (roots.length === 0) {
		return {
			kind: "firstRun"
		}
	}

	const enabled = roots.filter(root => root.enabled)

	if (enabled.length === 0) {
		return {
			kind: "allOff"
		}
	}

	const notices: VideoRootNotice[] = []

	for (const root of enabled) {
		if (root.phase === "incomplete" || root.phase === "missing") {
			notices.push({
				uuid: root.uuid,
				kind: root.phase,
				hasPrevious: (root.status?.completeAt ?? 0) > 0
			})
		}
	}

	const empty = enabled.filter(root => root.phase === "empty").map(root => root.uuid)
	let content: VideosHomeContent

	if (titleCount > 0) {
		content = {
			kind: "rows"
		}
	} else if (empty.length > 0 && !enabled.some(root => isRootBusy(root.phase))) {
		content = {
			kind: "empty",
			roots: empty
		}
	} else {
		content = {
			kind: "pending"
		}
	}

	return {
		kind: "library",
		notices,
		content
	}
}

export type VideosSection = { kind: "home" | "movies" | "shows" | "other" | "search" } | { kind: "library"; uuid: string }

const SECTION_PATHS = ["movies", "shows", "other", "search"] as const

// The section a /videos pathname shows under the shared header; null for routes that bring their own (a
// title page, the player).
export function videosSection(pathname: string): VideosSection | null {
	if (!pathIsUnder(pathname, "/videos")) {
		return null
	}

	const rest = pathname.replace(/^\/videos\/?/, "").replace(/\/$/, "")

	if (rest === "") {
		return {
			kind: "home"
		}
	}

	const section = SECTION_PATHS.find(path => path === rest)

	if (section !== undefined) {
		return {
			kind: section
		}
	}

	const library = /^library\/([^/]+)$/.exec(rest)?.[1]

	return library === undefined
		? null
		: {
				kind: "library",
				uuid: library
			}
}

// The phone tabs: only once the library has a directory, and never over the player.
export function showsVideosTabs({ pathname, hasLibrary }: { pathname: string; hasLibrary: boolean }): boolean {
	return hasLibrary && !pathIsUnder(pathname, "/videos/watch")
}
