import { describe, expect, it } from "vitest"
import type { ActiveLibraryRoot } from "@filen/shared/videos"
import {
	busyRoots,
	detectedRootKind,
	finishedScans,
	libraryRootViews,
	scanTransition,
	showsVideosTabs,
	videoKindCounts,
	videoRootPhase,
	videosHomeView,
	videosSection
} from "@/features/videos/lib/libraryView.logic"
import type { VideoRootStatus, VideoTitleView } from "@/workers/videoIndex/view.logic"
import { videos } from "@/locales/en/videos"

function status(uuid: string, patch: Partial<VideoRootStatus> = {}): VideoRootStatus {
	return {
		uuid,
		name: uuid,
		type: "auto",
		state: "ready",
		walking: false,
		current: true,
		completeAt: 1_000,
		walkedEntries: 10,
		size: "normal",
		videos: 3,
		moved: false,
		incomplete: false,
		...patch
	}
}

function root(dirUuid: string, enabled = true): ActiveLibraryRoot {
	return {
		dirUuid,
		type: "auto",
		enabled,
		added: 0
	}
}

function title(id: string, kind: VideoTitleView["kind"], roots: string[], episodeCount = 0): VideoTitleView {
	return {
		id,
		kind,
		title: id,
		sortName: id,
		year: 0,
		mergeKey: "",
		roots,
		dateAdded: 0,
		poster: null,
		backdrop: null,
		logo: null,
		primary: null,
		container: "mkv",
		height: 0,
		unsupported: false,
		disc: false,
		versionCount: 1,
		extraCount: 0,
		seasonCount: 0,
		episodeCount
	}
}

describe("videoRootPhase", () => {
	it("reads scanning before anything else, then missing, incomplete, loading, empty and ready", () => {
		expect(videoRootPhase(undefined)).toBe("loading")
		expect(videoRootPhase(status("a", { walking: true, state: "missing" }))).toBe("scanning")
		expect(videoRootPhase(status("a", { state: "missing" }))).toBe("missing")
		expect(videoRootPhase(status("a", { state: "error", completeAt: 0, videos: 0 }))).toBe("incomplete")
		expect(videoRootPhase(status("a", { incomplete: true }))).toBe("incomplete")
		expect(videoRootPhase(status("a", { state: "loading", completeAt: 0, videos: 0 }))).toBe("loading")
		expect(videoRootPhase(status("a", { videos: 0 }))).toBe("empty")
		expect(videoRootPhase(status("a"))).toBe("ready")
	})
})

describe("libraryRootViews", () => {
	it("joins the configured directories with the index's statuses; a disabled one has none", () => {
		const views = libraryRootViews([root("a"), root("b", false), root("c")], [status("a"), status("b"), status("x")])

		expect(views.map(view => [view.uuid, view.enabled, view.status?.uuid, view.phase])).toEqual([
			["a", true, "a", "ready"],
			["b", false, undefined, "loading"],
			["c", true, undefined, "loading"]
		])
	})
})

describe("videoKindCounts", () => {
	it("counts each title once in the total and once per directory it was found in", () => {
		const { total, byRoot } = videoKindCounts([
			title("heat", "movie", ["a", "b"]),
			title("lost", "series", ["b"], 12),
			title("clip", "other", ["a"])
		])

		expect(total).toEqual({ movies: 1, shows: 1, other: 1, episodes: 12 })
		expect(byRoot.get("a")).toEqual({ movies: 1, shows: 0, other: 1, episodes: 0 })
		expect(byRoot.get("b")).toEqual({ movies: 1, shows: 1, other: 0, episodes: 12 })
	})
})

describe("detectedRootKind", () => {
	it("takes the kind most titles have, movies winning a tie, and nothing before any", () => {
		expect(detectedRootKind(undefined)).toBeNull()
		expect(detectedRootKind({ movies: 0, shows: 0, other: 0, episodes: 0 })).toBeNull()
		expect(detectedRootKind({ movies: 1, shows: 4, other: 2, episodes: 30 })).toBe("shows")
		expect(detectedRootKind({ movies: 2, shows: 2, other: 0, episodes: 3 })).toBe("movies")
		expect(detectedRootKind({ movies: 0, shows: 0, other: 5, episodes: 0 })).toBe("other")
	})
})

describe("scan line", () => {
	it("names the enabled directories scanned or waiting for their first scan", () => {
		const views = libraryRootViews(
			[root("a"), root("b"), root("c"), root("d", false)],
			[status("a", { walking: true }), status("b", { state: "loading" }), status("c")]
		)

		expect(busyRoots(views).map(view => view.uuid)).toEqual(["a", "b"])
	})

	it("reports a finished scan only when it completed, not when it was stopped or failed", () => {
		const after = libraryRootViews(
			[root("a"), root("b"), root("c"), root("d")],
			[status("a"), status("b", { incomplete: true }), status("c", { videos: 0 }), status("d", { walking: true })]
		)

		expect(finishedScans(["a", "b", "c", "d"], after).map(view => view.uuid)).toEqual(["a", "c"])
		expect(finishedScans([], after)).toEqual([])
	})
})

describe("videosHomeView", () => {
	it("is the first run without any directory, and all-off when every directory is turned off", () => {
		expect(videosHomeView({ roots: [], titleCount: 0 })).toEqual({ kind: "firstRun" })
		expect(videosHomeView({ roots: libraryRootViews([root("a", false)], []), titleCount: 0 })).toEqual({ kind: "allOff" })
	})

	it("waits for titles while a directory is scanned, even when another came back empty", () => {
		const roots = libraryRootViews([root("a"), root("b")], [status("a", { walking: true, videos: 0 }), status("b", { videos: 0 })])

		expect(videosHomeView({ roots, titleCount: 0 })).toEqual({ kind: "library", notices: [], content: { kind: "pending" } })
	})

	it("offers the empty state once every scan is over and nothing was found", () => {
		const roots = libraryRootViews([root("a"), root("b")], [status("a", { videos: 0 }), status("b", { state: "missing" })])

		expect(videosHomeView({ roots, titleCount: 0 })).toEqual({
			kind: "library",
			notices: [{ uuid: "b", kind: "missing", hasPrevious: true }],
			content: { kind: "empty", roots: ["a"] }
		})
	})

	it("shows the rows with a notice per incomplete directory, telling whether an earlier scan stands in", () => {
		const roots = libraryRootViews(
			[root("a"), root("b"), root("c")],
			[status("a"), status("b", { incomplete: true }), status("c", { state: "error", completeAt: 0, videos: 0, incomplete: true })]
		)

		expect(videosHomeView({ roots, titleCount: 5 })).toEqual({
			kind: "library",
			notices: [
				{ uuid: "b", kind: "incomplete", hasPrevious: true },
				{ uuid: "c", kind: "incomplete", hasPrevious: false }
			],
			content: { kind: "rows" }
		})
	})
})

describe("videosSection", () => {
	it("maps the /videos routes under the shared header", () => {
		expect(videosSection("/videos")).toEqual({ kind: "home" })
		expect(videosSection("/videos/")).toEqual({ kind: "home" })
		expect(videosSection("/videos/movies")).toEqual({ kind: "movies" })
		expect(videosSection("/videos/search")).toEqual({ kind: "search" })
		expect(videosSection("/videos/library/abc")).toEqual({ kind: "library", uuid: "abc" })
	})

	it("leaves routes with their own header, and other modules, alone", () => {
		expect(videosSection("/videos/title/abc")).toBeNull()
		expect(videosSection("/videos/watch/abc")).toBeNull()
		expect(videosSection("/videosx")).toBeNull()
		expect(videosSection("/photos")).toBeNull()
	})
})

describe("showsVideosTabs", () => {
	it("shows the phone tabs once there is a library, never over the player", () => {
		expect(showsVideosTabs({ pathname: "/videos", hasLibrary: false })).toBe(false)
		expect(showsVideosTabs({ pathname: "/videos", hasLibrary: true })).toBe(true)
		expect(showsVideosTabs({ pathname: "/videos/shows", hasLibrary: true })).toBe(true)
		expect(showsVideosTabs({ pathname: "/videos/watch/abc", hasLibrary: true })).toBe(false)
	})
})

describe("stopped scans", () => {
	it("leave the scan line at once, read as incomplete and are never announced as scanned", () => {
		const roots = libraryRootViews(
			[root("first"), root("again")],
			[status("first", { state: "error", completeAt: 0, videos: 0, incomplete: true }), status("again", { incomplete: true })]
		)

		expect(roots.map(view => view.phase)).toEqual(["incomplete", "incomplete"])
		expect(busyRoots(roots)).toEqual([])
		expect(scanTransition(["first", "again"], roots)).toEqual({ done: [], started: undefined })
	})
})

describe("scanTransition", () => {
	it("names completed scans and the first newly started one", () => {
		const roots = libraryRootViews(
			[root("a"), root("b"), root("c")],
			[status("a"), status("b", { walking: true }), status("c", { state: "loading" })]
		)
		const { done, started } = scanTransition(["a", "c"], roots)

		expect(done.map(view => view.uuid)).toEqual(["a"])
		expect(started?.uuid).toBe("b")
	})
})

describe("videos copy", () => {
	it("shows no running count while scanning: a scan's result arrives at once", () => {
		expect(videos.videosScanning).not.toContain("count")
	})

	it("formats every count and file total as a number", () => {
		for (const text of Object.values(videos)) {
			for (const match of text.matchAll(/{{(count|files)\b[^}]*}}/g)) {
				expect(match[0]).toMatch(/, number}}$/)
			}
		}
	})

	it("names the directory the chooser adds", () => {
		expect(videos.videosChooserConfirm).toContain("{{name}}")
	})
})
