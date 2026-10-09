// @vitest-environment jsdom

import { createElement, type ReactNode } from "react"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, render } from "@testing-library/react"
import { QueryClient, QueryClientProvider } from "@tanstack/react-query"
import type { AnyFile } from "@filen/sdk-rs"
import "@/lib/i18n"
import type { VideoFileRef, VideoPartView, VideoTitleDetail, VideoTitleView } from "@/workers/videoIndex/view.logic"

// What opening a movie's page costs: the title's full probe (one), its NFO (one read), and art only for what is on
// screen after the dwell. Nothing is listed, walked or checked against the drive; the extras' stills wait for the
// shelf to come near.

const { sdkApi, library, observers } = vi.hoisted(() => ({
	sdkApi: {
		readVideoProbe: vi.fn(),
		makeArtFromRange: vi.fn(),
		storeArt: vi.fn(),
		cancelPreviewDownload: vi.fn(),
		cancelPreviewDownloadIfEarly: vi.fn(),
		cancelPreviewDownloadUnlessFinishing: vi.fn(),
		openRangePort: vi.fn(),
		downloadFileBytes: vi.fn(),
		getLastDriveEventId: vi.fn(),
		listVideoRoot: vi.fn(),
		videoIndexCheck: vi.fn(),
		videoIndexConfigure: vi.fn(),
		listDirectory: vi.fn(),
		videoTitleDetail: vi.fn(),
		videoTitleArt: vi.fn(),
		videoItemArt: vi.fn(),
		playableVideoFiles: vi.fn()
	},
	library: {
		snapshot: { titles: [] as VideoTitleView[], hidden: [], byId: new Map<string, VideoTitleView>(), roots: [], loaded: true },
		items: [],
		retainHome: vi.fn(() => () => undefined)
	},
	observers: [] as { callback: IntersectionObserverCallback; options: IntersectionObserverInit | undefined; targets: Set<Element> }[]
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi }))
vi.mock("@/features/videos/workers/frame.worker.ts?worker", () => ({ default: vi.fn() }))
vi.mock("@/features/drive/lib/thumbCache", () => ({ readThumbnailBlob: () => Promise.resolve(null) }))
vi.mock("@/lib/storage/adapter", () => ({
	kvGetJson: () => Promise.resolve(null),
	kvSetJson: () => Promise.resolve(),
	kvKeys: () => Promise.resolve([]),
	kvEntriesJson: () => Promise.resolve([]),
	kvDelete: () => Promise.resolve()
}))
vi.mock("@tanstack/react-router", () => {
	const navigate = vi.fn()
	const router = { history: { canGoBack: () => false, back: vi.fn() } }

	return {
		useNavigate: () => navigate,
		useRouter: () => router,
		useRouterState: ({ select }: { select: (state: unknown) => unknown }) =>
			select({ location: { state: { __TSR_key: "entry" }, href: "/videos/title/movie" } }),
		Link: ({
			children,
			to: _to,
			params: _params,
			search: _search,
			...props
		}: {
			children: ReactNode
			to: string
			params?: unknown
			search?: unknown
		}) => createElement("a", { ...props, href: "#" }, children)
	}
})
vi.mock("@/features/videos/lib/watchState", () => {
	const snapshot = { doc: { v: 1, items: {}, series: {} }, loaded: true, unsynced: false, dirty: false, lastPlayedCeilingMs: undefined }

	return {
		watchState: {
			subscribe: () => () => undefined,
			getSnapshot: () => snapshot,
			markWatched: vi.fn(),
			markUnwatched: vi.fn(),
			hide: vi.fn()
		}
	}
})
// The library client as the index's worker feeds it: details, art lookups and file rebuilds are worker calls.
vi.mock("@/features/videos/lib/index", () => {
	const artOf = async (art: Promise<{ candidates: unknown[]; files: VideoFileRef[] } | null>) => {
		const resolved = await art

		return resolved === null
			? null
			: {
					candidates: resolved.candidates,
					fileUuid: (file: number) => resolved.files[file]?.uuid ?? null,
					resolveFile: (file: number) => {
						const ref = resolved.files[file]

						return Promise.resolve(
							ref === undefined ? null : { uuid: ref.uuid, size: ref.size, file: { uuid: ref.uuid } as unknown as AnyFile }
						)
					}
				}
	}

	return {
		VIDEO_TITLE_DETAIL_KEY_PREFIX: ["videos", "title"],
		videoLibrary: {
			subscribe: () => () => undefined,
			getSnapshot: () => library.snapshot,
			getHomeItems: () => library.items,
			retainHome: library.retainHome,
			open: () => Promise.resolve(),
			titleArt: (id: string, slot: string) => artOf(sdkApi.videoTitleArt(id, slot) as Promise<never>),
			itemArt: (id: string, stableUuid: string, slot: string) => artOf(sdkApi.videoItemArt(id, stableUuid, slot) as Promise<never>),
			playableFiles: (stableUuids: string[]) => sdkApi.playableVideoFiles(stableUuids) as Promise<unknown[]>
		}
	}
})

import { VideoTitleScreen } from "@/features/videos/screens/videoTitle"
import { CARD_PROBE_DWELL_MS } from "@/features/videos/lib/probeService.logic"
import { title } from "@/tests/support/videoHome"

class FakeIntersectionObserver {
	private readonly entry: { callback: IntersectionObserverCallback; options: IntersectionObserverInit | undefined; targets: Set<Element> }

	public constructor(callback: IntersectionObserverCallback, options?: IntersectionObserverInit) {
		this.entry = { callback, options, targets: new Set() }
		observers.push(this.entry)
	}

	public observe(target: Element): void {
		this.entry.targets.add(target)
	}

	public unobserve(target: Element): void {
		this.entry.targets.delete(target)
	}

	public disconnect(): void {
		this.entry.targets.clear()
	}
}

function intersect(which: (options: IntersectionObserverInit | undefined) => boolean): void {
	for (const observer of observers.filter(entry => which(entry.options))) {
		observer.callback(
			[...observer.targets].map(
				target => ({ target, isIntersecting: true, intersectionRatio: 1 }) as unknown as IntersectionObserverEntry
			),
			observer as unknown as IntersectionObserver
		)
	}
}

const isVisibleObserver = (options: IntersectionObserverInit | undefined) => options?.threshold === 0.5
const isShelfObserver = (options: IntersectionObserverInit | undefined) => options?.rootMargin === "50% 50%"

function ref(name: string): VideoFileRef {
	return { stableUuid: `${name}-s`, uuid: `${name}-u`, name, size: 1_000 }
}

function part(name: string, patch: Partial<VideoPartView> = {}): VideoPartView {
	return {
		file: ref(name),
		container: "mkv",
		unsupported: false,
		height: 1080,
		subtitles: [],
		nfo: null,
		thumb: null,
		edl: null,
		...patch
	}
}

const MOVIE = title("movie", {
	title: "Heat",
	year: 1995,
	backdrop: ref("fanart.jpg"),
	poster: ref("poster.jpg"),
	versionCount: 1,
	extraCount: 2
})
const DETAIL: VideoTitleDetail = {
	kind: "movie",
	title: MOVIE,
	versions: [{ label: "", parts: [part("Heat (1995).mkv")] }],
	extras: [
		{ type: "trailer", title: "Trailer", part: part("Heat-trailer.mkv") },
		{ type: "featurette", title: "Making of", part: part("Heat-featurette.mkv") }
	],
	nfo: ref("movie.nfo")
}

const NFO =
	"<movie><title>Heat</title><mpaa>R</mpaa><runtime>101</runtime><plot>A crew of thieves.</plot><actor><name>Al Pacino</name></actor></movie>"

function networkCalls() {
	return {
		probes: sdkApi.readVideoProbe.mock.calls.length,
		nfoReads: sdkApi.downloadFileBytes.mock.calls.length,
		art: sdkApi.makeArtFromRange.mock.calls.length,
		listings:
			sdkApi.listDirectory.mock.calls.length +
			sdkApi.listVideoRoot.mock.calls.length +
			sdkApi.videoIndexCheck.mock.calls.length +
			sdkApi.getLastDriveEventId.mock.calls.length
	}
}

async function flush(): Promise<void> {
	await act(async () => {
		await vi.advanceTimersByTimeAsync(0)
	})
}

beforeEach(() => {
	vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })
	vi.stubGlobal("IntersectionObserver", FakeIntersectionObserver)
	vi.stubGlobal(
		"matchMedia",
		(query: string) =>
			({
				matches: true,
				media: query,
				addEventListener: () => undefined,
				removeEventListener: () => undefined
			}) as unknown as MediaQueryList
	)
	observers.length = 0
	library.snapshot = { titles: [MOVIE], hidden: [], byId: new Map([[MOVIE.id, MOVIE]]), roots: [], loaded: true }
	sdkApi.videoTitleDetail.mockResolvedValue(DETAIL)
	sdkApi.playableVideoFiles.mockImplementation((stableUuids: string[]) =>
		Promise.resolve(stableUuids.map(stableUuid => ({ uuid: stableUuid.replace(/-s$/, "-u"), stableUUID: stableUuid, size: 1_000n })))
	)
	sdkApi.downloadFileBytes.mockResolvedValue(new TextEncoder().encode(NFO))
	sdkApi.readVideoProbe.mockResolvedValue(null)
	sdkApi.makeArtFromRange.mockResolvedValue({ type: "notImage" })
	sdkApi.videoTitleArt.mockImplementation((_id: string, slot: string) =>
		Promise.resolve({
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
			files: [ref(slot === "b" ? "fanart.jpg" : "poster.jpg")]
		})
	)
	sdkApi.videoItemArt.mockImplementation((_id: string, stableUuid: string) =>
		Promise.resolve({ candidates: [{ source: "frame", file: 0 }, { source: "gradient" }], files: [{ ...ref(stableUuid), stableUuid }] })
	)
})

afterEach(() => {
	vi.useRealTimers()
	vi.unstubAllGlobals()
})

describe("title page request counts", () => {
	it("costs one full probe, one NFO read and the art on screen after the dwell, and lists nothing", async () => {
		let deliverNfo: () => void = () => undefined

		sdkApi.downloadFileBytes.mockImplementation(
			() =>
				new Promise(resolve => {
					deliverNfo = () => {
						resolve(new TextEncoder().encode(NFO))
					}
				})
		)

		const client = new QueryClient()
		const view = render(createElement(QueryClientProvider, { client }, createElement(VideoTitleScreen, { titleId: MOVIE.id })))

		// The probe and art gates hold a free slot briefly (tens of ms) for higher-ranked work.
		await act(async () => {
			await vi.advanceTimersByTimeAsync(100)
		})
		await flush()

		// While the NFO the index knows of is read, its synopsis and cast lines hold their room.
		expect(view.container.querySelectorAll("[data-slot=skeleton]").length).toBeGreaterThanOrEqual(3)

		deliverNfo()
		await flush()
		await flush()

		expect(view.container.querySelectorAll("[data-slot=skeleton]")).toHaveLength(0)
		// The metadata line reads as one run, with separators.
		expect(view.getByText("1995 · 1h 41m · R")).toBeTruthy()
		// Only this title's watch state is read: the library's video list is never asked for.
		expect(library.retainHome).not.toHaveBeenCalled()

		expect(view.getByRole("heading", { level: 1 }).textContent).toBe("Heat")
		expect(networkCalls()).toEqual({ probes: 1, nfoReads: 1, art: 0, listings: 0 })
		expect(sdkApi.readVideoProbe.mock.calls[0]?.[2]).toBe("full")
		// The NFO's synopsis and cast show.
		expect(view.getByText("A crew of thieves.")).toBeTruthy()
		expect(view.getByText("Cast: Al Pacino")).toBeTruthy()
		// The extras' stills are not even looked up before the shelf comes near.
		expect(sdkApi.videoItemArt).not.toHaveBeenCalled()

		// The poster and the backdrop on screen: their downloads only after the dwell.
		act(() => {
			intersect(isVisibleObserver)
		})
		await act(async () => {
			await vi.advanceTimersByTimeAsync(CARD_PROBE_DWELL_MS - 1)
		})

		expect(networkCalls().art).toBe(0)

		await act(async () => {
			await vi.advanceTimersByTimeAsync(200)
		})

		expect(networkCalls()).toEqual({ probes: 1, nfoReads: 1, art: 2, listings: 0 })

		// The shelf comes near: its stills are looked up (worker calls, local reads), still no listing.
		act(() => {
			intersect(isShelfObserver)
		})
		await flush()

		expect(sdkApi.videoItemArt).toHaveBeenCalledTimes(2)
		expect(networkCalls().listings).toBe(0)

		view.unmount()
	})

	it("moves Ends at on with the clock while the page stays open", async () => {
		vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] })
		vi.setSystemTime(new Date(2026, 9, 9, 20, 0, 30))

		const client = new QueryClient()
		const view = render(createElement(QueryClientProvider, { client }, createElement(VideoTitleScreen, { titleId: MOVIE.id })))

		await act(async () => {
			await vi.advanceTimersByTimeAsync(100)
		})
		await flush()
		await flush()

		const endsAt = () => view.getByText(/^Ends at /).textContent

		const first = endsAt()

		await act(async () => {
			await vi.advanceTimersByTimeAsync(60_000)
		})

		expect(endsAt()).not.toBe(first)
		view.unmount()
	})
})
