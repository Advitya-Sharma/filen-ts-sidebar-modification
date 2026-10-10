// @vitest-environment jsdom

import { createElement, type ReactNode } from "react"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, render } from "@testing-library/react"
import { QueryClient, QueryClientProvider } from "@tanstack/react-query"
import type { AnyFile } from "@filen/sdk-rs"
import "@/lib/i18n"
import type { HomeQuery, VideoCatalog } from "@/workers/videoIndex/catalog.logic"

// What opening Videos Home costs: rendering it from the library's catalog (one worker call) makes no listing or probe
// at all, and a card's artwork may only cost a download once that card has been on screen for the dwell.

const { sdkApi, catalog, observers } = vi.hoisted(() => {
	return {
		sdkApi: {
			readVideoProbe: vi.fn(),
			makeArtFromRange: vi.fn(),
			storeArt: vi.fn(),
			cancelPreviewDownload: vi.fn(),
			cancelPreviewDownloadIfEarly: vi.fn(),
			openRangePort: vi.fn(),
			getLastDriveEventId: vi.fn(),
			listVideoRoot: vi.fn(),
			videoIndexCheck: vi.fn(),
			videoIndexConfigure: vi.fn(),
			listDirectory: vi.fn(),
			videoTitleArt: vi.fn(),
			videoItemArt: vi.fn(),
			videoCatalogHome: vi.fn(),
			playableVideoFiles: vi.fn()
		},
		catalog: { current: null as VideoCatalog | null },
		observers: [] as { callback: IntersectionObserverCallback; options: IntersectionObserverInit | undefined; targets: Set<Element> }[]
	}
})

vi.mock("@/lib/sdk/client", () => ({ sdkApi }))
vi.mock("@/features/videos/workers/frame.worker.ts?worker", () => ({ default: vi.fn() }))
vi.mock("@/features/drive/lib/thumbCache", () => ({ readThumbnailBlob: () => Promise.resolve(null) }))
vi.mock("@/lib/storage/adapter", () => ({
	kvGetJson: () => Promise.resolve(null),
	kvSetJson: () => Promise.resolve(),
	kvKeys: () => Promise.resolve([]),
	kvEntriesJson: () => Promise.resolve([]),
	kvEntriesJsonInSlices: () => Promise.resolve([]),
	kvDelete: () => Promise.resolve()
}))
vi.mock("@tanstack/react-router", () => {
	const navigate = vi.fn()
	const router = { history: { canGoBack: () => false, back: vi.fn() } }

	return {
		useNavigate: () => navigate,
		useRouter: () => router,
		useRouterState: ({ select }: { select: (state: unknown) => unknown }) =>
			select({ location: { state: { __TSR_key: "entry" }, href: "/videos" } }),
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
vi.mock("@/features/videos/hooks/useVideoLibraryView", () => ({ useVideoRootNames: () => new Map([["root", "Library"]]) }))
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
// The library client as the index's worker feeds it: art lookups and file rebuilds are worker calls, never requests.
vi.mock("@/features/videos/lib/index", () => ({
	VIDEO_CATALOG_KEY_PREFIX: ["videos", "catalog"],
	videoLibrary: {
		markTargets: () => Promise.resolve([]),
		titleArt: async (id: string, slot: string) => {
			const art = (await sdkApi.videoTitleArt(id, slot)) as { candidates: unknown[]; files: { uuid: string; size: number }[] }

			return {
				candidates: art.candidates,
				fileUuid: (file: number) => art.files[file]?.uuid ?? null,
				resolveFile: (file: number) => {
					const ref = art.files[file]

					return Promise.resolve(
						ref === undefined ? null : { uuid: ref.uuid, size: ref.size, file: { uuid: ref.uuid } as unknown as AnyFile }
					)
				}
			}
		},
		itemArt: () => Promise.resolve(null)
	}
}))

import { VideosHomeRows } from "@/features/videos/components/homeRows"
import { CARD_PROBE_DWELL_MS } from "@/features/videos/lib/probeService.logic"
import type { VideoLibraryRootView } from "@/features/videos/lib/libraryView.logic"
import { homeCatalog, movieItem, title } from "@/tests/support/videoHome"

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

class SizedResizeObserver implements ResizeObserver {
	private readonly callback: ResizeObserverCallback

	public constructor(callback: ResizeObserverCallback) {
		this.callback = callback
	}

	public observe(target: Element): void {
		this.callback([{ target, contentRect: { width: 1200, height: 800 } } as unknown as ResizeObserverEntry], this)
	}

	public unobserve(): void {
		// Nothing to stop.
	}

	public disconnect(): void {
		// Nothing to stop.
	}
}

// Reports `targets` intersecting to the observers whose options match `which`.
function intersect(which: (options: IntersectionObserverInit | undefined) => boolean, targets: (all: Element[]) => Element[]): void {
	for (const observer of observers.filter(entry => which(entry.options))) {
		const chosen = targets([...observer.targets])

		observer.callback(
			chosen.map(target => ({ target, isIntersecting: true, intersectionRatio: 1 }) as unknown as IntersectionObserverEntry),
			observer as unknown as IntersectionObserver
		)
	}
}

const isRowObserver = (options: IntersectionObserverInit | undefined) => options?.rootMargin === "100% 0px"
const isArtRowObserver = (options: IntersectionObserverInit | undefined) => options?.rootMargin === "50% 0px"
const isCardObserver = (options: IntersectionObserverInit | undefined) => options?.rootMargin === "0px 50%"
const isVisibleObserver = (options: IntersectionObserverInit | undefined) => options?.threshold === 0.5

function networkCalls(): number {
	return (
		sdkApi.readVideoProbe.mock.calls.length +
		sdkApi.makeArtFromRange.mock.calls.length +
		sdkApi.getLastDriveEventId.mock.calls.length +
		sdkApi.listVideoRoot.mock.calls.length +
		sdkApi.videoIndexCheck.mock.calls.length +
		sdkApi.listDirectory.mock.calls.length
	)
}

const ROOTS: VideoLibraryRootView[] = [{ uuid: "root", type: "auto", enabled: true, status: undefined, phase: "ready" }]

async function flush(): Promise<void> {
	await act(async () => {
		await vi.advanceTimersByTimeAsync(0)
	})
}

beforeEach(() => {
	Element.prototype.scrollIntoView = vi.fn()
	vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] })
	vi.stubGlobal("IntersectionObserver", FakeIntersectionObserver)
	vi.stubGlobal("ResizeObserver", SizedResizeObserver)
	observers.length = 0

	const titles = Array.from({ length: 30 }, (_, i) => title(`m${String(i).padStart(2, "0")}`, { dateAdded: i }))

	catalog.current = homeCatalog({ titles, items: titles.map(item => movieItem(item.id, { dateAdded: item.dateAdded })) })
	sdkApi.videoCatalogHome.mockImplementation((query: HomeQuery) => Promise.resolve(catalog.current?.home(query)))
	// Every poster comes from a sidecar image: stored nowhere yet, so only a download can make it.
	sdkApi.videoTitleArt.mockImplementation((id: string) =>
		Promise.resolve({
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
			files: [{ stableUuid: `${id}-p`, uuid: `${id}-p`, name: "poster.jpg", size: 1000 }]
		})
	)
	sdkApi.makeArtFromRange.mockResolvedValue({ type: "notImage" })
})

afterEach(() => {
	vi.useRealTimers()
})

describe("Videos home request counts", () => {
	it("renders from the catalog with no listing or probe, and downloads art only for cards on screen after the dwell", async () => {
		const scroll = document.createElement("div")
		const client = new QueryClient()
		const view = render(
			createElement(QueryClientProvider, { client }, createElement(VideosHomeRows, { roots: ROOTS, scrollElement: scroll }))
		)

		await flush()

		// The rows: one worker read, no request.
		expect(sdkApi.videoCatalogHome).toHaveBeenCalledTimes(1)
		expect(view.container.querySelectorAll("section")).toHaveLength(3)
		// The first row holds the keyboard cursor, so it mounts at once (the cards in view and the next); the rest wait
		// for the viewport.
		expect(view.container.querySelectorAll("[role=listitem]")).toHaveLength(8)

		// The first two rows come within a viewport; their cards mount, still reading nothing.
		act(() => {
			intersect(isRowObserver, all => all.slice(0, 2))
		})
		await flush()

		expect(view.container.querySelectorAll("[role=listitem]")).toHaveLength(16)
		expect(sdkApi.videoTitleArt).not.toHaveBeenCalled()

		// Only the first row is within half a viewport: only its cards may read art, once near its visible part.
		act(() => {
			intersect(isArtRowObserver, all => all.slice(0, 1))
		})
		await flush()

		expect(sdkApi.videoTitleArt).not.toHaveBeenCalled()

		// Seven cards of the first row near its visible part: their art is looked up (worker calls, local reads).
		act(() => {
			intersect(isCardObserver, all => all.slice(0, 7))
		})
		await flush()

		expect(sdkApi.videoTitleArt.mock.calls.length).toBeGreaterThanOrEqual(7)
		expect(networkCalls()).toBe(0)

		// Three of them on screen: nothing until the dwell has passed, then exactly their three downloads.
		act(() => {
			intersect(isVisibleObserver, all => all.slice(0, 3))
		})
		await act(async () => {
			await vi.advanceTimersByTimeAsync(CARD_PROBE_DWELL_MS - 1)
		})

		expect(networkCalls()).toBe(0)

		await act(async () => {
			await vi.advanceTimersByTimeAsync(200)
		})

		expect(sdkApi.makeArtFromRange).toHaveBeenCalledTimes(3)
		expect(sdkApi.readVideoProbe).not.toHaveBeenCalled()
		expect(sdkApi.getLastDriveEventId).not.toHaveBeenCalled()
		expect(sdkApi.listVideoRoot).not.toHaveBeenCalled()
		expect(sdkApi.videoIndexCheck).not.toHaveBeenCalled()
		expect(sdkApi.listDirectory).not.toHaveBeenCalled()

		view.unmount()
	})
})
