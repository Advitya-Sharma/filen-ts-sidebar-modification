// @vitest-environment jsdom

import { createElement, type ComponentType, type ReactNode } from "react"
import { mkdirSync, rmSync, writeFileSync } from "node:fs"
import { dirname } from "node:path"
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from "vitest"
import { act, fireEvent, render, waitFor } from "@testing-library/react"
import { QueryClient, QueryClientProvider } from "@tanstack/react-query"
import { transformFileSync } from "@babel/core"
import "@/lib/i18n"
import type { VideoLibraryRootView } from "@/features/videos/lib/libraryView.logic"
import type { HomeQuery } from "@/workers/videoIndex/catalog.logic"
import type { VideoCatalog } from "@/workers/videoIndex/catalog.logic"

// What a keyboard cursor move on Home costs in renders: only the card losing and the card gaining the tab stop. The
// tests run uncompiled, so the home rows and the hooks they use are compiled here the way vite.config.ts compiles
// them, into a git-ignored directory inside the package (so the test runner transforms them and resolves aliases).

const { catalog, renders, compiledPath } = vi.hoisted(() => ({
	compiledPath: (name: string) => `${process.cwd()}/test-results/videos-compiled-${String(process.pid)}/${name}.jsx`,
	catalog: { current: null as VideoCatalog | null },
	renders: { count: 0 }
}))

// The rows come from the library's worker; here, the same catalog code in-process.
vi.mock("@/lib/sdk/client", () => ({
	sdkApi: { videoCatalogHome: (query: HomeQuery) => Promise.resolve(catalog.current?.home(query)) }
}))
vi.mock("@/features/videos/workers/frame.worker.ts?worker", () => ({ default: vi.fn() }))
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
vi.mock("@/features/videos/hooks/useWatchStateLoaded", () => import(/* @vite-ignore */ compiledPath("useWatchStateLoaded")))
vi.mock("@/features/videos/hooks/useVideoCardActions", () => import(/* @vite-ignore */ compiledPath("useVideoCardActions")))
vi.mock("@/features/videos/hooks/useCardFocus", () => import(/* @vite-ignore */ compiledPath("useCardFocus")))
vi.mock("@/features/videos/components/videoCard", () => import(/* @vite-ignore */ compiledPath("videoCard")))
vi.mock("@/features/videos/hooks/useVideoLibraryView", () => ({ useVideoRootNames: () => new Map([["root", "Library"]]) }))
// Every VideoCard render calls this once.
vi.mock("@/features/videos/hooks/useVideoArt", () => ({
	useVideoArt: () => {
		renders.count++

		return { kind: "none" }
	}
}))
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
vi.mock("@/features/videos/lib/index", () => ({
	VIDEO_CATALOG_KEY_PREFIX: ["videos", "catalog"],
	videoLibrary: { markTargets: () => Promise.resolve([]) }
}))

import { homeCatalog, movieItem, title } from "@/tests/support/videoHome"

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

class InertIntersectionObserver {
	public observe(): void {
		// Never reports: the rows holding the cursor mount regardless.
	}

	public unobserve(): void {
		// Nothing to stop.
	}

	public disconnect(): void {
		// Nothing to stop.
	}
}

const ROOTS: VideoLibraryRootView[] = [{ uuid: "root", type: "auto", enabled: true, status: undefined, phase: "ready" }]

let VideosHomeRows: ComponentType<{ roots: readonly VideoLibraryRootView[]; scrollElement: HTMLElement | null }>

const COMPILED: readonly (readonly [string, string])[] = [
	["src/features/videos/components/homeRows.tsx", "homeRows"],
	["src/features/videos/components/videoCard.tsx", "videoCard"],
	["src/features/videos/hooks/useWatchStateLoaded.ts", "useWatchStateLoaded"],
	["src/features/videos/hooks/useVideoCardActions.ts", "useVideoCardActions"],
	["src/features/videos/hooks/useCardFocus.ts", "useCardFocus"]
]

beforeAll(async () => {
	mkdirSync(dirname(compiledPath("x")), { recursive: true })

	for (const [source, name] of COMPILED) {
		const result = transformFileSync(source, {
			babelrc: false,
			configFile: false,
			presets: ["@babel/preset-typescript"],
			plugins: ["@babel/plugin-syntax-jsx", "babel-plugin-react-compiler"]
		})

		writeFileSync(compiledPath(name), result?.code ?? "")
	}

	const compiled = (await import(/* @vite-ignore */ compiledPath("homeRows"))) as { VideosHomeRows: typeof VideosHomeRows }

	VideosHomeRows = compiled.VideosHomeRows
})

afterAll(() => {
	rmSync(dirname(compiledPath("x")), { recursive: true, force: true })
})

beforeEach(() => {
	Element.prototype.scrollIntoView = vi.fn()
	vi.stubGlobal("IntersectionObserver", InertIntersectionObserver)
	vi.stubGlobal("ResizeObserver", SizedResizeObserver)

	const titles = Array.from({ length: 30 }, (_, i) => title(`m${String(i).padStart(2, "0")}`, { dateAdded: i }))

	catalog.current = homeCatalog({ titles, items: titles.map(item => movieItem(item.id, { dateAdded: item.dateAdded })) })
	renders.count = 0
})

afterEach(() => {
	vi.unstubAllGlobals()
})

describe("Videos home render counts", () => {
	it("re-renders only the two cards a cursor move touches", async () => {
		const view = render(
			createElement(
				QueryClientProvider,
				{ client: new QueryClient() },
				createElement(VideosHomeRows, { roots: ROOTS, scrollElement: document.createElement("div") })
			)
		)

		await waitFor(() => {
			expect(view.container.querySelector('[data-video-card][tabindex="0"]')).not.toBeNull()
		})

		const first = view.container.querySelector<HTMLElement>('[data-video-card][tabindex="0"]')

		expect(first).not.toBeNull()
		// The cards in view and the next; focusing the row mounts the rest.
		expect(view.container.querySelectorAll("[data-video-card]")).toHaveLength(8)

		act(() => {
			first?.focus()
		})

		await waitFor(() => {
			expect(view.container.querySelectorAll("[data-video-card]")).toHaveLength(30)
		})

		const before = renders.count

		act(() => {
			if (first !== null) {
				fireEvent.keyDown(first, { key: "ArrowRight" })
			}
		})

		expect(view.container.querySelector('[data-video-card][tabindex="0"]')?.getAttribute("aria-label")).toMatch(/^m28/)
		expect(renders.count - before).toBe(2)

		view.unmount()
	})
})
