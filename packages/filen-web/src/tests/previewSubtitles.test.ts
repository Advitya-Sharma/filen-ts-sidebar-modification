import { afterEach, describe, expect, it, vi } from "vitest"
import type { DriveItem } from "@/features/drive/lib/item"

// A preview's sidecar subtitles come from a listing already held (never a fetch) and cost one read when
// chosen: through the session for the drive, through the unauthenticated client for a public link.

const mocks = vi.hoisted(() => ({
	downloadFileBytes: vi.fn<() => Promise<Uint8Array>>(),
	downloadLinkedFileBytesAnon: vi.fn<() => Promise<Uint8Array>>(),
	cancelPreviewDownload: vi.fn(() => Promise.resolve()),
	queries: [] as { queryKey: unknown[]; state: { data: unknown } }[]
}))

vi.mock("@/lib/sdk/client", () => ({
	sdkApi: {
		downloadFileBytes: mocks.downloadFileBytes,
		downloadLinkedFileBytesAnon: mocks.downloadLinkedFileBytesAnon,
		cancelPreviewDownload: mocks.cancelPreviewDownload
	}
}))
vi.mock("@/queries/patch", () => ({ cachedQueriesWithPrefix: () => mocks.queries }))
vi.mock("@/features/drive/queries/drive", () => ({ DRIVE_LISTING_KEY_PREFIX: ["drive", "listing"] }))

const { cachedSiblings, loadSidecar, previewSidecars } = await import("@/features/preview/lib/previewSubtitles")
const { linkedFileItem } = await import("@/tests/fixtures/sdk")

const SRT = new TextEncoder().encode("1\n00:00:01,000 --> 00:00:02,000\nHello\n")

function listing(variant: string, uuid: string | null, items: DriveItem[]): { queryKey: unknown[]; state: { data: unknown } } {
	return { queryKey: ["drive", "listing", { variant, uuid }], state: { data: items } }
}

afterEach(() => {
	vi.clearAllMocks()
	mocks.queries = []
})

describe("cachedSiblings", () => {
	it("reads the cached directory listing holding the video, never a flat one", () => {
		const video = linkedFileItem("Movie.mkv")
		const subtitle = linkedFileItem("Movie.en.srt")

		mocks.queries = [listing("recents", null, [video]), listing("drive", "dir", [video, subtitle])]

		expect(cachedSiblings(video.data.uuid)).toEqual([video, subtitle])
	})

	it("is empty when no directory listing holds it", () => {
		const video = linkedFileItem("Movie.mkv")

		mocks.queries = [listing("favorites", null, [video]), listing("sharedIn", null, [video])]

		expect(cachedSiblings(video.data.uuid)).toEqual([])
	})
})

describe("loadSidecar", () => {
	it("reads a chosen subtitle file once through the session and parses it", async () => {
		const subtitle = linkedFileItem("Movie.en.srt", { size: BigInt(SRT.length) })
		const [found] = previewSidecars("Movie.mkv", [linkedFileItem("Movie.mkv"), subtitle])

		mocks.downloadFileBytes.mockResolvedValue(SRT)

		if (found === undefined) {
			throw new Error("no sidecar found")
		}

		const result = await loadSidecar(found, "authed", new AbortController().signal)

		expect(result.status === "ok" ? result.cues.texts : []).toEqual(["Hello"])
		expect(mocks.downloadFileBytes).toHaveBeenCalledOnce()
		expect(mocks.downloadLinkedFileBytesAnon).not.toHaveBeenCalled()
	})

	it("reads a public link's subtitle through the unauthenticated client", async () => {
		const [found] = previewSidecars("Movie.mkv", [linkedFileItem("Movie.srt", { size: BigInt(SRT.length) })])

		mocks.downloadLinkedFileBytesAnon.mockResolvedValue(SRT)

		if (found === undefined) {
			throw new Error("no sidecar found")
		}

		expect((await loadSidecar(found, "anon", new AbortController().signal)).status).toBe("ok")
		expect(mocks.downloadLinkedFileBytesAnon).toHaveBeenCalledOnce()
		expect(mocks.downloadFileBytes).not.toHaveBeenCalled()
	})

	it("calls a file without cues unreadable", async () => {
		const [found] = previewSidecars("Movie.mkv", [linkedFileItem("Movie.srt", { size: 5n })])

		mocks.downloadFileBytes.mockResolvedValue(new TextEncoder().encode("junk"))

		if (found === undefined) {
			throw new Error("no sidecar found")
		}

		expect((await loadSidecar(found, "authed", new AbortController().signal)).status).toBe("unreadable")
	})
})
