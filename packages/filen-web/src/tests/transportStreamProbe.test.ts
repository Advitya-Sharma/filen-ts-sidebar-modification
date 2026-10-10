import { beforeEach, describe, expect, it, vi } from "vitest"
import type { AnyFile } from "@filen/sdk-rs"
import type { VideoProbe } from "@filen/shared/videos"

// A `.ts` gate's probe is read at the tier the player's ladder needs, so a transport stream's head is read
// once: the drive's through the probe service's cache, a public link's handed to the player directly.

const mocks = vi.hoisted(() => ({
	readLinkedVideoProbeAnon: vi.fn<(file: AnyFile, tier: string, token: string) => Promise<VideoProbe | null>>(),
	cancelPreviewDownload: vi.fn(() => Promise.resolve()),
	request: vi.fn<(target: unknown, tier: string, options: unknown) => Promise<{ probe: VideoProbe | null } | null>>()
}))

vi.mock("@/lib/sdk/client", () => ({
	sdkApi: { readLinkedVideoProbeAnon: mocks.readLinkedVideoProbeAnon, cancelPreviewDownload: mocks.cancelPreviewDownload }
}))
vi.mock("@/features/videos/lib/probeService", () => ({ probeService: { request: mocks.request } }))

const { probeTransportStream, takeLinkedTransportStreamProbe } = await import("@/features/preview/lib/transportStreamProbe")

function fileOf(uuid: string): AnyFile {
	return { uuid, size: 10_000_000n } as unknown as AnyFile
}

function probeOf(container: VideoProbe["container"]): VideoProbe {
	return { container } as VideoProbe
}

beforeEach(() => {
	vi.clearAllMocks()
	takeLinkedTransportStreamProbe("a")
	takeLinkedTransportStreamProbe("b")
})

describe("probeTransportStream", () => {
	it("asks the drive's probe service for the full record the player reads next", async () => {
		mocks.request.mockResolvedValue({ probe: probeOf("ts") })

		await expect(probeTransportStream(fileOf("a"), "authed", new AbortController().signal)).resolves.toBe(true)
		expect(mocks.request.mock.calls[0]?.[1]).toBe("full")
		expect(takeLinkedTransportStreamProbe("a")).toBeNull()
	})

	it("hands a public link's transport stream probe to the player once, for that file only", async () => {
		const probe = probeOf("ts")

		mocks.readLinkedVideoProbeAnon.mockResolvedValue(probe)

		await expect(probeTransportStream(fileOf("a"), "anon", new AbortController().signal)).resolves.toBe(true)
		expect(mocks.readLinkedVideoProbeAnon.mock.calls[0]?.[1]).toBe("full")
		expect(takeLinkedTransportStreamProbe("b")).toBeNull()
		expect(takeLinkedTransportStreamProbe("a")).toBe(probe)
		expect(takeLinkedTransportStreamProbe("a")).toBeNull()
	})

	it("keeps nothing for a public link file that is not a transport stream, or whose probe failed", async () => {
		mocks.readLinkedVideoProbeAnon
			.mockResolvedValueOnce(probeOf("mp4"))
			.mockResolvedValueOnce(null)
			.mockRejectedValueOnce(new Error("offline"))

		await expect(probeTransportStream(fileOf("a"), "anon", new AbortController().signal)).resolves.toBe(false)
		await expect(probeTransportStream(fileOf("a"), "anon", new AbortController().signal)).resolves.toBe(false)
		await expect(probeTransportStream(fileOf("a"), "anon", new AbortController().signal)).resolves.toBeNull()
		expect(takeLinkedTransportStreamProbe("a")).toBeNull()
	})
})
