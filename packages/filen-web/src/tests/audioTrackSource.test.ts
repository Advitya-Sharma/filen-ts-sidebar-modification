import { afterEach, describe, expect, it, vi } from "vitest"
import type { AnyFile } from "@filen/sdk-rs"

const { downloadFileBytes, releasePreviewStream, streaming } = vi.hoisted(() => ({
	downloadFileBytes: vi.fn(),
	releasePreviewStream: vi.fn<(id: string) => void>(),
	streaming: { value: false }
}))

vi.mock("@/features/preview/lib/previewStream", () => ({
	isMediaStreamAvailable: () => streaming.value,
	previewStreamUrl: (_file: unknown, _name: string, _type: string, id: string) => Promise.resolve(`/sw/download/${id}`),
	releasePreviewStream
}))
vi.mock("@/lib/sdk/client", () => ({ sdkApi: { downloadFileBytes, cancelPreviewDownload: vi.fn() } }))

import { holdStreamRegistration, resolveTrackSource } from "@/features/audio/lib/bytes"
import type { QueueTrack } from "@/features/audio/store/audioQueue"

const track: QueueTrack = { uuid: "a", name: "a.mp3", mime: "audio/mpeg", contentType: null, file: {} as AnyFile }

describe("resolveTrackSource", () => {
	const urls: string[] = []

	afterEach(() => {
		for (const url of urls.splice(0)) {
			URL.revokeObjectURL(url)
		}
	})

	it("carries the downloaded Blob on a blob source, so a local read never fetches the object URL", async () => {
		const fetchSpy = vi.fn()

		vi.stubGlobal("fetch", fetchSpy)
		downloadFileBytes.mockResolvedValue(new Uint8Array([1, 2, 3]))

		const source = await resolveTrackSource(track, new AbortController().signal)

		if (source.kind !== "blob") {
			throw new Error("expected a blob source")
		}

		urls.push(source.url)

		expect(source.url.startsWith("blob:")).toBe(true)
		expect(source.blob.type).toBe("audio/mpeg")
		expect(new Uint8Array(await source.blob.arrayBuffer())).toEqual(new Uint8Array([1, 2, 3]))
		expect(fetchSpy).not.toHaveBeenCalled()
	})

	it("hands back stream registrations older than the playing and the warmed track", async () => {
		streaming.value = true

		try {
			const streamed = { ...track, contentType: "audio/mpeg" }
			const ids: string[] = []

			for (let index = 0; index < 4; index++) {
				const source = await resolveTrackSource(streamed, new AbortController().signal)

				ids.push(source.url.slice("/sw/download/".length))
			}

			expect(releasePreviewStream.mock.calls.map(([id]) => id)).toEqual([ids[0], ids[1]])

			// The bookkeeping alone: only ever the last two are kept.
			holdStreamRegistration("x")
			expect(releasePreviewStream).toHaveBeenLastCalledWith(ids[2])
		} finally {
			streaming.value = false
		}
	})
})
