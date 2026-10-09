// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest"
import type { File } from "@filen/sdk-rs"
import { driveFile, hexUuid } from "@/tests/support/videoDrive"

const { playableVideoFiles, videoTitleArt } = vi.hoisted(() => ({
	playableVideoFiles: vi.fn<(stableUuids: string[]) => Promise<File[]>>(),
	videoTitleArt: vi.fn()
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi: { playableVideoFiles, videoTitleArt } }))
vi.mock("@/queries/client", async () => ({ queryClient: (await import("@/tests/testQueryClient")).createTestQueryClient() }))
vi.mock("@/queries/persist", () => ({ persister: { persistQuery: () => Promise.resolve() } }))

import { videoLibrary } from "@/features/videos/lib/index"

const PARENT = hexUuid()
const files = [driveFile(PARENT, "poster.jpg"), driveFile(PARENT, "fanart.jpg"), driveFile(PARENT, "Heat (1995).mkv")]

function ref(file: File) {
	return { stableUuid: file.stableUUID ?? file.uuid, uuid: file.uuid, name: "x", size: 1_000 }
}

beforeEach(() => {
	videoLibrary.reset()
	playableVideoFiles.mockImplementation(stableUuids => Promise.resolve(files.filter(file => stableUuids.includes(file.stableUUID ?? ""))))
})

describe("art file resolution", () => {
	it("asks the worker once for every file asked for in the same task, and keeps nothing keyed afterwards", async () => {
		const resolved = await Promise.all([...files, ...files].map(file => videoLibrary.resolveArtFile(ref(file))))

		expect(playableVideoFiles).toHaveBeenCalledTimes(1)
		expect(resolved.map(entry => entry?.uuid)).toEqual([...files, ...files].map(file => file.uuid))
		expect(resolved[0]?.file).toBe(files[0])

		await videoLibrary.resolveArtFile(ref(files[1] ?? driveFile(PARENT, "x.jpg")))

		expect(playableVideoFiles).toHaveBeenCalledTimes(2)
	})

	it("resolves a ref the index no longer holds in that version to nothing, and asks again next time", async () => {
		const stale = { ...ref(files[0] ?? driveFile(PARENT, "x.jpg")), uuid: hexUuid() }

		await expect(videoLibrary.resolveArtFile(stale)).resolves.toBeNull()
		await videoLibrary.resolveArtFile(stale)

		expect(playableVideoFiles).toHaveBeenCalledTimes(2)
	})

	it("hands the art loader a resolver over the candidates' keyless refs", async () => {
		videoTitleArt.mockResolvedValue({
			candidates: [{ source: "sidecar", file: 0 }, { source: "typographic" }],
			files: [ref(files[0] ?? driveFile(PARENT, "x.jpg"))]
		})

		const art = await videoLibrary.titleArt("title", "p")

		expect(art?.candidates).toHaveLength(2)
		expect(art?.fileUuid(0)).toBe(files[0]?.uuid)
		expect(art?.fileUuid(5)).toBeNull()
		expect(playableVideoFiles).not.toHaveBeenCalled()
		await expect(art?.resolveFile(0)).resolves.toMatchObject({ uuid: files[0]?.uuid })
		await expect(art?.resolveFile(5)).resolves.toBeNull()
	})
})
