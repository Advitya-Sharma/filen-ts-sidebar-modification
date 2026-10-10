// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { renderHook, waitFor } from "@testing-library/react"

const mocks = vi.hoisted(() => ({
	load: vi.fn(),
	loadLocal: vi.fn(),
	titleArt: vi.fn()
}))

vi.mock("@/lib/storage/opfsBlobStore", () => ({ readStoreBlob: () => Promise.resolve(null) }))
vi.mock("@/features/videos/lib/index", () => ({ videoLibrary: { titleArt: mocks.titleArt } }))
vi.mock("@/features/videos/lib/artwork", () => ({ artLoader: { load: mocks.load, loadLocal: mocks.loadLocal } }))

import { useUpNextStill } from "@/features/videos/components/player/useTitleStill"

afterEach(() => {
	vi.clearAllMocks()
	vi.unstubAllGlobals()
})

const ART = {
	candidates: [{ source: "sidecar", file: 0 }],
	fileUuid: () => "still-u",
	resolveFile: () => Promise.resolve(null)
}

describe("the Up Next still", () => {
	it("never downloads or generates the title's still for the card", async () => {
		mocks.titleArt.mockResolvedValue(ART)
		mocks.loadLocal.mockResolvedValue("network")

		const { result } = renderHook(() => useUpNextStill({ titleId: "t", title: "Show", fileUuid: "next-u", enabled: true }))

		await waitFor(() => {
			expect(mocks.loadLocal).toHaveBeenCalledTimes(1)
		})
		await Promise.resolve()

		expect(mocks.load).not.toHaveBeenCalled()
		expect(result.current).toBeNull()
	})

	it("shows the title's still stored earlier", async () => {
		vi.stubGlobal("URL", Object.assign(URL, { createObjectURL: () => "blob:still", revokeObjectURL: () => undefined }))
		mocks.titleArt.mockResolvedValue(ART)
		mocks.loadLocal.mockResolvedValue({ type: "image", blob: new Blob(["x"]) })

		const { result } = renderHook(() => useUpNextStill({ titleId: "t", title: "Show", fileUuid: null, enabled: true }))

		await waitFor(() => {
			expect(result.current).toBe("blob:still")
		})
		expect(mocks.load).not.toHaveBeenCalled()
	})
})
