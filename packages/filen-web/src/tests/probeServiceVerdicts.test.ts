import { describe, expect, it, vi } from "vitest"

// The Videos probe service feeds file-kind verdicts in every session, from the moment it exists: Videos
// browsing alone settles an ambiguous .ts, .webm or .ogg, with no request of its own.

const adopt = vi.hoisted(() => vi.fn())

vi.mock("@/features/drive/lib/contentVerdicts", () => ({ adoptProbeRecords: adopt }))
vi.mock("@/features/videos/lib/records", () => ({
	probeRecords: { get: vi.fn(), refresh: vi.fn(), put: vi.fn(), peek: vi.fn(), reset: vi.fn() }
}))

const { probeService } = await import("@/features/videos/lib/probeService")
// Read now: the suite clears mock calls before each test.
const adoptedAtImport = adopt.mock.calls.map(([service]: unknown[]) => service)

describe("probeService", () => {
	it("is adopted by the file-kind verdicts when it is created", () => {
		expect(adoptedAtImport).toEqual([probeService])
	})
})
