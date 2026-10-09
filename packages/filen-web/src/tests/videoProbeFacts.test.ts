import { beforeEach, describe, expect, it, vi } from "vitest"
import type { ProbeRecord } from "@/features/videos/lib/probeService.logic"

// The grids' runtimes and resolutions: the stored probe records are scanned at most once per session, and the probes
// this session runs keep the facts current without another scan.

const { kvEntriesJson, recorded } = vi.hoisted(() => ({
	kvEntriesJson: vi.fn(),
	recorded: { listener: null as ((uuid: string, record: ProbeRecord) => void) | null }
}))

vi.mock("@/lib/storage/adapter", () => ({ kvEntriesJson }))
vi.mock("@/features/videos/lib/probeService", () => ({
	probeService: {
		onRecord: (listener: (uuid: string, record: ProbeRecord) => void) => {
			recorded.listener = listener

			return () => undefined
		}
	}
}))

import { probeFacts } from "@/features/videos/lib/probeFacts"

function record(durationSec: number, h: number): ProbeRecord {
	return {
		probe: {
			v: 1,
			container: "mkv",
			durationSec,
			index: "cues",
			video: { codec: null, fourcc: "", codecString: null, w: 0, h, par: 1, rotation: 0, bitDepth: null, transfer: "sdr" },
			audio: [],
			subs: [],
			chapters: [],
			cover: null,
			coverLand: null,
			partial: false
		},
		tiers: { card: true, full: false, cover: false },
		at: 1
	}
}

beforeEach(() => {
	probeFacts.reset()
	kvEntriesJson.mockReset()
	kvEntriesJson.mockResolvedValue([["videos.v1.probe.a", record(6000, 1080)]])
})

describe("probeFacts", () => {
	it("scans the stored records once, then follows new probes", async () => {
		expect((await probeFacts.load()).get("a")).toEqual({ durationSec: 6000, height: 1080 })

		recorded.listener?.("b", record(120, 2160))

		const facts = await probeFacts.load()

		expect(facts.get("b")).toEqual({ durationSec: 120, height: 2160 })
		expect(kvEntriesJson).toHaveBeenCalledTimes(1)
	})

	it("scans again only after a sign-out", async () => {
		await Promise.all([probeFacts.load(), probeFacts.load()])
		probeFacts.reset()
		await probeFacts.load()

		expect(kvEntriesJson).toHaveBeenCalledTimes(2)
	})
})
