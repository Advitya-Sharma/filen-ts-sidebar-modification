import { describe, expect, it } from "vitest"
import { boundedVerdicts, isStoredVerdict, VERDICT_KV_PREFIX, VERDICT_MAX_RECORDS } from "@filen/shared"

describe("verdict rows", () => {
	it("live under a versioned prefix, the newest few thousand", () => {
		expect(VERDICT_KV_PREFIX).toBe("files.v1.kind.")
		expect(VERDICT_MAX_RECORDS).toBe(4096)
	})
})

describe("boundedVerdicts", () => {
	it("keeps the newest, the session's winning, and drops the rest", () => {
		const { keep, drop } = boundedVerdicts(
			[
				["a", { v: "audio", at: 1 }],
				["b", { v: "video", at: 2 }],
				["c", { v: "audio", at: 3 }]
			],
			[["a", { v: "transportStream", at: 4 }]],
			2
		)

		expect(keep).toEqual([
			["c", { v: "audio", at: 3 }],
			["a", { v: "transportStream", at: 4 }]
		])
		expect(drop).toEqual(["b"])
	})
})

describe("isStoredVerdict", () => {
	it("accepts only a known verdict with a finite time", () => {
		expect(isStoredVerdict({ v: "audio", at: 1 })).toBe(true)
		expect(isStoredVerdict({ v: "notTransportStream", at: 0 })).toBe(true)
		expect(isStoredVerdict({ v: "picture", at: 1 })).toBe(false)
		expect(isStoredVerdict({ v: "audio", at: Number.NaN })).toBe(false)
		expect(isStoredVerdict({ v: "audio" })).toBe(false)
		expect(isStoredVerdict(null)).toBe(false)
		expect(isStoredVerdict("audio")).toBe(false)
	})
})
