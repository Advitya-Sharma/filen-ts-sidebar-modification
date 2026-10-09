import { describe, expect, it } from "vitest"
import { KeyIndex, MAX_SEEK_SCAN_BYTES, seekReach } from "@/workers/remux/keyIndex"

describe("KeyIndex", () => {
	it("finds the last key packet at or before a time", () => {
		const index = new KeyIndex()

		for (const time of [0.021, 10.021, 20.021, 30.021]) {
			index.add(time)
		}

		expect(index.atOrBefore(25)).toBe(20.021)
		expect(index.atOrBefore(20.021)).toBe(20.021)
		expect(index.atOrBefore(100)).toBe(30.021)
		expect(index.atOrBefore(0)).toBeNull()
	})

	it("keeps order and drops duplicates when a pump re-reads an earlier stretch", () => {
		const index = new KeyIndex()

		for (const time of [20, 30, 10, 20, 0]) {
			index.add(time)
		}

		expect([index.atOrBefore(5), index.atOrBefore(15), index.atOrBefore(25), index.atOrBefore(35)]).toEqual([0, 10, 20, 30])
	})
})

describe("seekReach", () => {
	const MiB = 1_048_576

	it("seeks freely within what has been scanned", () => {
		expect(seekReach({ time: 50, scannedUntil: 60, size: 150 * MiB, duration: 10_800 })).toEqual({ time: 50, clamped: false })
	})

	it("scans on to a target whose clusters fit the budget, a small file entirely", () => {
		expect(seekReach({ time: 20, scannedUntil: 0, size: 6 * MiB, duration: 30 })).toEqual({ time: 20, clamped: false })
		expect(seekReach({ time: 25, scannedUntil: 0, size: 6 * MiB, duration: 0 })).toEqual({ time: 25, clamped: false })
	})

	it("stops a far seek at the end of what has been scanned", () => {
		expect(seekReach({ time: 3600, scannedUntil: 60, size: 150 * MiB, duration: 10_800 })).toEqual({ time: 60, clamped: true })
		expect(seekReach({ time: 25, scannedUntil: 5, size: MAX_SEEK_SCAN_BYTES + 1, duration: 0 })).toEqual({ time: 5, clamped: true })
	})
})
