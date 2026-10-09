import { describe, expect, it } from "vitest"
import { hevcKeyIsIdr, hevcLengthSize, IDR_REACH_BYTES, IDR_REACH_SECONDS, idrReachSeconds } from "@/workers/remux/hevc.logic"

// HEVC NAL unit header: forbidden bit, 6-bit type, 6-bit layer id, 3-bit temporal id.
function nal(type: number, payload = 4): number[] {
	return [(type << 1) & 0x7e, 0x01, ...new Array<number>(payload).fill(0xaa)]
}

function lengthPrefixed(units: number[][], size = 4): Uint8Array {
	const out: number[] = []

	for (const unit of units) {
		const length = unit.length

		for (let shift = size - 1; shift >= 0; shift--) {
			out.push((length >> (shift * 8)) & 0xff)
		}

		out.push(...unit)
	}

	return new Uint8Array(out)
}

const VPS = 32
const SPS = 33
const PPS = 34
const SEI = 39
const CRA = 21
const IDR_W_RADL = 19
const IDR_N_LP = 20

describe("hevcKeyIsIdr", () => {
	it("looks past parameter sets and SEI to the first slice", () => {
		expect(hevcKeyIsIdr(lengthPrefixed([nal(VPS), nal(SPS), nal(PPS), nal(SEI), nal(IDR_W_RADL)]), 4)).toBe(true)
		expect(hevcKeyIsIdr(lengthPrefixed([nal(VPS), nal(SPS), nal(PPS), nal(CRA)]), 4)).toBe(false)
		expect(hevcKeyIsIdr(lengthPrefixed([nal(IDR_N_LP)], 2), 2)).toBe(true)
	})

	it("reads Annex B start codes when there is no configuration record", () => {
		const annexB = new Uint8Array([0, 0, 0, 1, ...nal(SPS), 0, 0, 1, ...nal(CRA)])

		expect(hevcKeyIsIdr(annexB, null)).toBe(false)
		expect(hevcKeyIsIdr(new Uint8Array([0, 0, 1, ...nal(IDR_W_RADL)]), null)).toBe(true)
	})

	it("does not refuse an access unit whose slices it cannot find", () => {
		expect(hevcKeyIsIdr(new Uint8Array([0, 0, 0]), 4)).toBe(true)
		expect(hevcKeyIsIdr(lengthPrefixed([nal(SPS)]), 4)).toBe(true)
	})
})

describe("hevcLengthSize", () => {
	it("reads lengthSizeMinusOne from the configuration record", () => {
		const record = new Uint8Array(23)

		record[21] = 0xf3
		expect(hevcLengthSize(record)).toBe(4)
		record[21] = 0xf1
		expect(hevcLengthSize(record)).toBe(2)
		expect(hevcLengthSize(null)).toBeNull()
		expect(hevcLengthSize(new Uint8Array(4))).toBeNull()
	})
})

describe("idrReachSeconds", () => {
	it("caps the reach in seconds and in bytes at the file's bitrate", () => {
		expect(idrReachSeconds(0)).toBe(IDR_REACH_SECONDS)
		expect(idrReachSeconds(1_000_000)).toBe(IDR_REACH_SECONDS)
		expect(idrReachSeconds(80_000_000)).toBeCloseTo((IDR_REACH_BYTES * 8) / 80_000_000)
	})
})
