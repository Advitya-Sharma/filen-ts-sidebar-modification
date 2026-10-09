import { describe, expect, it } from "vitest"
import { ac3Layout, downmixToStereo, dtsLayout, layoutFor, type Speaker } from "@/workers/remux/channelLayout.logic"

// An AC-3 syncframe head: sync, CRC, fscod/frmsizecod, bsid 8, then acmod and what follows it.
function ac3Frame(modeByte: number): Uint8Array {
	return Uint8Array.of(0x0b, 0x77, 0, 0, 0x14, 8 << 3, modeByte, 0)
}

// An E-AC-3 syncframe head: bsid 16, acmod and lfeon in byte 4.
function eac3Frame(acmod: number, lfeon: number): Uint8Array {
	return Uint8Array.of(0x0b, 0x77, 0, 0, (acmod << 1) | lfeon, 16 << 3, 0, 0)
}

// A DTS core frame header with the given AMODE, extension and LFF fields.
function dtsFrame(params: { amode: number; extension: boolean; extensionId: number; lff: number }): Uint8Array {
	const bits: number[] = []
	const push = (value: number, width: number) => {
		for (let bit = width - 1; bit >= 0; bit--) {
			bits.push((value >> bit) & 1)
		}
	}

	push(0x7ffe, 16)
	push(0x8001, 16)
	push(0, 1 + 5 + 1 + 7 + 14)
	push(params.amode, 6)
	push(0, 4 + 5 + 1 + 1 + 1 + 1 + 1)
	push(params.extensionId, 3)
	push(params.extension ? 1 : 0, 1)
	push(0, 1)
	push(params.lff, 2)

	while (bits.length % 8 !== 0) {
		bits.push(0)
	}

	return Uint8Array.from({ length: bits.length / 8 }, (_, byte) =>
		bits.slice(byte * 8, byte * 8 + 8).reduce((value, bit) => (value << 1) | bit, 0)
	)
}

function constant(channels: number, value: (channel: number) => number): Float32Array[] {
	return Array.from({ length: channels }, (_, channel) => Float32Array.of(value(channel)))
}

describe("ac3Layout", () => {
	it("reads acmod and lfeon past the mix-level fields each mode carries", () => {
		// 3/2 + LFE: acmod 111, cmixlev, surmixlev, lfeon 1.
		expect(ac3Layout(ac3Frame(0b11100001))).toEqual(["FL", "FR", "FC", "LFE", "SL", "SR"])
		// 3/1: acmod 101, cmixlev, surmixlev, lfeon 0.
		expect(ac3Layout(ac3Frame(0b10100000))).toEqual(["FL", "FR", "FC", "BC"])
		// 2/2 + LFE: acmod 110, surmixlev, lfeon 1.
		expect(ac3Layout(ac3Frame(0b11000100))).toEqual(["FL", "FR", "LFE", "SL", "SR"])
		// 2/0 + LFE: acmod 010, dsurmod, lfeon 1.
		expect(ac3Layout(ac3Frame(0b01000100))).toEqual(["FL", "FR", "LFE"])
	})

	it("reads E-AC-3 and refuses what is no syncframe", () => {
		expect(ac3Layout(eac3Frame(7, 1))).toEqual(["FL", "FR", "FC", "LFE", "SL", "SR"])
		expect(ac3Layout(eac3Frame(4, 0))).toEqual(["FL", "FR", "BC"])
		expect(ac3Layout(Uint8Array.of(1, 2, 3, 4, 5, 6, 7, 8))).toBeNull()
	})
})

describe("dtsLayout", () => {
	it("reads AMODE, LFE and the XCh back centre of DTS-ES", () => {
		expect(dtsLayout(dtsFrame({ amode: 9, extension: false, extensionId: 0, lff: 1 }))).toEqual(["FL", "FR", "FC", "LFE", "SL", "SR"])
		expect(dtsLayout(dtsFrame({ amode: 9, extension: true, extensionId: 0, lff: 2 }))).toEqual([
			"FL",
			"FR",
			"FC",
			"LFE",
			"BC",
			"SL",
			"SR"
		])
		// X96 is no extra channel.
		expect(dtsLayout(dtsFrame({ amode: 7, extension: true, extensionId: 2, lff: 0 }))).toEqual(["FL", "FR", "FC", "BC"])
		expect(dtsLayout(dtsFrame({ amode: 12, extension: false, extensionId: 0, lff: 0 }))).toBeNull()
	})
})

describe("layoutFor", () => {
	it("uses the header when it names as many channels as were decoded", () => {
		expect(layoutFor(["FL", "FR", "FC", "BC"], 4)).toEqual(["FL", "FR", "FC", "BC"])
	})

	it("adds the back pair a dependent substream decodes on top of the header's 5.1", () => {
		expect(layoutFor(["FL", "FR", "FC", "LFE", "SL", "SR"], 8)).toEqual(["FL", "FR", "FC", "LFE", "BL", "BR", "SL", "SR"])
	})

	it("falls back to FFmpeg's default layout for the count", () => {
		expect(layoutFor(null, 6)).toEqual(["FL", "FR", "FC", "LFE", "SL", "SR"])
		expect(layoutFor(["FL", "FR"], 6)).toEqual(["FL", "FR", "FC", "LFE", "SL", "SR"])
	})
})

describe("downmixToStereo", () => {
	const at = (planes: [Float32Array, Float32Array]) => [planes[0][0] ?? Number.NaN, planes[1][0] ?? Number.NaN]

	it("keeps 5.1 dialogue in the centre at -3 dB on both sides, not attenuated by the channel count", () => {
		const layout: Speaker[] = ["FL", "FR", "FC", "LFE", "SL", "SR"]
		const [left, right] = at(
			downmixToStereo(
				constant(6, channel => (channel === 2 ? 0.5 : 0)),
				layout
			)
		)

		expect(left).toBeCloseTo(0.5 * Math.SQRT1_2, 5)
		expect(right).toBeCloseTo(0.5 * Math.SQRT1_2, 5)
	})

	it("drops LFE", () => {
		const layout: Speaker[] = ["FL", "FR", "FC", "LFE", "SL", "SR"]

		expect(
			at(
				downmixToStereo(
					constant(6, channel => (channel === 3 ? 0.8 : 0)),
					layout
				)
			)
		).toEqual([0, 0])
	})

	it("routes surrounds by position: 2/2 sides to their side, a 3/1 single surround to both at -6 dB", () => {
		const quad: Speaker[] = ["FL", "FR", "SL", "SR"]
		const [quadLeft, quadRight] = at(
			downmixToStereo(
				constant(4, channel => (channel === 2 ? 0.5 : 0)),
				quad
			)
		)

		expect(quadLeft).toBeCloseTo(0.5 * Math.SQRT1_2, 5)
		expect(quadRight).toBe(0)

		const threeOne: Speaker[] = ["FL", "FR", "FC", "BC"]
		const [left, right] = at(
			downmixToStereo(
				constant(4, channel => (channel === 3 ? 0.8 : 0)),
				threeOne
			)
		)

		expect(left).toBeCloseTo(0.4, 5)
		expect(right).toBeCloseTo(0.4, 5)
	})

	it("limits peaks instead of clipping when every channel is at full scale", () => {
		const layout: Speaker[] = ["FL", "FR", "FC", "LFE", "BC", "SL", "SR"]
		const [left, right] = at(
			downmixToStereo(
				constant(7, () => 1),
				layout
			)
		)

		expect(left).toBeLessThanOrEqual(1)
		expect(left).toBeGreaterThan(0.95)
		expect(right).toBeCloseTo(left ?? Number.NaN, 6)
	})

	it("plays mono at full level on both sides and passes stereo through", () => {
		expect(at(downmixToStereo([Float32Array.of(0.3)], ["FC"])).map(value => Math.round(value * 1000) / 1000)).toEqual([0.3, 0.3])

		const stereo: [Float32Array, Float32Array] = [Float32Array.of(0.5), Float32Array.of(-0.25)]

		expect(downmixToStereo(stereo, ["FL", "FR"])).toEqual(stereo)
	})
})
