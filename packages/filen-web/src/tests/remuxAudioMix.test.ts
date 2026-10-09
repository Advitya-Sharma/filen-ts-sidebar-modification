import { describe, expect, it } from "vitest"
import {
	audioSpecificConfig,
	FirstChunkConfig,
	LinearResampler,
	nextBlockTimestamp,
	pickEncoderTarget
} from "@/workers/remux/audioMix.logic"

describe("pickEncoderTarget", () => {
	it("takes the first target the page accepts that this engine can encode", async () => {
		const aacOnly = (codec: string) => Promise.resolve(codec === "mp4a.40.2")
		const opusOnly = (codec: string) => Promise.resolve(codec === "opus")

		expect(await pickEncoderTarget(["aac", "opus"], aacOnly)).toEqual({ target: "aac", codec: "mp4a.40.2" })
		expect(await pickEncoderTarget(["aac", "opus"], opusOnly)).toEqual({ target: "opus", codec: "opus" })
		// Opus the encoder could do, but the page's MediaSource does not take it in MP4 (Safari).
		expect(await pickEncoderTarget(["aac"], opusOnly)).toBeNull()
	})
})

describe("nextBlockTimestamp", () => {
	it("stays continuous across small drift and re-anchors on a jump in the source", () => {
		expect(nextBlockTimestamp({ expectedUs: null, sourceUs: 500, blockUs: 21_333 })).toEqual({ timestampUs: 500, reanchored: true })
		expect(nextBlockTimestamp({ expectedUs: 21_333, sourceUs: 21_340, blockUs: 21_333 })).toEqual({
			timestampUs: 21_333,
			reanchored: false
		})
		// Two seconds of the file missing: the output follows the source, not the running total.
		expect(nextBlockTimestamp({ expectedUs: 1_000_000, sourceUs: 3_000_000, blockUs: 21_333 })).toEqual({
			timestampUs: 3_000_000,
			reanchored: true
		})
	})
})

describe("FirstChunkConfig", () => {
	it("hands a dropped chunk's decoder config to the first chunk written", () => {
		const configs = new FirstChunkConfig<string>()

		expect(configs.take("config", true)).toBeUndefined()
		expect(configs.take(undefined, false)).toBe("config")
		expect(configs.take(undefined, false)).toBeUndefined()
	})

	it("passes a written chunk's own config through", () => {
		const configs = new FirstChunkConfig<string>()

		expect(configs.take("own", false)).toBe("own")
		expect(configs.take("later", false)).toBe("later")
	})
})

describe("LinearResampler", () => {
	it("keeps a ramp a ramp across block boundaries", () => {
		const resampler = new LinearResampler(32_000, 48_000)
		const out: number[] = []

		for (let block = 0; block < 4; block++) {
			const plane = Float32Array.from({ length: 320 }, (_, index) => block * 320 + index)
			const [resampled] = resampler.process([plane])

			out.push(...(resampled ?? []))
		}

		// 1280 input frames at 2/3 of the output rate: about 1920 output frames, evenly spaced.
		expect(out.length).toBeGreaterThanOrEqual(1917)
		expect(out.length).toBeLessThanOrEqual(1920)

		for (let index = 1; index < out.length; index++) {
			expect((out[index] ?? 0) - (out[index - 1] ?? 0)).toBeCloseTo(2 / 3, 3)
		}
	})
})

describe("audioSpecificConfig", () => {
	it("passes a bare AudioSpecificConfig through", () => {
		const asc = Uint8Array.of(0x11, 0x90)

		expect(audioSpecificConfig(asc)).toBe(asc)
	})

	it("digs the AudioSpecificConfig out of a whole ES_Descriptor", () => {
		// What WebKit's AAC encoder reports: ES_Descriptor > DecoderConfigDescriptor > DecoderSpecificInfo,
		// each with a four-byte size, followed by an SLConfigDescriptor.
		const esds = Uint8Array.of(
			3,
			128,
			128,
			128,
			34,
			0,
			0,
			0,
			4,
			128,
			128,
			128,
			20,
			64,
			20,
			0,
			24,
			0,
			0,
			0,
			0,
			0,
			0,
			0,
			0,
			0,
			5,
			128,
			128,
			128,
			2,
			17,
			144,
			6,
			128,
			128,
			128,
			1,
			2
		)

		expect(Array.from(audioSpecificConfig(esds))).toEqual([0x11, 0x90])
	})

	it("leaves a truncated descriptor as it is", () => {
		const broken = Uint8Array.of(3, 128, 128)

		expect(audioSpecificConfig(broken)).toBe(broken)
	})
})
