import { describe, expect, it } from "vitest"
import { frameAtOffset, offsetOfFrame, WAV_HEADER_BYTES, wavHeader, wavLayout } from "@/workers/remux/wav.logic"
import { outputPlanes, PcmAssembler } from "@/workers/remux/pcmAssembler.logic"
import { audioRoute, AudioDemotions } from "@/lib/media/audioRoute.logic"

const HOUR = 3600

function ascii(bytes: Uint8Array, at: number, length: number): string {
	return String.fromCharCode(...bytes.subarray(at, at + length))
}

describe("wavLayout", () => {
	it("keeps the source's rate and at most two channels, sized exactly", () => {
		expect(wavLayout({ duration: 60, sampleRate: 44_100, channels: 6 })).toEqual({
			sampleRate: 44_100,
			channels: 2,
			decimation: 1,
			frames: 2_646_000,
			blockAlign: 4,
			size: WAV_HEADER_BYTES + 2_646_000 * 4
		})
		expect(wavLayout({ duration: 10, sampleRate: 48_000, channels: 1 })?.channels).toBe(1)
	})

	it("goes mono, then halves the rate, for audio too long for a 4 GiB WAV file", () => {
		const seven = wavLayout({ duration: 7 * HOUR, sampleRate: 48_000, channels: 2 })

		expect(seven).toMatchObject({ channels: 1, sampleRate: 48_000, decimation: 1, size: WAV_HEADER_BYTES + 7 * HOUR * 48_000 * 2 })

		expect(wavLayout({ duration: 20 * HOUR, sampleRate: 48_000, channels: 2 })).toMatchObject({
			channels: 1,
			sampleRate: 24_000,
			decimation: 2
		})

		const thirty = wavLayout({ duration: 30 * HOUR, sampleRate: 48_000, channels: 2 })

		expect(thirty).toMatchObject({ channels: 1, sampleRate: 12_000, decimation: 4 })
		expect(thirty?.size).toBeLessThanOrEqual(0xffff_ffff + 8)
	})

	it("has no layout for audio no WAV file can hold, nor for an unknown duration", () => {
		expect(wavLayout({ duration: 400 * HOUR, sampleRate: 48_000, channels: 2 })).toBeNull()
		expect(wavLayout({ duration: 0, sampleRate: 48_000, channels: 2 })).toBeNull()
		expect(wavLayout({ duration: Number.NaN, sampleRate: 48_000, channels: 2 })).toBeNull()
	})

	it("writes a PCM header whose sizes and rates match the layout", () => {
		const layout = wavLayout({ duration: 2, sampleRate: 44_100, channels: 2 })

		if (layout === null) {
			throw new Error("no layout")
		}

		const header = wavHeader(layout)
		const view = new DataView(header.buffer)

		expect([ascii(header, 0, 4), ascii(header, 8, 4), ascii(header, 12, 4), ascii(header, 36, 4)]).toEqual([
			"RIFF",
			"WAVE",
			"fmt ",
			"data"
		])
		expect(view.getUint32(4, true)).toBe(layout.size - 8)
		expect(view.getUint16(20, true)).toBe(1)
		expect(view.getUint16(22, true)).toBe(2)
		expect(view.getUint32(24, true)).toBe(44_100)
		expect(view.getUint32(28, true)).toBe(44_100 * 4)
		expect(view.getUint16(32, true)).toBe(4)
		expect(view.getUint16(34, true)).toBe(16)
		expect(view.getUint32(40, true)).toBe(layout.size - WAV_HEADER_BYTES)
	})

	it("maps a byte offset to the output frame it holds, and back", () => {
		const layout = wavLayout({ duration: HOUR, sampleRate: 48_000, channels: 2 })

		if (layout === null) {
			throw new Error("no layout")
		}

		// Seconds × byte rate past the header is that second's first frame.
		expect(frameAtOffset(layout, WAV_HEADER_BYTES + 1800 * 48_000 * 4)).toEqual({ frame: 1800 * 48_000, within: 0 })
		expect(frameAtOffset(layout, WAV_HEADER_BYTES + 10 * 4 + 3)).toEqual({ frame: 10, within: 3 })
		expect(frameAtOffset(layout, 0)).toEqual({ frame: 0, within: 0 })
		expect(offsetOfFrame(layout, 48_000)).toBe(WAV_HEADER_BYTES + 48_000 * 4)
	})
})

// One plane whose sample at source frame n is n / 32767, so a 16-bit output sample reads back the frame index.
function ramp(first: number, count: number): Float32Array {
	return Float32Array.from({ length: count }, (_, i) => (first + i) / 32767)
}

function samples(bytes: Uint8Array): number[] {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)

	return Array.from({ length: bytes.length / 2 }, (_, i) => view.getInt16(i * 2, true))
}

function monoLayout(frames: number, decimation = 1) {
	return { sampleRate: 1000 / decimation, channels: 1, decimation, frames, blockAlign: 2, size: WAV_HEADER_BYTES + frames * 2 }
}

describe("PcmAssembler", () => {
	it("starts exactly at its frame, whatever the first block's start", () => {
		const assembler = new PcmAssembler(monoLayout(100), 10)

		assembler.push([ramp(4, 10)], 4)

		expect(assembler.nextOffset).toBe(WAV_HEADER_BYTES + 10 * 2)
		expect(samples(assembler.take(100))).toEqual([10, 11, 12, 13])
	})

	it("fills a gap with silence and drops an overlap, so the output never drifts", () => {
		const assembler = new PcmAssembler(monoLayout(100), 0)

		assembler.push([ramp(0, 4)], 0)
		// A gap of 3 frames, then an overlap of 4.
		assembler.push([ramp(7, 3)], 7)
		assembler.push([ramp(6, 6)], 6)

		expect(samples(assembler.take(1000))).toEqual([0, 1, 2, 3, 0, 0, 0, 7, 8, 9, 10, 11])
	})

	it("queues a long gap as a count, making its silence only as it is taken", () => {
		const frames = 6 * HOUR * 48_000
		const assembler = new PcmAssembler(
			{ sampleRate: 48_000, channels: 1, decimation: 1, frames, blockAlign: 2, size: WAV_HEADER_BYTES + frames * 2 },
			0
		)

		assembler.push([ramp(0, 4)], 0)
		// A timestamp jump to the end of a six-hour declared duration: gigabytes of silence, none of it made yet.
		assembler.push([ramp(1, 2)], frames - 2)

		expect(assembler.queued).toBe(frames * 2)
		expect(samples(assembler.take(12))).toEqual([0, 1, 2, 3, 0, 0])

		assembler.skip(frames * 2 - 12 - 4)

		expect(samples(assembler.take(100))).toEqual([1, 2])
		expect(assembler.complete).toBe(true)
	})

	it("completes a decimation group with silence before counting the rest", () => {
		const assembler = new PcmAssembler(monoLayout(10, 2), 0)

		assembler.push([Float32Array.from([2, 4, 6].map(v => v / 32767))], 0)
		assembler.silence(4)
		assembler.push([Float32Array.from([8, 10].map(v => v / 32767))], 7)

		// The 10 waits in a group of its own.
		expect(samples(assembler.take(1000))).toEqual([3, 3, 0, 4])
	})

	it("takes a frame or two of rounding drift as continuous", () => {
		const assembler = new PcmAssembler({ ...monoLayout(1000), sampleRate: 48_000 }, 0)

		assembler.push([ramp(0, 4)], 0)
		assembler.push([ramp(4, 4)], 5)

		expect(samples(assembler.take(1000))).toEqual([0, 1, 2, 3, 4, 5, 6, 7])
	})

	it("averages whole groups for a decimated output, carrying a partial group to the next block", () => {
		const assembler = new PcmAssembler(monoLayout(10, 2), 0)

		assembler.push([Float32Array.from([2, 4, 6].map(v => v / 32767))], 0)
		assembler.push([Float32Array.from([8, 10, 12].map(v => v / 32767))], 3)

		expect(samples(assembler.take(1000))).toEqual([3, 7, 11])
	})

	it("stops at the declared length and pads up to it when the decoder ends short", () => {
		const assembler = new PcmAssembler(monoLayout(6), 0)

		assembler.push([ramp(0, 4)], 0)
		assembler.finish()

		expect(samples(assembler.take(1000))).toEqual([0, 1, 2, 3, 0, 0])
		expect(assembler.complete).toBe(true)

		assembler.push([ramp(6, 4)], 6)
		expect(assembler.queued).toBe(0)
	})

	it("hands bytes out across a frame boundary without losing its place", () => {
		const assembler = new PcmAssembler({ ...monoLayout(4), channels: 2, blockAlign: 4, size: WAV_HEADER_BYTES + 16 }, 0)

		assembler.push([ramp(0, 4), ramp(100, 4)], 0)

		expect(assembler.take(3).length).toBe(3)
		expect(assembler.nextOffset).toBe(WAV_HEADER_BYTES + 3)
		expect(assembler.take(100).length).toBe(13)
		expect(assembler.complete).toBe(true)
	})

	it("clamps samples beyond full scale", () => {
		const assembler = new PcmAssembler(monoLayout(2), 0)

		assembler.push([Float32Array.from([2, -2])], 0)

		expect(samples(assembler.take(4))).toEqual([32767, -32767])
	})
})

describe("outputPlanes", () => {
	it("doubles mono for a stereo output and folds a stereo pair for a mono one", () => {
		const mono = Float32Array.from([0.5, -0.5])

		expect(outputPlanes([mono], 2, ["FC"])).toEqual([mono, mono])
		expect(outputPlanes([Float32Array.from([1, 0]), Float32Array.from([0, 1])], 1, ["FL", "FR"])).toEqual([
			Float32Array.from([0.5, 0.5])
		])
	})

	it("downmixes surround to stereo by its layout", () => {
		const one = Float32Array.from([1])
		const zero = Float32Array.from([0])
		const [left, right] = outputPlanes([zero, zero, one, zero, zero, zero], 2, ["FL", "FR", "FC", "LFE", "BL", "BR"])

		expect(left?.[0]).toBeCloseTo(Math.SQRT1_2)
		expect(right?.[0]).toBeCloseTo(Math.SQRT1_2)
	})
})

describe("audioRoute", () => {
	it("plays as it is what the browser may, and decodes the rest", () => {
		expect(audioRoute({ extension: "mp3", contentType: "audio/mpeg", demoted: false })).toBe("direct")
		expect(audioRoute({ extension: "m4b", contentType: "audio/mp4", demoted: false })).toBe("direct")

		for (const extension of ["mka", "ac3", "eac3", "ec3", "dts"]) {
			expect(audioRoute({ extension, contentType: "audio/x-matroska", demoted: false })).toBe("transcode")
		}

		expect(audioRoute({ extension: "m4a", contentType: "audio/mp4", demoted: true })).toBe("transcode")
		expect(audioRoute({ extension: "mp3", contentType: null, demoted: false })).toBe("transcode")
	})

	it("remembers a bounded set of demoted files, oldest out first", () => {
		const demotions = new AudioDemotions(2)

		demotions.add("a")
		demotions.add("b")
		demotions.add("a")
		demotions.add("c")

		expect([demotions.has("a"), demotions.has("b"), demotions.has("c")]).toEqual([true, false, true])
	})
})
