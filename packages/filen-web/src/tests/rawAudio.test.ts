import { describe, expect, it } from "vitest"
import { findRawSync, measureRawStream, nextRawUnit, parseRawFrame, rawUnitAt } from "@/workers/remux/rawAudio.logic"
import { alacExtradata, rawPackets, rawStreamOf } from "@/workers/remux/audioInput"
import { BlockCache } from "@/lib/media/blockCache"

// Frame headers written field by field, the rest of each frame zeros: enough for the parser, nothing a
// decoder would accept.
function bitsToBytes(fields: [number, number][], size: number): Uint8Array {
	const out = new Uint8Array(size)
	let bit = 0

	for (const [value, count] of fields) {
		for (let i = count - 1; i >= 0; i--) {
			if (((value >> i) & 1) === 1) {
				out[bit >> 3] = (out[bit >> 3] ?? 0) | (0x80 >> (bit & 7))
			}

			bit++
		}
	}

	return out
}

// 48 kHz, 192 kb/s (frmsizecod 20: 384 words), bsid 8, 3/2 with LFE. At 44.1 kHz (fscod 1) the same rate is
// 417 words, or 418 with frmsizecod's padding bit set.
function ac3Frame(fscod = 0, frmsizecod = 20): Uint8Array {
	const words = fscod === 1 ? 417 + (frmsizecod & 1) : 384

	return bitsToBytes(
		[
			[0x0b77, 16],
			[0, 16],
			[fscod, 2],
			[frmsizecod, 6],
			[8, 5],
			[0, 3],
			[7, 3],
			[0, 2],
			[0, 2],
			[1, 1]
		],
		words * 2
	)
}

// An E-AC-3 frame: `type` 0 independent, 1 dependent; 48 kHz, six blocks, 3/2 with LFE.
function eac3Frame(type: number, bytes: number): Uint8Array {
	return bitsToBytes(
		[
			[0x0b77, 16],
			[type, 2],
			[0, 3],
			[bytes / 2 - 1, 11],
			[0, 2],
			[3, 2],
			[7, 3],
			[1, 1],
			[16, 5]
		],
		bytes
	)
}

// A DTS core frame: 16 blocks of 32 samples at 48 kHz, 3/2 (AMODE 9) with LFE.
function dtsFrame(bytes: number): Uint8Array {
	return bitsToBytes(
		[
			[0x7ffe, 16],
			[0x8001, 16],
			[1, 1],
			[31, 5],
			[0, 1],
			[15, 7],
			[bytes - 1, 14],
			[9, 6],
			[13, 4],
			[15, 5],
			[0, 10],
			[1, 2]
		],
		bytes
	)
}

function concat(...parts: Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
	let at = 0

	for (const part of parts) {
		out.set(part, at)
		at += part.length
	}

	return out
}

describe("parseRawFrame", () => {
	it("reads an AC-3 frame's size, rate and channels", () => {
		expect(parseRawFrame(ac3Frame(), 0)).toEqual({
			codec: "ac3",
			bytes: 768,
			samples: 1536,
			sampleRate: 48_000,
			channels: 6,
			unitStart: true
		})
	})

	it("tells an E-AC-3 dependent frame from the independent one it extends", () => {
		expect(parseRawFrame(eac3Frame(0, 512), 0)).toMatchObject({
			codec: "eac3",
			bytes: 512,
			samples: 1536,
			sampleRate: 48_000,
			channels: 6,
			unitStart: true
		})
		expect(parseRawFrame(eac3Frame(1, 256), 0)?.unitStart).toBe(false)
	})

	it("reads a DTS core frame", () => {
		expect(parseRawFrame(dtsFrame(2048), 0)).toEqual({
			codec: "dts",
			bytes: 2048,
			samples: 512,
			sampleRate: 48_000,
			channels: 6,
			unitStart: true
		})
	})

	it("finds nothing in other bytes, nor in a header cut short", () => {
		expect(parseRawFrame(new Uint8Array(64), 0)).toBeNull()
		expect(parseRawFrame(ac3Frame().subarray(0, 8), 0)).toBeNull()
	})
})

describe("measureRawStream", () => {
	it("measures a stream behind leading junk, confirmed by the frame that follows", () => {
		const frame = ac3Frame()
		const junk = new Uint8Array(100)
		const info = measureRawStream(concat(junk, frame, frame, frame), 100 + 768 * 1000)

		expect(info).toEqual({
			codec: "ac3",
			start: 100,
			unitBytes: 768,
			unitSamples: 1536,
			sampleRate: 48_000,
			channels: 6,
			units: 1000,
			duration: 32,
			variable: false,
			timeline: null
		})
	})

	it("counts an E-AC-3 unit's dependent frames into it", () => {
		const unit = concat(eac3Frame(0, 512), eac3Frame(1, 256))
		const info = measureRawStream(concat(unit, unit, unit), 768 * 10)

		expect(info).toMatchObject({ codec: "eac3", unitBytes: 768, units: 10 })
	})

	it("counts a DTS-HD extension between core frames into the unit", () => {
		const unit = concat(dtsFrame(2048), Uint8Array.from([0x64, 0x58, 0x20, 0x25]), new Uint8Array(508))
		const info = measureRawStream(concat(unit, unit, unit), 2560 * 4)

		expect(info).toMatchObject({ codec: "dts", unitBytes: 2560, units: 4, duration: (4 * 512) / 48_000 })
	})

	it("is no stream for a lone sync word", () => {
		expect(measureRawStream(concat(ac3Frame(), new Uint8Array(2000)), 10_000)).toBeNull()
		expect(measureRawStream(new Uint8Array(4096), 4096)).toBeNull()
	})
})

describe("rawUnitAt", () => {
	it("maps a time to its unit's offset and timestamp, clamped to the stream", () => {
		const frame = ac3Frame()
		const info = measureRawStream(concat(frame, frame), 768 * 1000)

		if (info === null) {
			throw new Error("no stream")
		}

		expect(rawUnitAt(info, 10)).toEqual({ index: 312, offset: 312 * 768 })
		expect(rawUnitAt(info, 1000).index).toBe(999)
		expect(rawUnitAt(info, -1).index).toBe(0)
	})

	it("finds the next sync of a codec past damaged bytes", () => {
		const bytes = concat(new Uint8Array(10), ac3Frame())

		expect(findRawSync(bytes, 0, "ac3")).toBe(10)
		expect(findRawSync(bytes, 11, "ac3")).toBe(-1)
	})
})

describe("alacExtradata", () => {
	it("gives FFmpeg the 36-byte atom from Matroska's bare config or MP4's box payload", () => {
		const config = Uint8Array.from({ length: 24 }, (_, i) => i + 1)
		const atom = alacExtradata(config)

		expect(atom?.length).toBe(36)
		expect(Array.from(atom?.subarray(0, 12) ?? [])).toEqual([0, 0, 0, 36, 0x61, 0x6c, 0x61, 0x63, 0, 0, 0, 0])
		expect(atom?.subarray(12)).toEqual(config)
		expect(alacExtradata(concat(new Uint8Array(4), config))).toEqual(atom)
		expect(alacExtradata(new Uint8Array(10))).toBeNull()
	})
})

// A 44.1 kHz AC-3 stream whose frames alternate between the plain and the padded size, as encoders write it.
function paddedAc3(frames: number): Uint8Array {
	return concat(...Array.from({ length: frames }, (_, i) => ac3Frame(1, 20 + (i % 2))))
}

function cacheOf(bytes: Uint8Array): BlockCache {
	return new BlockCache(bytes.length, (start, end) => Promise.resolve(bytes.slice(start, end)))
}

describe("variable unit sizes", () => {
	it("measures the average unit of a stream whose frames alternate in size", () => {
		const stream = paddedAc3(200)
		const info = measureRawStream(stream, stream.length)

		expect(info).toMatchObject({ codec: "ac3", sampleRate: 44_100, unitBytes: 835, units: 200 })
	})

	it("walks each unit's own size and lands a far start on the unit its offset names", async () => {
		const stream = paddedAc3(2000)
		const info = measureRawStream(stream.subarray(0, 262_144), stream.length)

		if (info === null) {
			throw new Error("no stream")
		}

		const packets = rawPackets(cacheOf(stream), info)
		const unitSeconds = 1536 / 44_100
		const far = await packets.first(1500 * unitSeconds + 0.001)

		expect(far?.offset).toBe(1500 * 835)
		expect(far?.index).toBe(1500)

		const sizes: number[] = []
		let packet = far

		for (let i = 0; i < 4 && packet !== null; i++) {
			sizes.push(packet.data.length)
			packet = await packets.next(packet)
		}

		expect(sizes).toEqual([834, 836, 834, 836])
		expect(packet === null ? null : packets.timestamp(packet)).toBeCloseTo(1504 * unitSeconds)
	})

	it("resyncs past a DTS-HD extension of varying size", async () => {
		const unit = (extension: number) => concat(dtsFrame(2048), Uint8Array.from([0x64, 0x58, 0x20, 0x25]), new Uint8Array(extension - 4))
		const stream = concat(unit(512), unit(800), unit(512), unit(1200), unit(512))
		const info = measureRawStream(stream, stream.length)

		if (info === null) {
			throw new Error("no stream")
		}

		const packets = rawPackets(cacheOf(stream), info)
		const sizes: number[] = []

		for (let packet = await packets.first(0); packet !== null; packet = await packets.next(packet)) {
			sizes.push(packet.data.length)
		}

		expect(sizes).toEqual([2560, 2848, 2560, 3248, 2560])
	})

	it("passes over an E-AC-3 dependent frame's sync to the next unit start", () => {
		const unit = concat(eac3Frame(0, 512), eac3Frame(1, 256))
		const bytes = concat(eac3Frame(1, 256), unit, unit, unit)

		expect(findRawSync(bytes, 0, "eac3")).toBe(0)
		expect(nextRawUnit(bytes, 0, "eac3", true)).toBe(256)
		expect(nextRawUnit(new Uint8Array(4096), 0, "eac3", true)).toBe(-1)
	})
})

describe("variable-bitrate timeline", () => {
	// DTS core frames each followed by an extension substream growing from 100 to ~2000 bytes: the bitrate rises
	// through the file, as a lossless extension's does with denser content.
	function rampingDts(units: number): Uint8Array {
		return concat(
			...Array.from({ length: units }, (_, i) =>
				concat(dtsFrame(2048), Uint8Array.from([0x64, 0x58, 0x20, 0x25]), new Uint8Array(96 + Math.floor((1900 * i) / units)))
			)
		)
	}

	async function opened(units: number) {
		const stream = rampingDts(units)
		const cache = cacheOf(stream)
		const input = { getFormat: () => Promise.reject(new Error("not a container")) }
		const raw = await rawStreamOf(cache, input as never)

		if (raw === null) {
			throw new Error("no stream")
		}

		return { stream, info: raw.info, packets: rawPackets(cache, raw.info) }
	}

	it("samples the timeline: a duration near the real one, and a tail that plays to the last unit", async () => {
		const { stream, info, packets } = await opened(1200)
		const unitSeconds = 512 / 48_000

		expect(info.variable).toBe(true)
		expect(Math.abs(info.units - 1200)).toBeLessThan(1200 * 0.02)
		expect(info.duration).toBeCloseTo(info.units * unitSeconds)

		const last = await packets.first(info.duration - unitSeconds / 2)

		expect(last === null ? null : last.offset + last.data.length).toBe(stream.length)
		expect(last?.index).toBe(info.units - 1)
		expect(last === null ? null : await packets.next(last)).toBeNull()
	})

	it("gives a unit the same bytes whether a seek reaches it or playback walks into it", async () => {
		const { info, packets } = await opened(1200)
		const unitSeconds = 512 / 48_000
		const walked = new Map<number, number>()

		for (let packet = await packets.first(0); packet !== null; packet = await packets.next(packet)) {
			expect(walked.has(packet.index)).toBe(false)
			walked.set(packet.index, packet.offset)
		}

		expect(walked.size).toBe(info.units)

		for (const index of [1, 467, 468, 469, 700, 937, info.units - 2]) {
			const sought = await packets.first(index * unitSeconds + unitSeconds / 4)

			expect([sought?.index, sought?.offset]).toEqual([index, walked.get(index)])
		}
	})
})
