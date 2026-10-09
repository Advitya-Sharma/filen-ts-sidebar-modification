import { describe, expect, it } from "vitest"
import { estimateFrameStreamDuration, id3v2Bytes } from "@/workers/remux/audioDuration.logic"

// MPEG-1 Layer III, 128 kbps, 44.1 kHz, stereo: 417 bytes and 1152 samples a frame.
const MP3_FRAME = 417
const MP3_SECONDS = 1152 / 44100

function mp3Frames(count: number, fill = 0): number[] {
	const bytes: number[] = []

	for (let i = 0; i < count; i++) {
		bytes.push(0xff, 0xfb, 0x90, 0x00, ...new Array<number>(MP3_FRAME - 4).fill(fill))
	}

	return bytes
}

// The same frame carrying a Xing tag after its 32 bytes of side information.
function xingFrame(): number[] {
	const frame = mp3Frames(1)

	frame.splice(36, 4, ...new TextEncoder().encode("Xing"))

	return frame
}

// AAC LC, 48 kHz, stereo: 1024 samples a frame, whatever its length.
function adtsFrame(length: number): number[] {
	return [
		0xff,
		0xf1,
		0x4c,
		0x80 | ((length >> 11) & 0x03),
		(length >> 3) & 0xff,
		((length & 0x07) << 5) | 0x1f,
		0xfc,
		...new Array<number>(length - 7).fill(0x21)
	]
}

function id3(size: number, footer = false): number[] {
	return [0x49, 0x44, 0x33, 4, 0, footer ? 0x10 : 0, (size >> 21) & 0x7f, (size >> 14) & 0x7f, (size >> 7) & 0x7f, size & 0x7f]
}

function estimate(stream: "mp3" | "adts", bytes: number[], size = bytes.length): number | null {
	return estimateFrameStreamDuration({ stream, bytes: Uint8Array.from(bytes), size })
}

describe("id3v2Bytes", () => {
	it("measures a tag from its syncsafe size, footer included", () => {
		expect(id3v2Bytes(Uint8Array.from(id3(300)))).toBe(310)
		expect(id3v2Bytes(Uint8Array.from(id3(300, true)))).toBe(320)
		expect(id3v2Bytes(Uint8Array.from(id3(2 ** 21 + 5)))).toBe(2 ** 21 + 15)
	})

	it("finds none where no tag starts, or its size is not syncsafe", () => {
		expect(id3v2Bytes(Uint8Array.from(mp3Frames(1).slice(0, 10)))).toBe(0)
		expect(id3v2Bytes(Uint8Array.from([0x49, 0x44, 0x33, 4, 0, 0, 0, 0, 0x80, 0]))).toBe(0)
		expect(id3v2Bytes(Uint8Array.from([0x49, 0x44, 0x33]))).toBe(0)
	})
})

describe("estimateFrameStreamDuration", () => {
	it("times a constant-bitrate MP3 by its size over the first frames' byte rate", () => {
		const head = mp3Frames(100)
		const size = 10_000 * MP3_FRAME

		expect(estimate("mp3", head, size)).toBeCloseTo(10_000 * MP3_SECONDS, 6)
	})

	it("counts every frame of a stream that ends inside what it was given", () => {
		expect(estimate("mp3", mp3Frames(5))).toBeCloseTo(5 * MP3_SECONDS, 9)
	})

	it("starts at the first frame another frame follows, past junk and a stray sync word", () => {
		const junk = [0x00, 0xff, 0xfb, 0x90, 0x00, 0x12, 0x34]
		const head = [...junk, ...mp3Frames(20)]

		expect(estimate("mp3", head, junk.length + 1000 * MP3_FRAME)).toBeCloseTo(1000 * MP3_SECONDS, 6)
	})

	it("skips a Xing frame, which holds no audio", () => {
		const head = [...xingFrame(), ...mp3Frames(20)]

		expect(estimate("mp3", head, MP3_FRAME + 1000 * MP3_FRAME)).toBeCloseTo(1000 * MP3_SECONDS, 6)
	})

	it("averages ADTS frames of varying length", () => {
		const lengths = [300, 400, 350, 450, 380, 420, 360, 440, 390, 410]
		const head = lengths.flatMap(length => adtsFrame(length))
		const rate = head.length / ((lengths.length * 1024) / 48000)

		expect(estimate("adts", head, 1_000_000)).toBeCloseTo(1_000_000 / rate, 6)
	})

	it("tells nothing from too few frames, or none", () => {
		expect(estimate("mp3", mp3Frames(3), 100 * MP3_FRAME)).toBeNull()
		expect(estimate("adts", new Array<number>(4096).fill(0x55), 1_000_000)).toBeNull()
		expect(estimate("mp3", adtsFrame(400).concat(adtsFrame(400)), 1_000_000)).toBeNull()
		expect(estimate("mp3", [], 1000)).toBeNull()
	})
})
