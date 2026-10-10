import { AudioSample } from "mediabunny"
import { describe, expect, it, vi } from "vitest"
import { AUDIO_UNSUPPORTED_ERROR, SourceReadError, type RemuxStats } from "@/lib/media/remuxTypes"
import type { AudioInput } from "@/workers/remux/audioInput"
import type { SampleItem } from "@/workers/remux/audioTranscode"
import { pickCursor, PcmStream, RESTART_PREROLL_SECONDS } from "@/workers/remux/pcmStream"
import { WAV_HEADER_BYTES, wavHeader, wavLayout } from "@/workers/remux/wav.logic"

const RATE = 1000
const BLOCK = 100

function stats(): RemuxStats {
	return {
		blockReads: 0,
		bytesRead: 0,
		rangeReads: 0,
		readAheadBlocks: 0,
		segmentsOut: 0,
		bytesOut: 0,
		transcodeSeconds: 0,
		decodeWaitSeconds: 0,
		transcodedMediaSeconds: 0
	}
}

// A mono source of `seconds` whose sample at time t is (t × RATE) / 32767 (wrapping), decoded in blocks of
// 0.1 s from the block covering the time asked for, like a demuxer and decoder would: the 16-bit output at
// RATE reads back its frame index, so every byte names the time it came from. `declared` is the rate the
// container states, `rateOf` the rate each block decodes at.
function source(
	seconds: number,
	options: { declared?: number; rateOf?: (block: number) => number; declaredChannels?: number; decodedChannels?: number } = {}
) {
	const { declared = RATE, rateOf = () => RATE, declaredChannels = 1, decodedChannels = 1 } = options
	const made: AudioSample[] = []
	const firsts: number[] = []
	const close = vi.fn(() => Promise.resolve())
	const blocks = Math.ceil((seconds * RATE) / BLOCK)
	const block = (index: number): SampleItem | null => {
		if (index >= blocks) {
			return null
		}

		const rate = rateOf(index)
		const start = (index * BLOCK) / RATE
		const count = Math.round((Math.min(seconds, start + BLOCK / RATE) - start) * rate)
		const plane = Float32Array.from({ length: count }, (_, i) => (Math.round((start + i / rate) * RATE) % 32767) / 32767)
		const data = new Float32Array(count * decodedChannels)

		for (let channel = 0; channel < decodedChannels; channel++) {
			data.set(plane, channel * count)
		}

		const sample = new AudioSample({
			data,
			format: "f32-planar",
			numberOfChannels: decodedChannels,
			sampleRate: rate,
			timestamp: start
		})

		made.push(sample)

		return { timestamp: sample.timestamp, duration: sample.duration, type: "key", sample }
	}
	const input: AudioInput = {
		info: { codec: "aac", sampleRate: declared, channels: declaredChannels, duration: seconds, header: null },
		cache: null as never,
		reader: () => {
			let index = 0

			return {
				first: time => {
					firsts.push(time)
					index = Math.floor((time * RATE) / BLOCK)

					return Promise.resolve(block(index))
				},
				next: () => Promise.resolve(block(++index)),
				close: () => Promise.resolve()
			}
		},
		close
	}

	return { input, firsts, close, made }
}

async function open(seconds: number, options: Parameters<typeof source>[1] = {}) {
	const { input, firsts, close } = source(seconds, options)
	const stream = await PcmStream.open(input, stats())

	return { stream, firsts, layout: stream.layout, closed: close }
}

function frames(bytes: Uint8Array): number[] {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)

	return Array.from({ length: Math.floor(bytes.length / 2) }, (_, i) => view.getInt16(i * 2, true))
}

const never = new AbortController().signal

describe("PcmStream", () => {
	it("serves the header, then the frames each offset names", async () => {
		const { stream, layout } = await open(10)
		const head = await stream.read(0, WAV_HEADER_BYTES + 8, never)

		expect(head.subarray(0, WAV_HEADER_BYTES)).toEqual(wavHeader(layout))
		expect(frames(head.subarray(WAV_HEADER_BYTES))).toEqual([0, 1, 2, 3])
		expect(layout.size).toBe(WAV_HEADER_BYTES + 10 * RATE * 2)
	})

	it("continues the decode a read follows on from, and decodes through a short gap", async () => {
		const { stream, firsts } = await open(10)

		await stream.read(WAV_HEADER_BYTES, WAV_HEADER_BYTES + 200, never)
		// The next read starts where the last ended; one a second further on continues too.
		expect(frames(await stream.read(WAV_HEADER_BYTES + 200, WAV_HEADER_BYTES + 204, never))).toEqual([100, 101])
		expect(frames(await stream.read(WAV_HEADER_BYTES + 2 * 1100, WAV_HEADER_BYTES + 2 * 1102, never))).toEqual([1100, 1101])
		expect(firsts).toEqual([0])
	})

	it("restarts the decoder at the sample a far offset names, a little before it", async () => {
		const { stream, firsts } = await open(60)
		const at = WAV_HEADER_BYTES + 2 * 45_250

		expect(frames(await stream.read(at, at + 6, never))).toEqual([45_250, 45_251, 45_252].map(frame => frame % 32767))
		// The first decode is the one the stream opened with.
		expect(firsts).toEqual([0, 45.25 - RESTART_PREROLL_SECONDS])

		// Back to the start: the decode there is gone, so another restart.
		// Back near the start: the decode the stream opened with still stands there.
		expect(frames(await stream.read(WAV_HEADER_BYTES + 20, WAV_HEADER_BYTES + 24, never))).toEqual([10, 11])
		expect(firsts).toHaveLength(2)
	})

	it("keeps two decodes going for an element reading in two places", async () => {
		const { stream, firsts } = await open(60)
		const far = WAV_HEADER_BYTES + 2 * 50_000

		await stream.read(WAV_HEADER_BYTES, WAV_HEADER_BYTES + 100, never)
		await stream.read(far, far + 100, never)
		await stream.read(WAV_HEADER_BYTES + 100, WAV_HEADER_BYTES + 200, never)
		await stream.read(far + 100, far + 200, never)

		expect(firsts).toHaveLength(2)
	})

	it("pads the end the duration promised, and never reads past the stream", async () => {
		const { stream, layout } = await open(1)
		const tail = await stream.read(layout.size - 4, layout.size + 100, never)

		expect(frames(tail)).toEqual([998, 999])
	})

	it("answers reads one at a time, in order, and refuses one already aborted", async () => {
		const { stream } = await open(10)
		const aborted = new AbortController()

		aborted.abort()

		const [a, b] = await Promise.all([
			stream.read(WAV_HEADER_BYTES, WAV_HEADER_BYTES + 4, never),
			stream.read(WAV_HEADER_BYTES + 4, WAV_HEADER_BYTES + 8, never)
		])

		expect([...frames(a), ...frames(b)]).toEqual([0, 1, 2, 3])
		await expect(stream.read(WAV_HEADER_BYTES, WAV_HEADER_BYTES + 4, aborted.signal)).rejects.toThrow()
	})

	it("closes its input with it", async () => {
		const { stream, closed } = await open(10)

		await stream.close()

		expect(closed).toHaveBeenCalledOnce()
		await expect(stream.read(0, 4, never)).rejects.toThrow()
	})
})

describe("pickCursor", () => {
	it("takes the decode standing at or just before the offset, the nearest of them", () => {
		const a = { nextOffset: 100 }
		const b = { nextOffset: 180 }

		expect(pickCursor([a, b], 200, 50)).toBe(b)
		expect(pickCursor([a, b], 120, 50)).toBe(a)
		expect(pickCursor([a, b], 90, 50)).toBeNull()
		expect(pickCursor([a, b], 300, 50)).toBeNull()
	})
})

describe("PcmStream failures", () => {
	it("starts afresh after a decode failed, and keeps a decode a given-up read left", async () => {
		const { input, firsts } = source(10)
		const layout = wavLayout({ duration: 10, sampleRate: RATE, channels: 1 })

		if (layout === null) {
			throw new Error("no layout")
		}

		const failing = { once: true }
		const reader = input.reader.bind(input)
		const stream = new PcmStream({
			input: {
				...input,
				reader: (stats, isCancelled) => {
					const inner = reader(stats, isCancelled)

					return {
						...inner,
						next: (item: SampleItem) => {
							if (failing.once) {
								failing.once = false

								return Promise.reject(new Error("source lost"))
							}

							return inner.next(item)
						}
					}
				}
			},
			layout,
			stats: stats()
		})

		await expect(stream.read(WAV_HEADER_BYTES, WAV_HEADER_BYTES + 2 * 150, never)).rejects.toThrow("source lost")
		expect(frames(await stream.read(WAV_HEADER_BYTES + 2 * 150, WAV_HEADER_BYTES + 2 * 152, never))).toEqual([150, 151])
		expect(firsts).toHaveLength(2)

		const given = new AbortController()

		given.abort()
		await expect(stream.read(WAV_HEADER_BYTES + 2 * 152, WAV_HEADER_BYTES + 2 * 154, given.signal)).rejects.toThrow()
		expect(frames(await stream.read(WAV_HEADER_BYTES + 2 * 152, WAV_HEADER_BYTES + 2 * 154, never))).toEqual([152, 153])
		expect(firsts).toHaveLength(2)
	})
})

describe("PcmStream given-up reads", () => {
	// WebKit abandons a stream mid-read and asks again from where the bytes it got ended.
	it("continues the decode for a read again from where one given up mid-way started", async () => {
		const { input, firsts } = source(10)
		const given = new AbortController()
		const reader = input.reader.bind(input)
		const stream = await PcmStream.open(
			{
				...input,
				reader: (stats, isCancelled) => {
					const inner = reader(stats, isCancelled)

					return {
						...inner,
						next: (item: SampleItem) => {
							// The read is given up on while its second block decodes.
							given.abort()

							return inner.next(item)
						}
					}
				}
			},
			stats()
		)

		await expect(stream.read(WAV_HEADER_BYTES, WAV_HEADER_BYTES + 2 * 250, given.signal)).rejects.toThrow()
		expect(frames(await stream.read(WAV_HEADER_BYTES, WAV_HEADER_BYTES + 2 * 250, never))).toEqual(
			Array.from({ length: 250 }, (_, frame) => frame)
		)
		expect(firsts).toEqual([0])
	})
})

describe("PcmStream rates", () => {
	it("lays the stream out at the rate the decoder gives, not the one the container states", async () => {
		// HE-AAC: the container states the core rate, the decoder gives twice that.
		const { stream, layout } = await open(10, { declared: RATE / 2 })
		const at = WAV_HEADER_BYTES + 2 * 7_000

		expect(stream.info.sampleRate).toBe(RATE)
		expect(layout.frames).toBe(10 * RATE)
		expect(frames(await stream.read(at, at + 4, never))).toEqual([7_000, 7_001])
	})

	it("resamples blocks decoded at another rate mid-stream, keeping their place", async () => {
		const { stream } = await open(10, { rateOf: block => (block >= 30 ? RATE / 2 : RATE) })
		const at = WAV_HEADER_BYTES + 2 * 5_000
		const got = frames(await stream.read(at, at + 8, never))

		got.forEach((value, i) => {
			expect(Math.abs(value - (5_000 + i))).toBeLessThanOrEqual(1)
		})
		expect(frames(await stream.read(WAV_HEADER_BYTES + 2 * 2_000, WAV_HEADER_BYTES + 2 * 2_002, never))).toEqual([2_000, 2_001])
	})

	it("fails a file whose audio cannot be decoded as unsupported, and a read of the file as itself", async () => {
		const { input } = source(10)
		const broken = (error: Error): AudioInput => ({
			...input,
			reader: () => ({ first: () => Promise.reject(error), next: () => Promise.resolve(null), close: () => Promise.resolve() })
		})

		await expect(PcmStream.open(broken(new Error("unsupported config")), stats())).rejects.toMatchObject({
			name: AUDIO_UNSUPPORTED_ERROR
		})
		await expect(PcmStream.open(broken(new SourceReadError("offline")), stats())).rejects.toMatchObject({ name: "SourceReadError" })
	})
})

describe("PcmStream channels and memory", () => {
	it("keeps the container's channels over a narrower first block, and takes a wider one", async () => {
		expect((await open(2, { declaredChannels: 2 })).layout.channels).toBe(2)
		// Parametric stereo: the container says mono, the decoder gives two.
		expect((await open(2, { declaredChannels: 1, decodedChannels: 2 })).layout.channels).toBe(2)
		expect((await open(2)).layout.channels).toBe(1)
	})

	it("closes the block it opened with when its decode goes unread: closed, or crowded out", async () => {
		const closing = source(60)
		const closed = await PcmStream.open(closing.input, stats())
		const first = closing.made[0]

		if (first === undefined) {
			throw new Error("nothing decoded")
		}

		const spy = vi.spyOn(first, "close")

		await closed.close()
		expect(spy).toHaveBeenCalled()

		const crowded = source(60)
		const stream = await PcmStream.open(crowded.input, stats())
		const opener = crowded.made[0]

		if (opener === undefined) {
			throw new Error("nothing decoded")
		}

		const evicted = vi.spyOn(opener, "close")

		await stream.read(WAV_HEADER_BYTES + 2 * 20_000, WAV_HEADER_BYTES + 2 * 20_002, never)
		await stream.read(WAV_HEADER_BYTES + 2 * 40_000, WAV_HEADER_BYTES + 2 * 40_002, never)
		expect(evicted).toHaveBeenCalled()
	})
})
