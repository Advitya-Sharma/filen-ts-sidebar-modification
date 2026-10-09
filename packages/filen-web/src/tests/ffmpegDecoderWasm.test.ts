import { readFileSync } from "node:fs"
import { runInThisContext } from "node:vm"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import WORKER_SOURCE from "@/vendor/ffmpeg-decoders/decoders.worker.js?raw"
import type { DecodedFrame } from "@/workers/remux/ffmpegDecoder"

// The shipped module itself, run in Node behind a stand-in for its worker scope: the bridge and the worker
// loop exercised on real AC-3, E-AC-3 and DTS frames (synthetic tones, three 1536-sample AC-3 frames, three
// E-AC-3, two 512-sample DTS, all 5.1).

interface Answer {
	id: number
	handle?: number
	frames?: DecodedFrame[]
	error?: string
	fatal?: string
}

const AC3_FRAME = 768
const EAC3_FRAME = 384
const DTS_FRAME = 1024
// AV_CH_LAYOUT_5POINT1: FL FR FC LFE SL SR.
const MASK_5_1 = 0x60f

function fixture(name: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(`./fixtures/ffmpeg-decoders/${name}`, import.meta.url)))
}

function frames(bytes: Uint8Array, size: number): Uint8Array[] {
	return Array.from({ length: bytes.length / size }, (_, index) => bytes.slice(index * size, (index + 1) * size))
}

function startWorker() {
	const answers = new Map<number, (answer: Answer) => void>()
	let listener: ((event: { data: unknown }) => void) | null = null
	let nextId = 0
	const scope = {
		location: { href: "file:///decoders.worker.js" },
		closed: false,
		addEventListener: (_type: string, handler: (event: { data: unknown }) => void) => {
			listener = handler
		},
		postMessage: (answer: Answer) => {
			answers.get(answer.id)?.(answer)
		},
		close: () => {
			scope.closed = true
		}
	}
	const evaluate = runInThisContext(`(function (self, setInterval) {\n${WORKER_SOURCE}\n})`) as (
		self: unknown,
		setInterval: () => void
	) => void

	// The worker's keepalive timer would hold the test process open.
	evaluate(scope, () => undefined)

	const request = (message: Record<string, unknown>): Promise<Answer> => {
		const id = nextId++

		return new Promise(resolve => {
			answers.set(id, resolve)
			listener?.({ data: { id, ...message } })
		})
	}
	const decode = async (handle: number | undefined, bytes: Uint8Array, timestampUs = 0): Promise<Answer> =>
		request({ type: "decode", handle, data: bytes.slice().buffer, timestampUs })

	return { scope, request, decode }
}

let logged: unknown[][] = []

beforeEach(() => {
	logged = []

	for (const method of ["log", "warn", "error"] as const) {
		vi.spyOn(console, method).mockImplementation((...args: unknown[]) => {
			logged.push(args)
		})
	}
})

afterEach(() => {
	// FFmpeg's log stays quiet, damaged input included.
	expect(logged).toEqual([])
})

describe("FFmpeg decoders module", () => {
	it.each([
		["ac3", "ac3-51.ac3", AC3_FRAME, 1536],
		["eac3", "eac3-51.eac3", EAC3_FRAME, 1536],
		["dts", "dts-51.dts", DTS_FRAME, 512]
	] as const)("decodes %s 5.1 to planar float with its channel mask and timestamps", async (codec, file, size, perFrame) => {
		const worker = startWorker()
		const { handle } = await worker.request({ type: "open", codec })
		const packets = frames(fixture(file), size)

		for (const [index, packet] of packets.entries()) {
			// Whole microseconds travel through FFmpeg unchanged.
			const timestampUs = Math.round((index * perFrame * 1_000_000) / 48_000)
			const answer = await worker.decode(handle, packet, timestampUs)

			expect(answer.error).toBeUndefined()
			expect(answer.frames?.map(frame => ({ ...frame, data: frame.data.byteLength }))).toEqual([
				{
					data: perFrame * 6 * 4,
					format: "f32-planar",
					channels: 6,
					channelMask: MASK_5_1,
					sampleRate: 48_000,
					frames: perFrame,
					timestampUs
				}
			])
		}

		expect(await worker.request({ type: "flush", handle })).toMatchObject({ frames: [] })
		expect(await worker.request({ type: "close", handle })).toEqual({ id: packets.length + 2 })
	})

	it("skips packets the decoders reject, whatever the error, and decodes on", async () => {
		const worker = startWorker()
		const ac3 = (await worker.request({ type: "open", codec: "ac3" })).handle
		const [first, second] = frames(fixture("ac3-51.ac3"), AC3_FRAME)

		if (first === undefined || second === undefined) {
			throw new Error("fixture too short")
		}

		// No syncword at all, and a header with an impossible bitstream id (AC-3's own parse error code, which
		// once ended the whole audio track).
		const badBsid = first.slice()

		badBsid[5] = 0xf8

		for (const damaged of [new Uint8Array(AC3_FRAME), badBsid, first.slice(0, 40)]) {
			expect(await worker.decode(ac3, damaged)).toEqual({ id: expect.any(Number) as number, frames: [] })
		}

		expect((await worker.decode(ac3, second)).frames).toHaveLength(1)

		const eac3 = (await worker.request({ type: "open", codec: "eac3" })).handle
		const dts = (await worker.request({ type: "open", codec: "dts" })).handle
		const [eac3Frame] = frames(fixture("eac3-51.eac3"), EAC3_FRAME)
		const [dtsFrame] = frames(fixture("dts-51.dts"), DTS_FRAME)

		expect((await worker.decode(eac3, eac3Frame?.slice(0, 100) ?? new Uint8Array())).frames).toEqual([])
		expect((await worker.decode(eac3, eac3Frame ?? new Uint8Array())).frames).toHaveLength(1)
		expect((await worker.decode(dts, new Uint8Array(DTS_FRAME).fill(0x55))).frames).toEqual([])
		expect((await worker.decode(dts, dtsFrame ?? new Uint8Array())).frames).toHaveLength(1)
	})

	it("refuses an unknown codec without opening a decoder", async () => {
		const worker = startWorker()

		expect(await worker.request({ type: "open", codec: "aac" })).toEqual({ id: 0, error: "no FFmpeg decoder for aac" })
		expect(await worker.request({ type: "open", codec: "constructor" })).toEqual({ id: 1, error: "no FFmpeg decoder for constructor" })
		expect(worker.scope.closed).toBe(false)
	})

	it("answers a trap in the module as fatal and closes the worker, ignoring what comes after", async () => {
		const worker = startWorker()

		await worker.request({ type: "open", codec: "ac3" })

		const answer = await worker.decode(0x7ffffff0, new Uint8Array(AC3_FRAME))

		expect(answer.fatal).toEqual(expect.any(String))
		expect(answer.error).toBeUndefined()
		expect(worker.scope.closed).toBe(true)

		const late = await Promise.race([
			worker.request({ type: "open", codec: "ac3" }),
			new Promise(resolve => {
				setTimeout(() => {
					resolve("unanswered")
				}, 50)
			})
		])

		expect(late).toBe("unanswered")
	})
})
