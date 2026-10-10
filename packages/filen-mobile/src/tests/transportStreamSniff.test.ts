import { vi, describe, it, expect, beforeEach } from "vitest"

const { mockFetch } = vi.hoisted(() => ({
	mockFetch: vi.fn()
}))

vi.mock("expo/fetch", () => ({
	fetch: mockFetch
}))

import { fs } from "@/tests/mocks/expoFileSystem"
import { SNIFF_HEAD_BYTES, readRangeFromUrl, sniffTransportStream, transportStreamLayout, transportStreamOutcome } from "@/lib/transportStreamSniff"

const MIB = 1024 * 1024

// `packets` sync bytes every `stride` bytes from `offset`, the first packet carrying `first` (a PAT) when given.
function stream(stride: number, offset: number, first: number[] | null = null, packets = 5): Uint8Array {
	const bytes = new Uint8Array(SNIFF_HEAD_BYTES)

	for (let k = 0; k < packets; k++) {
		const at = offset + k * stride

		if (at < bytes.length) {
			bytes[at] = 0x47
		}
	}

	if (first !== null) {
		bytes.set([0x47, 0x40, 0x00, 0x10, 0x00, ...first], offset)
	}

	return bytes
}

// A PAT naming `programs` (program 0, the network PID, names none).
function pat(programs: number[]): number[] {
	const entries = programs.flatMap(program => [program >> 8, program & 0xff, 0xe1, 0x00])
	const length = 5 + entries.length + 4

	return [0x00, 0xb0 | (length >> 8), length & 0xff, 0x00, 0x01, 0xc1, 0x00, 0x00, ...entries, 0, 0, 0, 0]
}

function reading(bytes: Uint8Array) {
	return vi.fn(async (_start: number, length: number) => bytes.subarray(0, length))
}

// Text whose every line is `line`, so a character in one column repeats every `line.length` bytes.
function text(line: string): Uint8Array {
	return new TextEncoder().encode(line.repeat(Math.ceil(SNIFF_HEAD_BYTES / line.length)))
}

describe("sniffTransportStream", () => {
	it("reads only the head, once, and reports the packet layout", async () => {
		const readRange = reading(stream(188, 0))

		expect(SNIFF_HEAD_BYTES).toBe(1020)
		expect(await sniffTransportStream(readRange, 10 * MIB)).toEqual({ stride: 188, prefix: 0 })
		expect(readRange).toHaveBeenCalledTimes(1)
		expect(readRange).toHaveBeenCalledWith(0, SNIFF_HEAD_BYTES)
		expect(await sniffTransportStream(reading(stream(192, 4)), MIB)).toEqual({ stride: 192, prefix: 4 })
		expect(await sniffTransportStream(reading(stream(204, 0)), MIB)).toEqual({ stride: 204, prefix: 0 })
	})

	it("is a stream whatever its PAT names, a multiplex of several programs included", async () => {
		expect(await sniffTransportStream(reading(stream(188, 0, pat([1, 2]))), MIB)).toEqual({ stride: 188, prefix: 0 })
		expect(await sniffTransportStream(reading(stream(188, 0, pat([0, 1]))), MIB)).toEqual({ stride: 188, prefix: 0 })
	})

	it("finds a stream cut mid-packet at whatever offset its first sync byte sits", async () => {
		expect(await sniffTransportStream(reading(stream(188, 1)), MIB)).toEqual({ stride: 188, prefix: 1 })
		expect(await sniffTransportStream(reading(stream(188, 187)), MIB)).toEqual({ stride: 188, prefix: 187 })
		expect(await sniffTransportStream(reading(stream(192, 100)), MIB)).toEqual({ stride: 192, prefix: 100 })
		expect(await sniffTransportStream(reading(stream(204, 203)), MIB)).toEqual({ stride: 204, prefix: 203 })
	})

	it("is no stream for TypeScript, for fixed-width lines that repeat a G every packet, or for too few packets", async () => {
		expect(await sniffTransportStream(reading(new TextEncoder().encode("export const x = 1\n".repeat(60))), MIB)).toBe(false)
		expect(await sniffTransportStream(reading(text(`${"G".padEnd(46, "x")}\n`)), MIB)).toBe(false)
		expect(await sniffTransportStream(reading(stream(188, 0, null, 4)), MIB)).toBe(false)
	})

	it("reads nothing of a file too small for a player to sync, and a file smaller than the head whole", async () => {
		const tiny = vi.fn(async (_start: number, length: number) => new Uint8Array(length))

		expect(await sniffTransportStream(tiny, 939)).toBe(false)
		expect(tiny).not.toHaveBeenCalled()

		const small = reading(stream(188, 0))

		expect(await sniffTransportStream(small, 950)).toEqual({ stride: 188, prefix: 0 })
		expect(small).toHaveBeenCalledWith(0, 950)
	})

	it("is null when the read fails", async () => {
		expect(
			await sniffTransportStream(async () => {
				throw new Error("network")
			}, MIB)
		).toBeNull()
	})
})

describe("transportStreamLayout", () => {
	it("prefers 188-byte packets and checks no stride its head is too short for", () => {
		expect(transportStreamLayout(stream(188, 0).subarray(0, 940))).toEqual({ stride: 188, prefix: 0 })
		expect(transportStreamLayout(stream(204, 0).subarray(0, 1019))).toBe(false)
	})
})

describe("transportStreamOutcome", () => {
	it("is notTransportStream for no stream, on either platform", () => {
		expect(transportStreamOutcome(false, "android")).toBe("notTransportStream")
		expect(transportStreamOutcome(false, "ios")).toBe("notTransportStream")
	})

	it("keeps only a 188-byte stream on Android, which Media3 reads from any offset; BDAV and parity streams are unplayable", () => {
		expect(transportStreamOutcome({ stride: 188, prefix: 0 }, "android")).toBe("transportStream")
		expect(transportStreamOutcome({ stride: 188, prefix: 37 }, "android")).toBe("transportStream")
		expect(transportStreamOutcome({ stride: 192, prefix: 4 }, "android")).toBe("unplayable")
		expect(transportStreamOutcome({ stride: 204, prefix: 0 }, "android")).toBe("unplayable")
	})

	it("keeps any stream on iOS, which previews none", () => {
		expect(transportStreamOutcome({ stride: 188, prefix: 0 }, "ios")).toBe("transportStream")
		expect(transportStreamOutcome({ stride: 192, prefix: 4 }, "ios")).toBe("transportStream")
	})
})

// A body served in `chunk`-byte pieces on demand, counting what was pulled.
function servedBody(total: number, chunk: number) {
	const served = { bytes: 0, cancelled: false }
	const body = new ReadableStream<Uint8Array>(
		{
			pull(controller) {
				const size = Math.min(chunk, total - served.bytes)

				if (size <= 0) {
					controller.close()

					return
				}

				served.bytes += size
				controller.enqueue(new Uint8Array(size).fill(0x47))
			},
			cancel() {
				served.cancelled = true
			}
		},
		{ highWaterMark: 0 }
	)

	return { served, body }
}

function respond(status: number, body: ReadableStream<Uint8Array>) {
	return { status, body }
}

describe("readRangeFromUrl", () => {
	beforeEach(() => {
		mockFetch.mockReset()
	})

	it("reads a local copy's head through a file handle, with no request", async () => {
		fs.set("file:///cache/capture.ts", new Uint8Array([1, 2, 3, 4, 5, 6]))

		const bytes = await readRangeFromUrl("file:///cache/capture.ts", new AbortController().signal)(2, 3)

		expect([...bytes]).toEqual([3, 4, 5])
		expect(mockFetch).not.toHaveBeenCalled()
		await expect(readRangeFromUrl("file:///cache/capture.ts", new AbortController().signal)(4, 10)).rejects.toThrow("short read")
	})

	it("asks the HTTP provider for exactly the range and reads only it", async () => {
		const { served, body } = servedBody(817, 400)

		mockFetch.mockResolvedValue(respond(206, body))

		const bytes = await readRangeFromUrl("http://127.0.0.1:1/file?x", new AbortController().signal)(0, 817)

		expect(bytes.byteLength).toBe(817)
		expect(mockFetch).toHaveBeenCalledWith("http://127.0.0.1:1/file?x", expect.objectContaining({ headers: { Range: "bytes=0-816" } }))
		expect(served.bytes).toBe(817)
	})

	it("fails a short 206", async () => {
		mockFetch.mockResolvedValue(respond(206, servedBody(500, 400).body))

		await expect(readRangeFromUrl("http://127.0.0.1:1/file", new AbortController().signal)(0, 817)).rejects.toThrow("short read")
	})

	it("cuts off a server that answers with the whole file, having read no more than one piece past the request", async () => {
		const { served, body } = servedBody(64 * 1024 * 1024, 1024)
		let signal: AbortSignal | undefined

		mockFetch.mockImplementation(async (_url: string, init: { signal: AbortSignal }) => {
			signal = init.signal

			return respond(200, body)
		})

		await expect(readRangeFromUrl("http://127.0.0.1:1/file", new AbortController().signal)(0, 817)).rejects.toThrow("range not served")
		expect(served.bytes).toBe(1024)
		expect(served.cancelled).toBe(true)
		expect(signal?.aborted).toBe(true)
	})

	it("takes a 200 no longer than the request as the whole of a file that small", async () => {
		const { body } = servedBody(500, 400)

		mockFetch.mockResolvedValue(respond(200, body))

		const bytes = await readRangeFromUrl("http://127.0.0.1:1/file", new AbortController().signal)(0, 500)

		expect(bytes.byteLength).toBe(500)
	})

	it("never takes a 200 for a range that does not start the file, or any other status", async () => {
		mockFetch.mockResolvedValue(respond(200, servedBody(10, 10).body))

		await expect(readRangeFromUrl("http://127.0.0.1:1/file", new AbortController().signal)(4, 10)).rejects.toThrow("range not served")

		mockFetch.mockResolvedValue(respond(500, servedBody(10, 10).body))

		await expect(readRangeFromUrl("http://127.0.0.1:1/file", new AbortController().signal)(0, 10)).rejects.toThrow("range not served")
	})

	it("follows the caller's abort", async () => {
		const outer = new AbortController()
		let signal: AbortSignal | undefined

		mockFetch.mockImplementation(async (_url: string, init: { signal: AbortSignal }) => {
			signal = init.signal
			outer.abort()

			return respond(206, servedBody(10, 10).body)
		})

		await readRangeFromUrl("http://127.0.0.1:1/file", outer.signal)(0, 10)

		expect(signal?.aborted).toBe(true)
	})
})
