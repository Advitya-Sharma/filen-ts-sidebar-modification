import { describe, expect, it } from "vitest"
import { asByteStream, imageSniffStream } from "@/lib/media/probe/imageSniff"

// The SDK's image decoder reads its input with a BYOB reader, which a TransformStream's readable doesn't offer: art
// made from a byte range (sidecar posters, embedded covers) goes through asByteStream.

const PNG = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 1, 2, 3, 4])

async function readByob(stream: ReadableStream<Uint8Array>): Promise<number[]> {
	const reader = stream.getReader({ mode: "byob" })
	const out: number[] = []

	for (;;) {
		const { done, value } = await reader.read(new Uint8Array(5))

		if (done) {
			return out
		}

		out.push(...value)
	}
}

describe("asByteStream", () => {
	it("lets a BYOB reader read what a sniffing transform passes on", async () => {
		const sniff = imageSniffStream()

		expect(() => sniff.stream.readable.getReader({ mode: "byob" })).toThrow(TypeError)

		const writer = sniff.stream.writable.getWriter()
		const reading = readByob(asByteStream(sniff.stream.readable))

		await writer.write(PNG.subarray(0, 6))
		await writer.write(PNG.subarray(6))
		await writer.close()

		expect(await reading).toEqual([...PNG])
	})

	it("hands on a source error and cancels the source when read no more", async () => {
		const sniff = imageSniffStream()
		const writer = sniff.stream.writable.getWriter()
		const reading = readByob(asByteStream(sniff.stream.readable))

		void writer.write(new Uint8Array(16)).catch(() => undefined)

		await expect(reading).rejects.toThrow("not an image")
		expect(sniff.notImage()).toBe(true)

		let cancelled: unknown = null
		const source = new ReadableStream<Uint8Array>({
			cancel(reason) {
				cancelled = reason
			}
		})

		await asByteStream(source).cancel("done")

		expect(cancelled).toBe("done")
	})
})
