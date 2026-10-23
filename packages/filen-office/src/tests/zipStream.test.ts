import { deflateRawSync } from "node:zlib"
import { describe, expect, it } from "vitest"
import { ZipArchive, ZipError, bytesSource, type InflateHook, type ZipEntry } from "@filen/office/core/opc"
import { buildZip, sameBytes, text } from "./zipBuilder"

// ZipArchive.stream: an entry's content in chunks, never whole, checked as read() checks it.

const nativeInflate: InflateHook = input => ReadableStream.from(input).pipeThrough(new DecompressionStream("deflate-raw"))

function repetitive(size: number): Uint8Array {
	const out = new Uint8Array(size)

	for (let i = 0; i < size; i++) out[i] = (i * 7 + (i >> 9)) & 0xff

	return out
}

function entryOf(archive: ZipArchive, name: string): ZipEntry {
	const entry = archive.get(name)

	if (entry === undefined) throw new Error(`no ${name}`)

	return entry
}

async function collect(archive: ZipArchive, name: string): Promise<{ bytes: Uint8Array; largest: number; chunks: number }> {
	const parts: Uint8Array[] = []
	let largest = 0

	for await (const chunk of archive.stream(entryOf(archive, name))) {
		// A chunk is the caller's until the next one is asked for: copied here, as a consumer that keeps it must.
		parts.push(chunk.slice())
		largest = Math.max(largest, chunk.length)
	}

	const bytes = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
	let at = 0

	for (const part of parts) {
		bytes.set(part, at)
		at += part.length
	}

	return { bytes, largest, chunks: parts.length }
}

async function failure(run: () => Promise<unknown>): Promise<string> {
	try {
		await run()
	} catch (error) {
		return error instanceof ZipError ? `${error.code}: ${error.message}` : `not a ZipError: ${String(error)}`
	}

	return "no error"
}

const BIG = repetitive(3 * 1024 * 1024)

describe("ZipArchive.stream", () => {
	for (const [label, inflate] of [
		["fflate", undefined],
		["a native inflate", nativeInflate]
	] as const) {
		it(`streams a deflated entry in bounded chunks through ${label}`, async () => {
			const archive = await ZipArchive.open(
				bytesSource(buildZip([{ name: "big.xml", data: BIG }])),
				inflate === undefined ? {} : { inflate }
			)
			const { bytes, largest, chunks } = await collect(archive, "big.xml")

			expect(sameBytes(bytes, BIG)).toBe(true)
			expect(largest).toBeLessThanOrEqual(2 * 1024 * 1024)
			expect(chunks).toBeGreaterThan(4)
		})
	}

	it("streams a stored entry in bounded views", async () => {
		const archive = await ZipArchive.open(bytesSource(buildZip([{ name: "s.bin", data: BIG, method: 0 }])))
		const { bytes, largest } = await collect(archive, "s.bin")

		expect(sameBytes(bytes, BIG)).toBe(true)
		expect(largest).toBeLessThanOrEqual(1024 * 1024)
	})

	it("streams an empty entry as nothing", async () => {
		const archive = await ZipArchive.open(bytesSource(buildZip([{ name: "e", data: new Uint8Array(0) }])))

		expect((await collect(archive, "e")).bytes.length).toBe(0)
	})

	it("fails a CRC mismatch, a short entry and an overlong one exactly as read() does", async () => {
		const data = text("hello world, hello world, hello world")
		const cases = [
			buildZip([{ name: "a", data, crc: 0x12345678 }]),
			buildZip([{ name: "a", data, uncompressedSize: data.length + 5 }]),
			buildZip([{ name: "a", data, uncompressedSize: data.length - 5 }]),
			buildZip([{ name: "a", data, method: 0, crc: 1 }]),
			buildZip([{ name: "a", data: text("x"), compressed: new Uint8Array([0xff, 0xff, 0xff]) }])
		]

		for (const zip of cases) {
			for (const inflate of [undefined, nativeInflate]) {
				const archive = await ZipArchive.open(bytesSource(zip), inflate === undefined ? {} : { inflate })
				const read = await failure(() => archive.read(entryOf(archive, "a")))
				const streamed = await failure(() => collect(archive, "a"))

				expect(read).not.toBe("no error")
				expect(streamed).toBe(read)
			}
		}
	})

	it("stops inflating once the entry outgrows its declared size", async () => {
		const bomb = deflateRawSync(new Uint8Array(64 * 1024 * 1024))
		const archive = await ZipArchive.open(
			bytesSource(buildZip([{ name: "bomb", data: new Uint8Array(1000), compressed: bomb, uncompressedSize: 1000 }]))
		)
		let seen = 0
		const message = await failure(async () => {
			for await (const chunk of archive.stream(entryOf(archive, "bomb"))) seen += chunk.length
		})

		expect(message).toContain("inflates past its declared size")
		expect(seen).toBeLessThanOrEqual(1000)
	})

	it("refuses a compression method it does not read", async () => {
		const archive = await ZipArchive.open(bytesSource(buildZip([{ name: "m", data: text("x"), method: 99 }])))

		expect(await failure(() => collect(archive, "m"))).toBe(await failure(() => archive.read(entryOf(archive, "m"))))
	})
})
