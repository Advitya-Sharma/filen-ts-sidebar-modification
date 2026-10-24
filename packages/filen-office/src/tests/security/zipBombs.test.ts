import { describe, expect, it } from "vitest"
import { ZipArchive, ZipError, bytesSource, type ByteSource, type ZipLimits } from "@filen/office/core/opc"
import { openDocxPackage } from "@filen/office/core/docx"
import { buildZip } from "../zipBuilder"
import { deflateRawZeros } from "./generate"
import { budget } from "../timing"

// Zip-level bombs: every refusal is typed and comes before anything large is allocated or inflated.

async function codeOf(run: () => Promise<unknown>): Promise<string | undefined> {
	try {
		await run()

		return undefined
	} catch (error) {
		if (error instanceof ZipError) return error.code

		throw error
	}
}

const open = (bytes: Uint8Array, limits?: Partial<ZipLimits>): Promise<ZipArchive> =>
	ZipArchive.open(bytesSource(bytes), limits === undefined ? {} : { limits })

const encoder = new TextEncoder()

// Byte offset of the central directory (from the end record of an archive without a comment).
function directoryOffset(bytes: Uint8Array): number {
	return new DataView(bytes.buffer, bytes.byteOffset).getUint32(bytes.length - 6, true)
}

describe("overlapping entries", () => {
	it("refuses an entry whose data runs over the next entry's local header", async () => {
		// The non-recursive bomb: one deflate stream shared by many directory entries. Stretching the first
		// entry's compressed size over the second is the minimal form of that overlap.
		const bytes = buildZip([
			{ name: "a.xml", data: encoder.encode("a".repeat(5000)) },
			{ name: "b.xml", data: encoder.encode("b".repeat(5000)) }
		])
		const view = new DataView(bytes.buffer)
		const central = directoryOffset(bytes)

		view.setUint32(central + 20, bytes.length - 200, true)

		expect(await codeOf(() => open(bytes))).toBe("corrupt")
	})

	it("refuses many directory entries pointing at one local header", async () => {
		const data = encoder.encode("x".repeat(1000))
		const entries = Array.from({ length: 50 }, (_, k) => ({ name: `e${k}`, data, ...(k > 0 ? { localHeaderOf: 0 } : {}) }))

		expect(await codeOf(() => open(buildZip(entries)))).toBe("corrupt")
	})
})

describe("declared counts and sizes", () => {
	it("refuses a directory past the entry limit after reading only the directory", async () => {
		const entries = Array.from({ length: 101 }, (_, k) => ({ name: `p${k}`, data: new Uint8Array(0), method: 0 as const }))
		const bytes = buildZip(entries)
		const started = performance.now()

		expect(await codeOf(() => open(bytes, { maxEntries: 100 }))).toBe("tooManyEntries")
		// The many-tiny-entries CPU bomb at the default limit: refused in linear time.
		expect(
			await codeOf(() =>
				open(buildZip(Array.from({ length: 20_001 }, (_, k) => ({ name: `t${k}`, data: new Uint8Array(0), method: 0 as const }))))
			)
		).toBe("tooManyEntries")
		expect(performance.now() - started).toBeLessThan(budget(5000))
	})

	it("refuses an end record declaring more entries than the directory holds", async () => {
		const bytes = buildZip([{ name: "a", data: encoder.encode("a") }], { totalEntries: 2 })

		expect(await codeOf(() => open(bytes))).toBe("truncated")
	})

	it("refuses a ZIP64 end record declaring 2^40 entries before reading its directory", async () => {
		const bytes = buildZip([{ name: "a", data: encoder.encode("a") }], { zip64End: true, totalEntries: 2 ** 40 })
		const reads: number[] = []
		const inner = bytesSource(bytes)
		const source: ByteSource = { size: inner.size, read: (offset, length) => (reads.push(length), inner.read(offset, length)) }

		expect(await codeOf(() => ZipArchive.open(source))).toBe("tooManyEntries")
		expect(Math.max(...reads)).toBeLessThanOrEqual(bytes.length)
	})

	it("refuses ZIP64 sizes past the limits, and offsets past 2^53, without allocating them", async () => {
		const huge = buildZip([{ name: "a", data: encoder.encode("a"), zip64: true, uncompressedSize: 2 ** 52 }])

		expect(await codeOf(() => open(huge))).toBe("tooLarge")

		const far = buildZip([{ name: "a", data: encoder.encode("a"), zip64: true }])
		const view = new DataView(far.buffer)
		// The central ZIP64 extra field holds uncompressed, compressed, offset; set the offset's high word.
		const extra = directoryOffset(far) + 46 + 1 + 4

		view.setUint32(extra + 20, 0x0020_0000, true)

		expect(await codeOf(() => open(far))).toBe("tooLarge")
	})

	it("refuses a total that only the sum of many entries exceeds", async () => {
		const entries = Array.from({ length: 4 }, (_, k) => ({
			name: `big${k}`,
			data: encoder.encode("x"),
			uncompressedSize: 200 * 1024 * 1024
		}))

		expect(await codeOf(() => open(buildZip(entries)))).toBe("tooLarge")
	})
})

describe("inflate cut-off", () => {
	it("stops a 1 GiB stream declared as 1 MiB after one step past the declared size", async () => {
		const compressed = deflateRawZeros(1024)
		const bytes = buildZip([{ name: "bomb.xml", data: new Uint8Array(0), compressed, uncompressedSize: 1024 * 1024, crc: 0 }])
		const archive = await open(bytes)
		const entry = archive.get("bomb.xml")

		if (entry === undefined) throw new Error("missing entry")

		const started = performance.now()

		expect(await codeOf(() => archive.read(entry))).toBe("corrupt")
		// Inflating the whole stream takes far longer than the cut-off does.
		expect(performance.now() - started).toBeLessThan(budget(500))
	})

	it("stops pulling input from an injected inflate that never ends", async () => {
		let pulled = 0
		const bytes = buildZip([
			{ name: "x", data: new Uint8Array(0), compressed: new Uint8Array(4096), uncompressedSize: 64 * 1024, crc: 0 }
		])
		const archive = await ZipArchive.open(bytesSource(bytes), {
			inflate: async function* () {
				for (;;) {
					pulled++
					yield await Promise.resolve(new Uint8Array(16 * 1024))
				}
			}
		})
		const entry = archive.get("x")

		if (entry === undefined) throw new Error("missing entry")

		expect(await codeOf(() => archive.read(entry))).toBe("corrupt")
		expect(pulled).toBeLessThanOrEqual(5)
	})
})

describe("documents", () => {
	it("refuses a package past the limits with a typed reason, as a lower-memory platform asks", async () => {
		const entries = [
			{ name: "[Content_Types].xml", data: encoder.encode("<Types/>") },
			{ name: "word/document.xml", data: new Uint8Array(0), compressed: deflateRawZeros(64), uncompressedSize: 64 << 20, crc: 0 }
		]
		const result = await openDocxPackage(bytesSource(buildZip(entries)), { limits: { maxTotalUncompressed: 32 << 20 } })

		expect(result.ok ? undefined : result.reason).toBe("tooLarge")
	})
})
