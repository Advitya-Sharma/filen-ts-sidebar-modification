import { readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { crc32 as nodeCrc32, deflateRawSync } from "node:zlib"
import { unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import {
	ZipArchive,
	ZipError,
	bytesSource,
	crc32,
	memorySink,
	planZip,
	writeZip,
	type ByteSource,
	type InflateHook,
	type ZipEntry,
	type ZipLimits,
	ZIP64_COMPRESSED,
	ZIP64_OFFSET,
	ZIP64_UNCOMPRESSED
} from "@filen/office/core/opc"
import { buildZip, sameBytes, text } from "./zipBuilder"

const FIXTURES = join(import.meta.dirname, "fixtures")
const fixtures = readdirSync(FIXTURES).filter(name => name.endsWith(".docx"))

function fixture(name: string): Uint8Array {
	return new Uint8Array(readFileSync(join(FIXTURES, name)))
}

function open(bytes: Uint8Array, limits?: Partial<ZipLimits>): Promise<ZipArchive> {
	return ZipArchive.open(bytesSource(bytes), limits === undefined ? {} : { limits })
}

async function codeOf(run: () => Promise<unknown>): Promise<string | undefined> {
	try {
		await run()
	} catch (error) {
		return error instanceof ZipError ? error.code : `not a ZipError: ${String(error)}`
	}

	return undefined
}

function entry(archive: ZipArchive, name: string): ZipEntry {
	const found = archive.get(name)

	if (found === undefined) throw new Error(`no ${name}`)

	return found
}

async function write(archive: ZipArchive, plan = planZip(archive), forceZip64 = false): Promise<Uint8Array> {
	const sink = memorySink()

	await writeZip(plan, archive.source, sink, { forceZip64 })

	return sink.bytes()
}

// Repeated content, so it deflates at a few hundred to one.
function repetitive(size: number): Uint8Array {
	const out = new Uint8Array(size)

	for (let i = 0; i < size; i++) out[i] = 97 + (i % 7)

	return out
}

describe("crc32", () => {
	it("matches a reference implementation, also when continued", () => {
		const data = repetitive(100_003)

		data[5] = 0xff

		const expected = nodeCrc32(data)

		expect(crc32(data)).toBe(expected)
		expect(crc32(data.subarray(50_001), crc32(data.subarray(0, 50_001)))).toBe(expected)
		expect(crc32(text("123456789"))).toBe(0xcbf43926)
		expect(crc32(new Uint8Array(0))).toBe(0)
	})
})

describe("ZipArchive reading", () => {
	it.each(fixtures)("reads every entry of %s as fflate does", async name => {
		const bytes = fixture(name)
		const archive = await open(bytes)
		const reference = unzipSync(bytes)

		expect(archive.entries.map(item => item.name).sort()).toEqual(Object.keys(reference).sort())

		for (const item of archive.entries) {
			expect(sameBytes(await archive.read(item), reference[item.name])).toBe(true)
		}
	})

	it("reads stored and deflated entries, UTF-8 names and the archive comment", async () => {
		const data = repetitive(5000)
		const bytes = buildZip(
			[
				{ name: "stored.bin", data, method: 0 },
				{ name: "deflated.xml", data },
				{ name: "médias/图片.png", data: text("x"), flags: 0x0800 },
				{ name: "empty", data: new Uint8Array(0) }
			],
			{ comment: "hello" }
		)
		const archive = await open(bytes)

		expect(await archive.read(entry(archive, "stored.bin"))).toEqual(data)
		expect(await archive.read(entry(archive, "deflated.xml"))).toEqual(data)
		expect(await archive.read(entry(archive, "médias/图片.png"))).toEqual(text("x"))
		expect(await archive.read(entry(archive, "empty"))).toEqual(new Uint8Array(0))
		expect(entry(archive, "deflated.xml").compressedSize).toBeLessThan(data.length)
		expect(new TextDecoder().decode(archive.comment)).toBe("hello")
	})

	it("reads entries with data descriptors in every layout", async () => {
		const data = repetitive(3000)
		const bytes = buildZip([
			{ name: "signed", data, descriptor: "signed" },
			{ name: "unsigned", data, descriptor: "unsigned" },
			{ name: "signed64", data, descriptor: "signed64", zip64: true },
			{ name: "plain", data }
		])
		const archive = await open(bytes)

		for (const name of ["signed", "unsigned", "signed64", "plain"]) {
			expect(await archive.read(entry(archive, name))).toEqual(data)
		}

		expect((await archive.layout(entry(archive, "signed"))).descriptorLength).toBe(16)
		expect((await archive.layout(entry(archive, "unsigned"))).descriptorLength).toBe(12)
		expect((await archive.layout(entry(archive, "signed64"))).descriptorLength).toBe(24)
		expect((await archive.layout(entry(archive, "plain"))).descriptorLength).toBe(0)
	})

	it("reads ZIP64 end records and extra fields", async () => {
		const data = repetitive(2000)
		const bytes = buildZip(
			[
				{ name: "a", data, zip64: true },
				{ name: "b", data, zip64: true, method: 0 }
			],
			{ zip64End: true }
		)
		const archive = await open(bytes)

		expect(archive.entries.length).toBe(2)
		expect(await archive.read(entry(archive, "a"))).toEqual(data)
		expect(await archive.read(entry(archive, "b"))).toEqual(data)
		expect(entry(archive, "a").extra.length).toBe(0)
	})

	it("reads only the tail, the directory and the entries it is asked for", async () => {
		const bytes = buildZip([
			{ name: "big", data: repetitive(200_000), method: 0 },
			{ name: "small", data: text("x") }
		])
		const reads: [number, number][] = []
		const inner = bytesSource(bytes)
		const source: ByteSource = {
			size: inner.size,
			read(offset, length) {
				reads.push([offset, length])

				return inner.read(offset, length)
			}
		}
		const archive = await ZipArchive.open(source)

		await archive.read(entry(archive, "small"))

		const [tail, ...rest] = reads

		// The tail (end record plus the longest possible comment), then nothing inside the big entry.
		expect(tail?.[1]).toBe(0xffff + 22 + 20)
		expect(rest.some(([offset, length]) => offset < 200_000 && offset + length > 40)).toBe(false)
		expect(rest.reduce((sum, [, length]) => sum + length, 0)).toBeLessThan(300)
	})
})

describe("ZipArchive refusals", () => {
	it("refuses inputs that are not zips", async () => {
		expect(await codeOf(() => open(text("not a zip at all, just some text")))).toBe("notZip")
		expect(await codeOf(() => open(new Uint8Array(10)))).toBe("notZip")
		expect(await codeOf(() => open(Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...new Uint8Array(600)])))).toBe(
			"compoundFile"
		)
	})

	it("refuses truncated archives", async () => {
		const bytes = fixture("01-plain.docx")
		const archive = await open(bytes)
		const directoryStart = archive.dataEnd

		expect(await codeOf(() => open(bytes.subarray(directoryStart + 10)))).toBe("notZip")
		expect(await codeOf(() => open(bytes.subarray(0, bytes.length - 30)))).toBe("notZip")
	})

	it("refuses damaged entries when they are read", async () => {
		const data = repetitive(4000)
		const bytes = buildZip([
			{ name: "crc", data, crc: 1234 },
			{ name: "short", data, uncompressedSize: 5000 },
			{ name: "long", data, uncompressedSize: 3000 },
			{ name: "garbage", data, compressed: Uint8Array.from([0xff, 0xff, 0xff, 0xff, 0xff, 0xff]) },
			{ name: "method", data, method: 99 },
			{ name: "renamed", data, localName: "other!!" }
		])
		const archive = await open(bytes)

		expect(await codeOf(() => archive.read(entry(archive, "crc")))).toBe("corrupt")
		expect(await codeOf(() => archive.read(entry(archive, "short")))).toBe("corrupt")
		expect(await codeOf(() => archive.read(entry(archive, "long")))).toBe("corrupt")
		expect(await codeOf(() => archive.read(entry(archive, "garbage")))).toBe("corrupt")
		expect(await codeOf(() => archive.read(entry(archive, "method")))).toBe("unsupported")
		expect(await codeOf(() => archive.read(entry(archive, "renamed")))).toBe("corrupt")
	})

	it("stops inflating as soon as an entry outgrows its declared size", async () => {
		// 64 MiB of zeros declared as 1 KiB: refused without ever holding more than one inflate step.
		const zeros = new Uint8Array(64 * 1024 * 1024)
		const compressed = new Uint8Array(deflateRawSync(zeros, { level: 9 }))
		const bytes = buildZip([{ name: "bomb", data: new Uint8Array(0), compressed, uncompressedSize: 1024, crc: 0 }])
		const archive = await open(bytes)

		expect(await codeOf(() => archive.read(entry(archive, "bomb")))).toBe("corrupt")
	})

	it("enforces the entry, size and ratio limits before inflating anything", async () => {
		const data = repetitive(2 * 1024 * 1024)
		const bytes = buildZip([
			{ name: "a", data },
			{ name: "b", data }
		])
		expect(await codeOf(() => open(bytes, { maxEntries: 1 }))).toBe("tooManyEntries")
		expect(await codeOf(() => open(bytes, { maxTotalUncompressed: 3 * 1024 * 1024 }))).toBe("tooLarge")
		expect(await codeOf(() => open(bytes, { maxEntryUncompressed: 1024 * 1024 }))).toBe("tooLarge")
		expect(await codeOf(() => open(bytes))).toBeUndefined()

		const lying = buildZip([{ name: "x", data: text("x"), uncompressedSize: 4 * 1024 * 1024 * 1024 - 2 }])

		expect(await codeOf(() => open(lying))).toBe("tooLarge")
	})

	it("opens uniform data that deflate compresses past 1000:1", async () => {
		const uniform = new Uint8Array(5_000_000).fill(0xff)
		const compressed = new Uint8Array(deflateRawSync(uniform, { level: 9 }))
		const archive = await open(buildZip([{ name: "x.bin", data: uniform, compressed }]))

		expect(uniform.length / compressed.length).toBeGreaterThan(1000)
		expect(sameBytes(await archive.read(entry(archive, "x.bin")), uniform)).toBe(true)
	})

	it("refuses overlapping entries, duplicate names and encryption", async () => {
		const data = repetitive(1000)

		expect(
			await codeOf(() =>
				open(
					buildZip([
						{ name: "a", data },
						{ name: "b", data, localHeaderOf: 0 }
					])
				)
			)
		).toBe("corrupt")
		expect(
			await codeOf(() =>
				open(
					buildZip([
						{ name: "a", data },
						{ name: "a", data }
					])
				)
			)
		).toBe("duplicateName")
		expect(await codeOf(() => open(buildZip([{ name: "a", data, flags: 1 }])))).toBe("encrypted")
		expect(await codeOf(() => open(buildZip([{ name: "a", data, method: 0, uncompressedSize: 5 }])))).toBe("corrupt")
	})

	it("refuses a descriptor that disagrees with the central directory", async () => {
		const data = repetitive(1000)
		const damaged = buildZip([{ name: "a", data, descriptor: "signed" }])
		const archive = await open(damaged)
		const layout = await archive.layout(entry(archive, "a"))
		const descriptor = layout.dataOffset + entry(archive, "a").compressedSize

		damaged[descriptor + 4] = (damaged[descriptor + 4] ?? 0) ^ 0xff

		const reopened = await open(damaged)

		expect(await codeOf(() => reopened.read(entry(reopened, "a")))).toBe("corrupt")
	})
})

describe("writeZip", () => {
	it.each(fixtures)("copies every entry of %s back byte for byte", async name => {
		const bytes = fixture(name)
		const archive = await open(bytes)

		expect(sameBytes(await write(archive), bytes)).toBe(true)
	})

	it("rewrites only the replaced entry and keeps its metadata", async () => {
		const bytes = fixture("01-plain.docx")
		const archive = await open(bytes)
		const document = entry(archive, "word/document.xml")
		const replacement = text('<?xml version="1.0"?><w:document xmlns:w="urn:w"/>')
		const out = await write(archive, planZip(archive, { replace: [{ name: "word/document.xml", bytes: replacement }] }))
		const reopened = await open(out)
		const rewritten = entry(reopened, "word/document.xml")

		expect(unzipSync(out)["word/document.xml"]).toEqual(replacement)
		expect(await reopened.read(rewritten)).toEqual(replacement)
		expect(rewritten.index).toBe(document.index)
		expect([rewritten.modTime, rewritten.modDate, rewritten.method, rewritten.externalAttributes]).toEqual([
			document.modTime,
			document.modDate,
			document.method,
			document.externalAttributes
		])

		for (const original of archive.entries) {
			if (original.name === "word/document.xml") continue

			expect(sameBytes(await reopened.raw(entry(reopened, original.name)), await archive.raw(original))).toBe(true)
		}
	})

	it("adds and removes entries", async () => {
		const archive = await open(fixture("01-plain.docx"))
		const added = repetitive(10_000)
		const out = await write(
			archive,
			planZip(archive, {
				remove: ["docProps/app.xml"],
				add: [
					{ name: "word/media/filen-1.png", bytes: added, compress: false },
					{ name: "customXml/项目.xml", bytes: added }
				]
			})
		)
		const reference = unzipSync(out)

		expect(reference["docProps/app.xml"]).toBeUndefined()
		expect(reference["word/media/filen-1.png"]).toEqual(added)
		expect(reference["customXml/项目.xml"]).toEqual(added)

		const reopened = await open(out)

		expect(entry(reopened, "word/media/filen-1.png").method).toBe(0)
		expect(entry(reopened, "customXml/项目.xml").method).toBe(8)
		expect(entry(reopened, "customXml/项目.xml").flags & 0x0800).toBe(0x0800)
		expect(entry(reopened, "customXml/项目.xml").modDate).toBe(0x21)
	})

	it("refuses plans that name an entry twice or replace a missing one", async () => {
		const archive = await open(fixture("01-plain.docx"))

		expect(() => planZip(archive, { replace: [{ name: "missing.xml", bytes: text("x") }] })).toThrow(ZipError)
		expect(await codeOf(() => write(archive, planZip(archive, { add: [{ name: "word/document.xml", bytes: text("x") }] })))).toBe(
			"duplicateName"
		)
	})

	it("deflates new content with an injected deflate", async () => {
		const archive = await open(fixture("01-plain.docx"))
		const replacement = repetitive(50_000)
		const calls: number[] = []
		const sink = memorySink()

		await writeZip(planZip(archive, { replace: [{ name: "word/document.xml", bytes: replacement }] }), archive.source, sink, {
			deflate: data => {
				calls.push(data.length)

				return Promise.resolve(new Uint8Array(deflateRawSync(data)))
			}
		})

		expect(calls).toEqual([50_000])
		expect(sameBytes(unzipSync(sink.bytes())["word/document.xml"], replacement)).toBe(true)
	})

	it("keeps an archive's ZIP64 form when copying it", async () => {
		const data = repetitive(2000)
		const bytes = buildZip(
			[
				{ name: "a", data, zip64: true },
				{ name: "b", data }
			],
			{ zip64End: true }
		)
		const archive = await open(bytes)

		expect(archive.zip64).toBe(true)
		expect(entry(archive, "a").zip64Fields).toBe(ZIP64_UNCOMPRESSED | ZIP64_COMPRESSED | ZIP64_OFFSET)
		expect(entry(archive, "b").zip64Fields).toBe(0)
		expect(sameBytes(await write(archive), bytes)).toBe(true)
	})

	it("writes entry data in physical order and the directory in its own order", async () => {
		const data = repetitive(500)
		const bytes = buildZip([
			{ name: "first", data },
			{ name: "second", data }
		])
		const archive = await open(bytes)
		const [first, second] = archive.entries

		if (first === undefined || second === undefined) throw new Error("entries missing")

		// The same data with the central records swapped, so data order and directory order differ.
		const swapped = await write(archive, {
			items: [
				{ kind: "copy", entry: { ...first, index: 1 } },
				{ kind: "copy", entry: { ...second, index: 0 } }
			]
		})
		const reopened = await open(swapped)

		expect(reopened.entries.map(item => item.name)).toEqual(["second", "first"])
		expect(planZip(reopened).items.map(item => (item.kind === "copy" ? item.entry.name : item.name))).toEqual(["first", "second"])
		expect(sameBytes(await write(reopened), swapped)).toBe(true)
	})

	it("writes ZIP64 records that it and fflate read back", async () => {
		const archive = await open(fixture("03-header-footer-footnotes.docx"))
		const replacement = repetitive(5000)
		const out = await write(archive, planZip(archive, { replace: [{ name: "word/styles.xml", bytes: replacement }] }), true)
		const reopened = await open(out)
		const reference = unzipSync(out)

		for (const original of archive.entries) {
			const expected = original.name === "word/styles.xml" ? replacement : await archive.read(original)

			expect(sameBytes(await reopened.read(entry(reopened, original.name)), expected)).toBe(true)
			expect(sameBytes(reference[original.name], expected)).toBe(true)
		}
	})

	it("copies data descriptors, ZIP64 entries and unknown methods as they are", async () => {
		const data = repetitive(3000)
		const bytes = buildZip(
			[
				{ name: "signed", data, descriptor: "signed" },
				{ name: "unsigned", data, descriptor: "unsigned" },
				{ name: "signed64", data, descriptor: "signed64", zip64: true },
				{ name: "method", data, method: 99, compressed: data }
			],
			{ zip64End: true }
		)
		const archive = await open(bytes)
		const out = await write(archive)
		const reopened = await open(out)

		for (const name of ["signed", "unsigned", "signed64"]) {
			expect(await reopened.read(entry(reopened, name))).toEqual(data)
		}

		expect(await reopened.raw(entry(reopened, "method"))).toEqual(data)
		// Everything before the central directory is the original bytes.
		expect(sameBytes(out.subarray(0, reopened.dataEnd), bytes.subarray(0, archive.dataEnd))).toBe(true)
	})

	it("hands the sink few, large chunks", async () => {
		const archive = await open(fixture("08-large-300p.docx"))
		const chunks: number[] = []
		let closed = false

		await writeZip(planZip(archive), archive.source, {
			write(chunk) {
				chunks.push(chunk.length)

				return Promise.resolve()
			},
			close() {
				closed = true

				return Promise.resolve()
			}
		})

		expect(closed).toBe(true)
		expect(chunks.length).toBeLessThanOrEqual(2)
	})

	it("streams large copied entries in bounded reads", async () => {
		const data = repetitive(3 * 1024 * 1024)
		const bytes = buildZip([{ name: "big", data, method: 0 }])
		const archive = await open(bytes)
		const lengths: number[] = []
		const inner = bytesSource(bytes)
		const sink = memorySink()

		await writeZip(
			planZip(archive),
			{
				size: inner.size,
				read(offset, length) {
					lengths.push(length)

					return inner.read(offset, length)
				}
			},
			sink,
			{ copyChunkSize: 1024 * 1024 }
		)

		expect(Math.max(...lengths)).toBeLessThanOrEqual(1024 * 1024)
		expect(sameBytes(sink.bytes(), bytes)).toBe(true)
	})
})

// The web host's inflate: a native DecompressionStream fed by the archive's step-by-step input.
const nativeInflate: InflateHook = input => ReadableStream.from(input).pipeThrough(new DecompressionStream("deflate-raw"))

function concat(...parts: Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
	let offset = 0

	for (const part of parts) {
		out.set(part, offset)
		offset += part.length
	}

	return out
}

describe("ZipArchive repairs", () => {
	const base = buildZip([
		{ name: "a.xml", data: repetitive(5000) },
		{ name: "b.xml", data: text("hello"), method: 0 }
	])

	async function expectRepaired(bytes: Uint8Array): Promise<void> {
		const archive = await open(bytes)

		expect(archive.repaired).toBe(true)
		expect(sameBytes(await archive.read(entry(archive, "a.xml")), repetitive(5000))).toBe(true)
		expect(sameBytes(await archive.read(entry(archive, "b.xml")), text("hello"))).toBe(true)
		// A rewrite drops the junk and is the clean archive again.
		expect(sameBytes(await write(archive), base)).toBe(true)
	}

	it("opens a clean archive unrepaired", async () => {
		expect((await open(base)).repaired).toBe(false)
	})

	it("ignores bytes after the archive, also beyond the longest comment", async () => {
		await expectRepaired(concat(base, new Uint8Array(100)))
		await expectRepaired(concat(base, new Uint8Array(70_000)))
		await expectRepaired(concat(base, repetitive(300_000)))
	})

	it("clamps a comment length running past the end", async () => {
		const bytes = base.slice()

		new DataView(bytes.buffer).setUint16(bytes.length - 2, 10, true)

		await expectRepaired(bytes)
	})

	it("finds the archive behind prepended bytes", async () => {
		await expectRepaired(concat(new Uint8Array(16), base))
		await expectRepaired(concat(repetitive(70_000), base))
	})

	it("reads every central record whatever the declared count says, but refuses missing ones", async () => {
		const short = base.slice()
		const data = new DataView(short.buffer)

		data.setUint16(short.length - 14, 1, true)
		data.setUint16(short.length - 12, 1, true)

		const archive = await open(short)

		expect(archive.entries.map(item => item.name)).toEqual(["a.xml", "b.xml"])
		expect(archive.repaired).toBe(true)

		const long = base.slice()

		new DataView(long.buffer).setUint16(long.length - 12, 3, true)

		expect(await codeOf(() => open(long))).toBe("truncated")
	})

	it("reads names that are neither flagged nor valid UTF-8 as CP437, and writes them back as stored", async () => {
		const bytes = buildZip([
			{ name: "ab.xml", data: text("x") },
			{ name: "ac.xml", data: text("y") }
		])

		// Same-length byte swaps in both headers: "ab" -> 0x8E 0x8E, "ac" -> 0x8E 0x8F, neither valid UTF-8.
		for (let at = bytes.indexOf(0x61); at >= 0; at = bytes.indexOf(0x61, at + 1)) {
			if (bytes[at + 2] !== 0x2e) continue
			if (bytes[at + 1] === 0x62) bytes.set([0x8e, 0x8e], at)
			if (bytes[at + 1] === 0x63) bytes.set([0x8e, 0x8f], at)
		}

		const archive = await open(bytes)

		expect(archive.entries.map(item => item.name)).toEqual(["ÄÄ.xml", "ÄÅ.xml"])

		const out = await write(archive, planZip(archive, { replace: [{ name: "ÄÄ.xml", bytes: text("z") }] }))
		const rewritten = (await open(out)).entries[0]

		expect(Array.from(rewritten?.nameBytes ?? [])).toEqual([0x8e, 0x8e, 0x2e, 0x78, 0x6d, 0x6c])
		expect((rewritten?.flags ?? 0) & 0x0800).toBe(0)
	})

	it("reads the 64-bit descriptor of an empty entry whose local header is ZIP64", async () => {
		const bytes = buildZip([
			{ name: "e", data: new Uint8Array(0), method: 0, descriptor: "signed64", zip64: true },
			{ name: "f", data: text("f") }
		])
		const archive = await open(bytes)

		expect((await archive.layout(entry(archive, "e"))).descriptorLength).toBe(24)
		expect(sameBytes(await write(archive), bytes)).toBe(true)
	})
})

describe("inflating", () => {
	it("uses an injected inflate", async () => {
		const bytes = fixture("08-large-300p.docx")
		const archive = await ZipArchive.open(bytesSource(bytes), { inflate: nativeInflate })
		const reference = unzipSync(bytes)

		for (const item of archive.entries) {
			expect(sameBytes(await archive.read(item), reference[item.name])).toBe(true)
		}
	})

	it("stops an injected inflate as soon as an entry outgrows its declared size", async () => {
		const zeros = new Uint8Array(64 * 1024 * 1024)
		const compressed = new Uint8Array(deflateRawSync(zeros, { level: 9 }))
		const bytes = buildZip([{ name: "bomb", data: new Uint8Array(0), compressed, uncompressedSize: 1024, crc: 0 }])
		let pulled = 0
		const counting: InflateHook = input =>
			nativeInflate(
				(async function* () {
					for await (const step of input) {
						pulled += step.length

						yield step
					}
				})()
			)
		const archive = await ZipArchive.open(bytesSource(bytes), { inflate: counting })

		expect(await codeOf(() => archive.read(entry(archive, "bomb")))).toBe("corrupt")
		// Steps are sized to what may still come out (1 KiB here), so only a sliver of input is consumed.
		expect(pulled).toBeLessThan(64 * 1024)
	})

	it("refuses damaged data through an injected inflate", async () => {
		const bytes = buildZip([
			{ name: "garbage", data: repetitive(4000), compressed: Uint8Array.from([0xff, 0xff, 0xff, 0xff, 0xff, 0xff]) }
		])
		const archive = await ZipArchive.open(bytesSource(bytes), { inflate: nativeInflate })

		expect(await codeOf(() => archive.read(entry(archive, "garbage")))).toBe("corrupt")
	})
})

describe("writeZip details", () => {
	it("writes entries as they come from an async source, each one only once the one before is written", async () => {
		const written: string[] = []
		const sink = {
			chunks: [] as Uint8Array[],
			write(chunk: Uint8Array) {
				this.chunks.push(chunk)
				written.push(`bytes:${String(chunk.length)}`)

				return Promise.resolve()
			},
			close: () => Promise.resolve()
		}

		async function* entries() {
			for (const name of ["one.pdf", "two.pdf"]) {
				written.push(`made:${name}`)
				yield { kind: "data" as const, name, bytes: text(`%PDF ${name}`), compress: false }
			}
		}

		await writeZip({ items: entries() }, null, sink)

		// The second entry is asked for only after the first one's bytes went out.
		expect(written.indexOf("made:two.pdf")).toBeGreaterThan(written.findIndex(step => step.startsWith("bytes:")))

		const joined = new Uint8Array(sink.chunks.reduce((sum, chunk) => sum + chunk.length, 0))
		let at = 0

		for (const chunk of sink.chunks) {
			joined.set(chunk, at)
			at += chunk.length
		}

		const archive = await ZipArchive.open(bytesSource(joined))

		expect(archive.entries.map(entry => entry.name)).toEqual(["one.pdf", "two.pdf"])
		expect(new TextDecoder().decode(await archive.read(archive.entries[1] as ZipEntry))).toBe("%PDF two.pdf")
	})

	it("stamps a new entry with the time it is given, in DOS local time, and the DOS epoch without one", async () => {
		const sink = memorySink()
		const modified = new Date(2026, 9, 23, 14, 37, 58)

		await writeZip(
			{
				items: [
					{ kind: "data", name: "dated.pdf", bytes: text("%PDF"), compress: false, modified },
					{ kind: "data", name: "plain.txt", bytes: text("x"), compress: false },
					{ kind: "data", name: "old.txt", bytes: text("y"), compress: false, modified: new Date(1975, 0, 1) }
				]
			},
			null,
			sink
		)

		const archive = await ZipArchive.open(bytesSource(sink.bytes()))
		const [dated, plain, old] = archive.entries

		expect(dated?.modDate).toBe(((2026 - 1980) << 9) | (10 << 5) | 23)
		expect(dated?.modTime).toBe((14 << 11) | (37 << 5) | 29)
		expect(plain).toMatchObject({ modDate: (1 << 5) | 1, modTime: 0 })
		// Before the DOS epoch: the epoch itself.
		expect(old).toMatchObject({ modDate: (1 << 5) | 1, modTime: 0 })
		expect(new TextDecoder().decode(await archive.read(dated as ZipEntry))).toBe("%PDF")
	})

	it("reuses the layouts an archive already read instead of reading every local header again", async () => {
		const bytes = fixture("03-header-footer-footnotes.docx")
		let reads = 0
		const inner = bytesSource(bytes)
		const source: ByteSource = {
			size: inner.size,
			read(offset, length) {
				reads++

				return inner.read(offset, length)
			}
		}
		const archive = await ZipArchive.open(source)

		for (const item of archive.entries) await archive.layout(item)

		const before = reads
		const sink = memorySink()

		await writeZip(planZip(archive), archive, sink)

		expect(sameBytes(sink.bytes(), bytes)).toBe(true)
		// One read per copied entry's bytes, nothing for headers.
		expect(reads - before).toBe(archive.entries.length)
	})

	it("keeps forced ZIP64 off stored entries, so an ODF mimetype has no extra field", async () => {
		const mimetype = text("application/vnd.oasis.opendocument.text")
		const bytes = buildZip([
			{ name: "mimetype", data: mimetype, method: 0 },
			{ name: "content.xml", data: repetitive(3000) }
		])
		const archive = await open(bytes)
		const out = await write(archive, planZip(archive, { replace: [{ name: "mimetype", bytes: mimetype }] }), true)

		// Local header: 30 bytes, the name, then the data with no extra field between.
		expect(new DataView(out.buffer).getUint16(28, true)).toBe(0)
		expect(new TextDecoder().decode(out.subarray(38, 38 + mimetype.length))).toBe("application/vnd.oasis.opendocument.text")

		const reopened = await open(out)

		expect(reopened.zip64).toBe(true)
		expect(sameBytes(await reopened.read(entry(reopened, "content.xml")), repetitive(3000))).toBe(true)
	})
})
