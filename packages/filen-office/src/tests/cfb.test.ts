import { describe, expect, it } from "vitest"
import { CfbError, CfbFile, cfbBytes, compareNames, planCfb, writeCfb, type CfbErrorCode, type CfbWriteEntry } from "@filen/office/core/cfb"
import { BytePieces, bytesSource, memorySink, type ByteChunks } from "@filen/office/core/opc"

// Deterministic bytes (xorshift32).
function noise(length: number, seed: number): Uint8Array {
	const out = new Uint8Array(length)
	let state = seed || 1

	for (let k = 0; k < length; k++) {
		state ^= state << 13
		state ^= state >>> 17
		state ^= state << 5
		out[k] = state & 0xff
	}

	return out
}

// Byte equality without vitest's element-wise diff (which runs out of memory on megabytes).
function same(a: Uint8Array, b: Uint8Array): boolean {
	return Buffer.from(a.buffer, a.byteOffset, a.byteLength).equals(Buffer.from(b.buffer, b.byteOffset, b.byteLength))
}

function stream(path: string, bytes: Uint8Array): CfbWriteEntry {
	return { path: path.split("/"), bytes }
}

async function open(bytes: Uint8Array): Promise<CfbFile> {
	return CfbFile.open(bytesSource(bytes))
}

async function rejectsWith(promise: Promise<unknown>, code: CfbErrorCode): Promise<void> {
	await expect(promise).rejects.toSatisfy((error: unknown) => error instanceof CfbError && error.code === code)
}

function view(bytes: Uint8Array): DataView {
	return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

// The DataSpaces layout an encrypted OOXML package carries ([MS-OFFCRYPTO] §2.3.4.1), with stand-in contents.
function encryptedPackage(packageLength: number): CfbWriteEntry[] {
	return [
		stream("EncryptionInfo", noise(1_120, 1)),
		stream("EncryptedPackage", noise(packageLength, 2)),
		stream("\u0006DataSpaces/Version", noise(76, 3)),
		stream("\u0006DataSpaces/DataSpaceMap", noise(112, 4)),
		stream("\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace", noise(64, 5)),
		stream("\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary", noise(200, 6))
	]
}

async function expectStreams(file: CfbFile, entries: readonly CfbWriteEntry[]): Promise<void> {
	for (const entry of entries) {
		if (!("bytes" in entry)) continue

		expect(await file.read(entry.path), entry.path.join("/")).toEqual(entry.bytes)
	}
}

// Sector `index` of a version 3 file: its byte offset.
const sector = (index: number): number => (index + 1) * 512
const fatEntry = (index: number): number => sector(0) + index * 4

describe("compound file writer and reader", () => {
	for (const version of [3, 4] as const) {
		it(`round-trips an encrypted package layout (version ${version})`, async () => {
			const entries = encryptedPackage(70_001)
			const bytes = await cfbBytes(entries, { version })
			const file = await open(bytes)

			expect(file.version).toBe(version)
			expect(file.sectorSize).toBe(version === 4 ? 4096 : 512)
			await expectStreams(file, entries)
			expect(file.entry("\u0006DataSpaces/TransformInfo")?.kind).toBe("storage")
			expect([...file.walk()].map(entry => entry.path.join("/"))).toEqual([
				"",
				"\u0006DataSpaces",
				"\u0006DataSpaces/Version",
				"\u0006DataSpaces/DataSpaceMap",
				"\u0006DataSpaces/DataSpaceInfo",
				"\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace",
				"\u0006DataSpaces/TransformInfo",
				"\u0006DataSpaces/TransformInfo/StrongEncryptionTransform",
				"\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary",
				"EncryptionInfo",
				"EncryptedPackage"
			])
		})
	}

	it("writes the same bytes for the same input, whatever the entry order", async () => {
		const entries = encryptedPackage(9_000)
		const first = await cfbBytes(entries)

		expect(await cfbBytes([...entries].reverse())).toEqual(first)
		expect(await cfbBytes(entries)).toEqual(first)
	})

	it("writes nothing but what is needed", async () => {
		// Header, FAT, directory (2 entries), 10 sectors of stream.
		expect(planCfb([stream("A", noise(5_000, 1))]).size).toBe(13 * 512)
		// Header, FAT, directory, mini FAT, mini stream (2 x 64 bytes).
		expect(planCfb([stream("A", noise(100, 1))]).size).toBe(5 * 512)
		// Version 4: a 4096-byte header sector, then FAT, directory and one stream sector.
		expect(planCfb([stream("A", noise(4_096, 1))], { version: 4 }).size).toBe(4 * 4096)
		// An empty stream and an empty storage take a directory entry only.
		expect(planCfb([stream("A", new Uint8Array(0)), { path: ["S"], storage: true }]).size).toBe(3 * 512)
	})

	it("puts streams under 4096 bytes in the mini stream and the rest in sectors", async () => {
		const entries = [0, 1, 63, 64, 65, 4_095, 4_096, 4_097, 20_000].map((length, k) => stream(`S${k}`, noise(length, k + 1)))

		for (const version of [3, 4] as const) await expectStreams(await open(await cfbBytes(entries, { version })), entries)
	})

	it("keeps CLSIDs of storages and the root", async () => {
		const clsid = "{00020906-0000-0000-C000-000000000046}"
		const file = await open(
			await cfbBytes([stream("ObjectPool/_1/Data", noise(10, 1)), { path: ["ObjectPool", "_1"], storage: true, clsid }], {
				rootClsid: clsid
			})
		)

		expect(file.root.clsid).toBe(clsid)
		expect(file.entry("ObjectPool/_1")?.clsid).toBe(clsid)
		expect(file.entry("ObjectPool")?.clsid).toBeUndefined()
	})

	it("sorts siblings by length, then upper-cased, and finds names case-insensitively", async () => {
		const names = ["bb", "a", "B", "Ab", "ä", "aB2", "Zz", "ccc"].filter(
			(name, k, all) => all.findIndex(other => compareNames(other, name) === 0) === k
		)
		const file = await open(await cfbBytes(names.map((name, k) => stream(name, noise(10, k + 1)))))

		expect(file.root.children.map(entry => entry.name)).toEqual([...names].sort(compareNames))
		expect(file.entry("ZZ")?.name).toBe("Zz")
		expect(file.entry(["Ä"])?.name).toBe("ä")
		expect(file.entry("nope")).toBeUndefined()
	})

	it("writes content given in pieces without joining it, and refuses pieces of the wrong length", async () => {
		const parts = [noise(3_000, 1), noise(9_000, 2), noise(10, 3)]
		const pieces = BytePieces.of(parts)
		const file = await open(await cfbBytes([{ path: ["P"], bytes: pieces }]))

		expect(await file.read("P")).toEqual(pieces.copy(0, pieces.length))

		const lying: ByteChunks = {
			length: 100,
			async *chunks() {
				yield noise(50, 1)
			}
		}

		await rejectsWith(cfbBytes([{ path: ["P"], bytes: lying }]), "lengthMismatch")
	})

	it("streams to a sink in batches, ending at the planned size", async () => {
		const plan = planCfb(encryptedPackage(300_000))
		const sink = memorySink()
		const writes: number[] = []

		await writeCfb(plan, {
			write: chunk => {
				writes.push(chunk.length)

				return sink.write(chunk)
			},
			close: () => sink.close()
		})

		expect(sink.bytes().length).toBe(plan.size)
		await rejectsWith(writeCfb({ size: plan.size, version: 3 }, sink), "notFound")
		expect(writes.length).toBeLessThan(10)
		await expectStreams(await open(sink.bytes()), encryptedPackage(300_000))
	})

	it("lists FAT sectors past the header's 109 in DIFAT sectors", { timeout: 60_000 }, async () => {
		// 16 MiB of 512-byte sectors needs over 256 FAT sectors: 109 in the header, the rest in two DIFAT sectors.
		const big = new Uint8Array(16 * 1024 * 1024)

		big.set(noise(1_000, 7), big.length - 1_000)

		const bytes = await cfbBytes([stream("Big", big), stream("Small", noise(10, 8))])
		const header = view(bytes)

		expect(header.getUint32(0x2c, true)).toBe(259)
		expect(header.getUint32(0x48, true)).toBe(2)

		const file = await open(bytes)

		expect(same(await file.read("Big"), big)).toBe(true)
		expect(await file.read("Small")).toEqual(noise(10, 8))

		// A DIFAT chain pointing back at itself, or ending early.
		const difat = header.getUint32(0x44, true)
		const next = header.getUint32(sector(difat) + 508, true)

		header.setUint32(sector(difat) + 508, difat, true)
		await rejectsWith(open(bytes), "corrupt")
		header.setUint32(sector(difat) + 508, 0xfffffffe, true)
		await rejectsWith(open(bytes), "corrupt")
		header.setUint32(sector(difat) + 508, next, true)
		await expect(open(bytes)).resolves.toBeInstanceOf(CfbFile)
	})

	it("refuses bad names, duplicates and streams on a storage path", async () => {
		await rejectsWith(cfbBytes([stream("a/b", noise(1, 1)), { path: ["x:y"], bytes: noise(1, 1) }]), "invalidName")
		await rejectsWith(cfbBytes([{ path: ["x".repeat(32)], bytes: noise(1, 1) }]), "invalidName")
		await rejectsWith(cfbBytes([{ path: [], bytes: noise(1, 1) }]), "invalidName")
		await rejectsWith(cfbBytes([stream("Foo", noise(1, 1)), stream("FOO", noise(1, 1))]), "duplicateName")
		await rejectsWith(cfbBytes([stream("A", noise(1, 1)), stream("A/B", noise(1, 1))]), "duplicateName")
		await rejectsWith(
			cfbBytes([
				{ path: ["S"], storage: true },
				{ path: ["s"], storage: true }
			]),
			"duplicateName"
		)
		await rejectsWith(cfbBytes([{ path: ["S"], storage: true, clsid: "nope" }]), "invalidClsid")
	})
})

describe("compound file reader", () => {
	it("reads ranges of a stream, as views of the source within one run of sectors", async () => {
		const content = noise(50_000, 3)
		const bytes = await cfbBytes([stream("A", content), stream("B", noise(20_000, 4))])
		const file = await open(bytes)
		const source = file.stream("A")

		expect(source.size).toBe(content.length)

		const inside = await source.read(1_000, 4_000)

		expect(inside.buffer).toBe(bytes.buffer)
		expect(inside).toEqual(content.subarray(1_000, 5_000))
		expect(await source.read(49_990, 10)).toEqual(content.subarray(49_990))
		await rejectsWith(source.read(49_990, 11), "truncated")
		await rejectsWith(file.read("Nope"), "notFound")
		expect(() => file.stream("Nope")).toThrow(CfbError)
	})

	it("joins a stream scattered over the file", async () => {
		const content = noise(10_240, 9)
		const bytes = await cfbBytes([stream("A", content)])
		// Stream sectors 2..21 (after the FAT and the directory): swap the last two.
		const data = view(bytes)
		const copy = bytes.slice(sector(20), sector(21))

		bytes.copyWithin(sector(20), sector(21), sector(22))
		bytes.set(copy, sector(21))
		data.setUint32(fatEntry(19), 21, true)
		data.setUint32(fatEntry(21), 20, true)
		data.setUint32(fatEntry(20), 0xfffffffe, true)

		const file = await open(bytes)

		expect(await file.read("A")).toEqual(content)
		expect(await file.stream("A").read(9_000, 1_240)).toEqual(content.subarray(9_000))
	})

	it("reads a file whose last sector is cut short, and refuses a stream cut short", async () => {
		const content = noise(5_000, 5)
		const bytes = await cfbBytes([stream("A", content)])
		// The last sector holds 5000 - 9 * 512 = 392 bytes of the stream.
		const file = await open(bytes.subarray(0, bytes.length - 100))

		expect(await file.read("A")).toEqual(content)
		await rejectsWith((await open(bytes.subarray(0, bytes.length - 121))).read("A"), "truncated")
	})

	it("refuses a chain through the file's cut-short last sector before the stream's end", async () => {
		const bytes = await cfbBytes([stream("A", noise(5_000, 5))])
		const data = view(bytes)

		// Stream sectors 2..11; 11, the file's last, loses 100 bytes and moves up the chain: 2..9, 11, 10.
		data.setUint32(fatEntry(9), 11, true)
		data.setUint32(fatEntry(11), 10, true)
		data.setUint32(fatEntry(10), 0xfffffffe, true)
		await rejectsWith((await open(bytes.subarray(0, bytes.length - 100))).read("A"), "truncated")
	})

	it("ignores the high half of version 3 stream sizes", async () => {
		const bytes = await cfbBytes([stream("A", noise(100, 1))])

		view(bytes).setUint32(sector(1) + 128 + 0x7c, 0xdeadbeef, true)
		expect((await open(bytes)).entry("A")?.size).toBe(100)
	})

	it("refuses what is not a compound file, other versions and sector sizes", async () => {
		await rejectsWith(open(noise(1_000, 1)), "notCompoundFile")
		await rejectsWith(open(new Uint8Array(3)), "notCompoundFile")

		const good = await cfbBytes([stream("A", noise(100, 1))])

		await rejectsWith(open(good.subarray(0, 300)), "truncated")

		for (const [offset, value, code] of [
			[0x1a, 5, "unsupported"],
			[0x1c, 0xfeff, "corrupt"],
			[0x1e, 10, "unsupported"],
			[0x20, 7, "unsupported"]
		] as const) {
			const bytes = good.slice()

			view(bytes).setUint16(offset, value, true)
			await rejectsWith(open(bytes), code)
		}

		const cutoff = good.slice()

		view(cutoff).setUint32(0x38, 8192, true)
		await rejectsWith(open(cutoff), "unsupported")
		// A version 4 header with fewer than 4096 bytes.
		await rejectsWith(open((await cfbBytes([], { version: 4 })).subarray(0, 2_000)), "truncated")
	})

	it("holds to the size and entry limits", async () => {
		const bytes = await cfbBytes(Array.from({ length: 40 }, (_, k) => stream(`S${k}`, noise(10, k + 1))))

		await expect(CfbFile.open(bytesSource(bytes), { limits: { maxEntries: 41 } })).resolves.toBeInstanceOf(CfbFile)
		await rejectsWith(CfbFile.open(bytesSource(bytes), { limits: { maxEntries: 20 } }), "tooManyEntries")
		await rejectsWith(CfbFile.open(bytesSource(bytes), { limits: { maxSize: bytes.length - 1 } }), "tooLarge")
	})

	describe("hostile chains and trees", () => {
		// FAT, directory, mini FAT, mini stream; A: 5000 bytes in sectors 4..13; B: 6000 bytes in sectors
		// 14..25; C: 100 bytes in the mini stream.
		async function sample(): Promise<Uint8Array> {
			return cfbBytes([stream("A", noise(5_000, 1)), stream("B", noise(6_000, 2)), stream("C", noise(100, 3))])
		}

		function directoryEntry(bytes: Uint8Array, name: string): number {
			for (let at = sector(1); at < sector(2); at += 128) {
				let found = ""

				for (let k = 0; k < name.length; k++) found += String.fromCharCode(view(bytes).getUint16(at + k * 2, true))

				if (found === name && view(bytes).getUint16(at + 0x40, true) === (name.length + 1) * 2) return at
			}

			throw new Error(`no entry ${name}`)
		}

		async function readAll(bytes: Uint8Array): Promise<void> {
			const file = await open(bytes)

			for (const entry of file.walk()) if (entry.kind === "stream") await file.read(entry)
		}

		it("finds the sample where it expects it", async () => {
			const bytes = await sample()
			const a = directoryEntry(bytes, "A")

			expect(view(bytes).getUint32(a + 0x74, true)).toBe(4)
			await readAll(bytes)
		})

		it("refuses a chain that loops", async () => {
			const bytes = await sample()

			view(bytes).setUint32(fatEntry(8), 5, true)
			await rejectsWith(readAll(bytes), "corrupt")
		})

		it("refuses a chain shorter than its stream, or leaving the file", async () => {
			const short = await sample()

			view(short).setUint32(fatEntry(8), 0xfffffffe, true)
			await rejectsWith(readAll(short), "corrupt")

			const outside = await sample()

			view(outside).setUint32(fatEntry(8), 1_000_000, true)
			await rejectsWith(readAll(outside), "corrupt")
		})

		it("refuses two streams sharing sectors, and stream sizes beyond the file", async () => {
			const shared = await sample()

			view(shared).setUint32(directoryEntry(shared, "B") + 0x74, 4, true)
			await rejectsWith(readAll(shared), "corrupt")

			const huge = await sample()

			view(huge).setUint32(directoryEntry(huge, "B") + 0x78, 0x7fffffff, true)

			const file = await open(huge)

			await rejectsWith(file.read("B"), "corrupt")
			// The others stay readable.
			expect(await file.read("A")).toEqual(noise(5_000, 1))
		})

		it("refuses mini chains that loop or point past the mini stream", async () => {
			const loop = await sample()
			const miniFat = view(loop).getUint32(0x3c, true)

			view(loop).setUint32(sector(miniFat), 0, true)
			await rejectsWith(readAll(loop), "corrupt")

			const past = await sample()

			view(past).setUint32(directoryEntry(past, "C") + 0x74, 40, true)
			await rejectsWith(readAll(past), "corrupt")
		})

		it("refuses directory trees that loop or point outside the directory", async () => {
			const loop = await sample()

			view(loop).setUint32(directoryEntry(loop, "A") + 0x44, 0, true)
			await rejectsWith(open(loop), "corrupt")

			const self = await sample()
			const c = directoryEntry(self, "C")

			view(self).setUint32(c + 0x48, (c - sector(1)) / 128, true)
			await rejectsWith(open(self), "corrupt")

			const outside = await sample()

			view(outside).setUint32(directoryEntry(outside, "A") + 0x48, 900, true)
			await rejectsWith(open(outside), "corrupt")
		})

		it("refuses a directory chain that loops, or a FAT sector outside the file", async () => {
			const loop = await sample()

			view(loop).setUint32(fatEntry(1), 1, true)
			await rejectsWith(open(loop), "corrupt")

			const fat = await sample()

			view(fat).setUint32(0x4c, 5_000, true)
			await rejectsWith(open(fat), "corrupt")
		})

		it("refuses an unallocated entry or a second root in the tree, and a bad name length", async () => {
			for (const [field, value, width] of [
				[0x42, 0, 1],
				[0x42, 5, 1],
				[0x40, 65, 2],
				[0x40, 0, 2]
			] as const) {
				const bytes = await sample()
				const at = directoryEntry(bytes, "B") + field

				if (width === 1) view(bytes).setUint8(at, value)
				else view(bytes).setUint16(at, value, true)

				await rejectsWith(open(bytes), "corrupt")
			}
		})
	})
})
