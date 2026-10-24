import { describe, expect, it } from "vitest"
import { ZipArchive, ZipError } from "@filen/office/core/opc"
import { openDocxPackage } from "@filen/office/core/docx"
import { compoundFile, countingSource, directoryEntry, type CompoundFile } from "./generate"
import { budget } from "../timing"

// Telling encrypted OOXML from legacy binary files (core/cfb's reader under openDocxPackage): every walk is
// capped and no count in the file is trusted, so a hostile header costs a bounded number of reads and answers
// with a kind.

interface Sniffed {
	readonly reason: string | undefined
	readonly reads: number
	readonly bytesRead: number
}

async function sniff(file: CompoundFile | Uint8Array): Promise<Sniffed> {
	const source = countingSource(file instanceof Uint8Array ? file : file.bytes)
	const result = await openDocxPackage(source)

	return {
		reason: result.ok ? undefined : result.reason,
		reads: source.reads.length,
		bytesRead: source.reads.reduce((sum, read) => sum + read.length, 0)
	}
}

// A minimal valid layout ([MS-CFB] §2.2): FAT in sector 0, directory in sector 1, empty streams, the root's
// children linked as a right-leaning tree in directory order (§2.6.4).
function wellFormed(names: readonly string[]): CompoundFile {
	const file = compoundFile(9, 3)
	const { view } = file

	view.setUint16(0x18, 0x3e, true)
	view.setUint16(0x1a, 3, true)
	view.setUint16(0x1c, 0xfffe, true)
	view.setUint16(0x20, 6, true)
	view.setUint32(0x2c, 1, true)
	view.setUint32(0x30, 1, true)
	view.setUint32(0x38, 4096, true)
	view.setUint32(0x3c, 0xfffffffe, true)
	view.setUint32(0x44, 0xfffffffe, true)
	view.setUint32(0x4c, 0, true)
	view.setUint32(file.sector(0) + 0, 0xfffffffd, true)
	view.setUint32(file.sector(0) + 4, 0xfffffffe, true)

	for (let k = 2; k < 128; k++) view.setUint32(file.sector(0) + k * 4, 0xffffffff, true)

	names.forEach((name, k) => directoryEntry(file, file.sector(1) + k * 128, name, k === 0 ? 5 : 2))

	const order = names
		.map((name, k) => ({ name, k }))
		.slice(1)
		.sort((a, b) => a.name.length - b.name.length || (a.name.toUpperCase() < b.name.toUpperCase() ? -1 : 1))

	for (let k = 0; k < names.length; k++) {
		const at = file.sector(1) + k * 128

		view.setUint32(at + 0x44, 0xffffffff, true)
		view.setUint32(at + 0x48, 0xffffffff, true)
		view.setUint32(at + 0x4c, 0xffffffff, true)
		view.setUint32(at + 0x74, 0xfffffffe, true)
	}

	view.setUint32(file.sector(1) + 0x4c, order[0]?.k ?? 0xffffffff, true)
	order.forEach((entry, k) => view.setUint32(file.sector(1) + entry.k * 128 + 0x48, order[k + 1]?.k ?? 0xffffffff, true))

	return file
}

describe("compound file kinds", () => {
	it("tells encrypted OOXML from legacy binary files and other compound files", async () => {
		expect((await sniff(wellFormed(["Root Entry", "EncryptionInfo", "EncryptedPackage"]))).reason).toBe("encrypted")
		expect((await sniff(wellFormed(["Root Entry", "WordDocument", "1Table"]))).reason).toBe("legacyBinary")
		expect((await sniff(wellFormed(["Root Entry", "Something"]))).reason).toBe("compoundFile")
	})

	it("reports a compound file from the zip reader with a typed error", async () => {
		const file = wellFormed(["Root Entry"])

		await expect(ZipArchive.open(countingSource(file.bytes))).rejects.toSatisfy(
			(error: unknown) => error instanceof ZipError && error.code === "compoundFile"
		)
	})
})

describe("hostile compound files", () => {
	it("stops on a directory chain that loops on itself", async () => {
		const file = wellFormed(["Root Entry", "EncryptionInfo"])

		file.view.setUint32(file.sector(0) + 4, 1, true)

		const result = await sniff(file)

		expect(result.reason).toBe("compoundFile")
		expect(result.reads).toBeLessThan(10)
	})

	it("stops on a DIFAT chain that loops and declares 2^32 - 1 sectors", async () => {
		const file = compoundFile(9, 4)
		const { view } = file

		// A directory sector whose FAT sector lies far past the header's 109 DIFAT slots.
		view.setUint32(0x30, 0xfff00000, true)
		view.setUint32(0x44, 2, true)
		view.setUint32(0x48, 0xffffffff, true)
		view.setUint32(file.sector(2) + 127 * 4, 2, true)

		const started = performance.now()
		const result = await sniff(file)

		expect(result.reason).toBe("compoundFile")
		expect(result.reads).toBeLessThan(10)
		expect(performance.now() - started).toBeLessThan(budget(1000))
	})

	it("caps a long directory chain at a bounded number of sector reads", async () => {
		const sectors = 6000
		const file = compoundFile(12, sectors + 8)
		const { view } = file
		const perSector = 1024

		view.setUint32(0x30, 8, true)

		for (let fat = 0; fat < 7; fat++) view.setUint32(0x4c + fat * 4, fat, true)
		for (let s = 8; s < sectors + 8; s++) view.setUint32(file.sector(Math.floor(s / perSector)) + (s % perSector) * 4, s + 1, true)

		const result = await sniff(file)

		expect(result.reason).toBe("compoundFile")
		expect(result.reads).toBeLessThanOrEqual(4096 + 16)
	})

	it.each([
		["a sector shift that is neither 9 nor 12", (file: CompoundFile) => file.view.setUint16(0x1e, 7, true)],
		["sector numbers past the file", (file: CompoundFile) => file.view.setUint32(0x30, 0xfffffff0, true)],
		["a FAT sector past the file", (file: CompoundFile) => file.view.setUint32(0x4c, 0x7fffffff, true)],
		[
			"random bytes after the signature",
			(file: CompoundFile) => {
				for (let k = 8; k < file.bytes.length; k++) file.bytes[k] = (k * 2654435761) >>> 24
			}
		],
		["directory names longer than an entry", (file: CompoundFile) => file.view.setUint16(file.sector(1) + 0x40, 0xffff, true)]
	])("answers a kind for %s", async (_, damage) => {
		const file = wellFormed(["Root Entry", "EncryptionInfo", "EncryptedPackage"])

		damage(file)

		const result = await sniff(file)

		expect(["encrypted", "legacyBinary", "compoundFile"]).toContain(result.reason)
		expect(result.bytesRead).toBeLessThanOrEqual(file.bytes.length * 2)
	})

	it("answers for a file holding only the signature", async () => {
		expect((await sniff(Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1))).reason).toBe("compoundFile")
	})
})
