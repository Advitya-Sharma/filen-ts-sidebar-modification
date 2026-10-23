import { constants, crc32 as nodeCrc32, deflateRawSync, inflateRawSync } from "node:zlib"
import { unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { ZipArchive, ZipError, bytesSource, memorySink, writeZip, type ZipPlanItem } from "@filen/office/core/opc"
import { SyncDeflate } from "../core/opc/zipWriter"

type Piece = Uint8Array | { readonly deflated: Uint8Array; readonly crc: number; readonly size: number }

// What a pre-deflating producer hands the writer: a fresh stream per piece, ending in a sync flush.
function predeflated(raw: Uint8Array): Piece {
	return { deflated: deflateRawSync(raw, { finishFlush: constants.Z_SYNC_FLUSH }), crc: nodeCrc32(raw), size: raw.length }
}

function markup(seed: number, size: number): Uint8Array {
	let xml = ""

	for (let k = 0; xml.length < size; k++) xml += `<w:p><w:r><w:t>Paragraph ${seed}.${k} Ж ${(k * 7919) % 104729}</w:t></w:r></w:p>`

	return new TextEncoder().encode(xml.slice(0, size))
}

function joined(raws: readonly Uint8Array[]): Uint8Array {
	const out = new Uint8Array(raws.reduce((size, raw) => size + raw.length, 0))
	let at = 0

	for (const raw of raws) {
		out.set(raw, at)
		at += raw.length
	}

	return out
}

// Content whose raw form (chunks) must not be read when it is deflated: that is the work the pieces spare.
function content(pieces: readonly Piece[], raws: readonly Uint8Array[], length = joined(raws).length) {
	return {
		length,
		chunks: async function* (): AsyncGenerator<Uint8Array> {
			yield* raws
		},
		pieces: async function* (): AsyncGenerator<Piece> {
			yield* pieces
		}
	}
}

async function zipOf(items: readonly ZipPlanItem[], forceZip64 = false, level?: 0 | 1 | 6): Promise<Uint8Array> {
	const sink = memorySink()

	await writeZip({ items }, null, sink, { forceZip64, ...(level === undefined ? {} : { level }) })

	return sink.bytes()
}

async function readBoth(zip: Uint8Array, name: string): Promise<{ fflate: Uint8Array | undefined; own: Uint8Array; crc: number }> {
	const archive = await ZipArchive.open(bytesSource(zip))
	const entry = archive.get(name)

	if (entry === undefined) throw new Error(`no entry ${name}`)

	return { fflate: unzipSync(zip)[name], own: await archive.read(entry), crc: entry.crc32 }
}

describe("writeZip with pre-deflated pieces", () => {
	it("splices a single pre-deflated piece into a valid entry with its CRC and sizes", async () => {
		const raw = markup(1, 300_000)
		const bytes = content([predeflated(raw)], [raw])
		let chunksRead = false
		const zip = await zipOf([
			{
				kind: "data",
				name: "word/document.xml",
				bytes: { ...bytes, chunks: () => ((chunksRead = true), bytes.chunks()) },
				compress: true
			}
		])
		const read = await readBoth(zip, "word/document.xml")

		expect(chunksRead).toBe(false)
		expect(read.fflate).toEqual(raw)
		expect(read.own).toEqual(raw)
		expect(read.crc).toBe(nodeCrc32(raw))
	})

	it("mixes raw and pre-deflated pieces in one deflate stream, ending on either kind", async () => {
		const [a, b, c, d, e, f] = [
			markup(2, 70_000),
			markup(3, 262_144),
			markup(4, 5),
			markup(5, 262_144),
			markup(6, 100_000),
			markup(7, 1)
		]
		const raws = [a, b, c, d, e, f]
		const endsRaw: Piece[] = [a, predeflated(b), c, predeflated(d), predeflated(e), f]
		const endsDeflated = endsRaw.slice(0, 5)
		const zip = await zipOf([
			{ kind: "data", name: "raw-end.xml", bytes: content(endsRaw, raws), compress: true },
			{ kind: "data", name: "deflated-end.xml", bytes: content(endsDeflated, raws.slice(0, 5)), compress: true }
		])

		for (const [name, expected] of [
			["raw-end.xml", joined(raws)],
			["deflated-end.xml", joined(raws.slice(0, 5))]
		] as const) {
			const read = await readBoth(zip, name)

			expect(read.fflate).toEqual(expected)
			expect(read.own).toEqual(expected)
			expect(read.crc).toBe(nodeCrc32(expected))
		}
	})

	it("ends a raw run before a pre-deflated piece on a byte boundary, wherever its last bit falls", async () => {
		const piece = markup(12, 3000)

		for (const level of [0, 1, 6] as const) {
			for (let size = 1; size < 400_000; size = Math.floor(size * 1.6) + 1) {
				const raw = markup(13, size)
				const zip = await zipOf(
					[{ kind: "data", name: "run.xml", bytes: content([raw, predeflated(piece)], [raw, piece]), compress: true }],
					false,
					level
				)

				expect(unzipSync(zip)["run.xml"]).toEqual(joined([raw, piece]))
			}
		}
	})

	it("ends a stored (level 0) deflate stream on a sync flush that other deflated content can follow", () => {
		for (const size of [0, 1, 65_535, 65_536, 140_000]) {
			const raw = markup(14, size)
			const stream = new SyncDeflate(0)

			stream.push(raw.subarray(0, size >> 1))
			stream.push(raw.subarray(size >> 1))

			const { parts, size: length } = stream.end(false)
			const flushed = joined(parts)

			expect(flushed.length).toBe(length)
			expect(Array.from(flushed.subarray(-5))).toEqual([0, 0, 0, 0xff, 0xff])
			expect(inflateRawSync(joined([flushed, new Uint8Array([0x03, 0x00])]))).toEqual(Buffer.from(raw))
		}
	})

	it("stores such content raw when the entry is not compressed", async () => {
		const raw = markup(8, 90_000)
		const zip = await zipOf([{ kind: "data", name: "stored.xml", bytes: content([predeflated(raw)], [raw]), compress: false }])
		const read = await readBoth(zip, "stored.xml")

		expect(read.own).toEqual(raw)
		expect(read.crc).toBe(nodeCrc32(raw))
	})

	it("writes pre-deflated pieces under ZIP64 fields", async () => {
		const [first, second] = [markup(9, 200_000), markup(10, 50_000)]
		const raws = [first, second]
		const zip = await zipOf(
			[{ kind: "data", name: "big.xml", bytes: content([predeflated(first), second], raws), compress: true }],
			true
		)
		const read = await readBoth(zip, "big.xml")

		expect(read.fflate).toEqual(joined(raws))
		expect(read.own).toEqual(joined(raws))
	})

	it("refuses pieces that do not add up to the declared length", async () => {
		const raw = markup(11, 1000)
		const error = await zipOf([
			{ kind: "data", name: "short.xml", bytes: content([predeflated(raw)], [raw], 1001), compress: true }
		]).catch((caught: unknown) => caught)

		expect(error).toBeInstanceOf(ZipError)
		expect((error as ZipError).code).toBe("corrupt")
	})
})
