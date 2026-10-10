import { brotliDecompressSync } from "node:zlib"

// Pulls the untransformed tables out of a WOFF2 file and wraps them in a minimal sfnt, which is all the
// metrics reader needs. glyf/loca (transformed by WOFF2) are left out: rebuilding outlines is the
// browser's job.

const SIGNATURE = 0x774f4632
const COLLECTION = 0x74746366
const HEADER_SIZE = 48

// The 63 tags a table directory can name by index, four characters each.
const KNOWN_TAGS =
	"cmapheadhheahmtxmaxpnameOS/2postcvt fpgmglyflocaprepCFF VORGEBDTEBLCgasphdmxkernLTSHPCLTVDMXvheavmtxBASEGDEFGPOSGSUBEBSCJSTF" +
	"MATHCBDTCBLCCOLRCPALSVG sbixacntavarbdatblocbslncvarfdscfeatfmtxfvargvarhstyjustlcarmortmorxopbdproptrakZapfSilfGlatGlocFeatSill"

export const METRIC_TABLES = ["cmap", "head", "hhea", "hmtx", "maxp", "name", "OS/2"]

interface Entry {
	tag: string
	transformed: boolean
	length: number
}

export function sfntFromWoff2(woff2: Uint8Array, tags: readonly string[] = METRIC_TABLES): Uint8Array {
	const data = new DataView(woff2.buffer, woff2.byteOffset, woff2.byteLength)

	if (woff2.length < HEADER_SIZE || data.getUint32(0) !== SIGNATURE) throw new Error("Not a WOFF2 file")

	const flavor = data.getUint32(4)

	if (flavor === COLLECTION) throw new Error("WOFF2 collections are not supported")

	const numTables = data.getUint16(12)
	const compressedLength = data.getUint32(20)
	const entries: Entry[] = []
	let at = HEADER_SIZE

	const base128 = (): number => {
		let value = 0

		for (let i = 0; i < 5; i++) {
			const byte = data.getUint8(at++)

			if (i === 0 && byte === 0x80) throw new Error("UIntBase128 with a leading zero")

			value = value * 128 + (byte & 0x7f)

			if ((byte & 0x80) === 0) return value
		}

		throw new Error("UIntBase128 longer than 5 bytes")
	}

	for (let i = 0; i < numTables; i++) {
		const flags = data.getUint8(at++)
		const index = flags & 0x3f
		let tag: string

		if (index === 63) {
			tag = String.fromCharCode(data.getUint8(at), data.getUint8(at + 1), data.getUint8(at + 2), data.getUint8(at + 3))
			at += 4
		} else {
			tag = KNOWN_TAGS.slice(index * 4, index * 4 + 4)
		}

		const version = flags >>> 6
		const length = base128()
		// glyf and loca use version 3 for "not transformed"; every other table uses 0.
		const transformed = tag === "glyf" || tag === "loca" ? version !== 3 : version !== 0

		entries.push({ tag, transformed, length: transformed ? base128() : length })
	}

	const stream = new Uint8Array(brotliDecompressSync(woff2.subarray(at, at + compressedLength)))
	const tables = new Map<string, Uint8Array>()
	let offset = 0

	for (const entry of entries) {
		if (tags.includes(entry.tag)) {
			if (entry.transformed) throw new Error(`Table ${entry.tag} is transformed`)

			tables.set(entry.tag, stream.subarray(offset, offset + entry.length))
		}

		offset += entry.length
	}

	return buildSfnt(flavor, tables)
}

export function buildSfnt(flavor: number, tables: ReadonlyMap<string, Uint8Array>): Uint8Array {
	const tags = [...tables.keys()].sort()
	let size = 12 + tags.length * 16

	for (const tag of tags) size += ((tables.get(tag)?.length ?? 0) + 3) & ~3

	const out = new Uint8Array(size)
	const view = new DataView(out.buffer)
	const power = 2 ** Math.floor(Math.log2(Math.max(1, tags.length)))

	view.setUint32(0, flavor)
	view.setUint16(4, tags.length)
	view.setUint16(6, power * 16)
	view.setUint16(8, Math.log2(power))
	view.setUint16(10, tags.length * 16 - power * 16)

	let offset = 12 + tags.length * 16

	tags.forEach((tag, i) => {
		const table = tables.get(tag) ?? new Uint8Array(0)
		const record = 12 + i * 16

		for (let k = 0; k < 4; k++) view.setUint8(record + k, tag.charCodeAt(k))

		view.setUint32(record + 8, offset)
		view.setUint32(record + 12, table.length)
		out.set(table, offset)
		offset += (table.length + 3) & ~3
	})

	return out
}
