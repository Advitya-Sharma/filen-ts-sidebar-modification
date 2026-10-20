import { deflateSync } from "fflate"
import { crc32 } from "@filen/office/core/opc"

// Hand-built zips for the layouts real writers produce rarely (ZIP64, data descriptors) and the damaged
// or hostile ones no writer should produce.

export interface BuildEntry {
	name: string
	data: Uint8Array
	method?: 0 | 8 | 99
	// Raw compressed bytes to store instead of compressing `data`.
	compressed?: Uint8Array
	descriptor?: "signed" | "unsigned" | "signed64"
	zip64?: boolean
	flags?: number
	crc?: number
	compressedSize?: number
	uncompressedSize?: number
	localName?: string
	// Point the central record at another entry's local header.
	localHeaderOf?: number
}

export interface BuildOptions {
	zip64End?: boolean
	comment?: string
	// Bytes after the end record.
	trailing?: Uint8Array
	totalEntries?: number
}

const encoder = new TextEncoder()

export { ZIP_ENTRY_TIME } from "./zipTime"

function u16(value: number): number[] {
	return [value & 0xff, (value >>> 8) & 0xff]
}

function u32(value: number): number[] {
	return [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff]
}

function u64(value: number): number[] {
	return [...u32(value % 0x1_0000_0000), ...u32(Math.floor(value / 0x1_0000_0000))]
}

export function buildZip(entries: BuildEntry[], options: BuildOptions = {}): Uint8Array {
	const chunks: Uint8Array[] = []
	const out = { length: 0 }
	const push = (bytes: ArrayLike<number>): void => {
		chunks.push(Uint8Array.from(bytes))
		out.length += bytes.length
	}
	const central: number[][] = []
	const offsets: number[] = []

	for (const entry of entries) {
		const method = entry.method ?? 8
		const compressed = entry.compressed ?? (method === 8 ? deflateSync(entry.data) : entry.data)
		const crc = entry.crc ?? crc32(entry.data)
		const compressedSize = entry.compressedSize ?? compressed.length
		const uncompressedSize = entry.uncompressedSize ?? entry.data.length
		const descriptor = entry.descriptor
		const flags = (entry.flags ?? 0) | (descriptor !== undefined ? 8 : 0)
		const name = encoder.encode(entry.name)
		const localName = encoder.encode(entry.localName ?? entry.name)
		const localExtra = entry.zip64 ? [...u16(1), ...u16(16), ...u64(uncompressedSize), ...u64(compressedSize)] : []
		const offset = out.length

		offsets.push(offset)

		push([...u32(0x04034b50), ...u16(entry.zip64 ? 45 : 20), ...u16(flags), ...u16(method), ...u16(0), ...u16(0x21)])

		if (descriptor !== undefined) {
			push([...u32(0), ...u32(0), ...u32(0)])
		} else {
			push([...u32(crc), ...u32(entry.zip64 ? 0xffffffff : compressedSize), ...u32(entry.zip64 ? 0xffffffff : uncompressedSize)])
		}

		push([...u16(localName.length), ...u16(localExtra.length)])
		push(localName)
		push(localExtra)
		push(compressed)

		if (descriptor === "signed") push([...u32(0x08074b50), ...u32(crc), ...u32(compressedSize), ...u32(uncompressedSize)])
		if (descriptor === "unsigned") push([...u32(crc), ...u32(compressedSize), ...u32(uncompressedSize)])
		if (descriptor === "signed64") push([...u32(0x08074b50), ...u32(crc), ...u64(compressedSize), ...u64(uncompressedSize)])

		const localOffset = entry.localHeaderOf !== undefined ? (offsets[entry.localHeaderOf] ?? 0) : offset
		const centralExtra = entry.zip64
			? [...u16(1), ...u16(24), ...u64(uncompressedSize), ...u64(compressedSize), ...u64(localOffset)]
			: []

		central.push([
			...u32(0x02014b50),
			...u16(entry.zip64 ? 45 : 20),
			...u16(entry.zip64 ? 45 : 20),
			...u16(flags),
			...u16(method),
			...u16(0),
			...u16(0x21),
			...u32(crc),
			...u32(entry.zip64 ? 0xffffffff : compressedSize),
			...u32(entry.zip64 ? 0xffffffff : uncompressedSize),
			...u16(name.length),
			...u16(centralExtra.length),
			...u16(0),
			...u16(0),
			...u16(0),
			...u32(0),
			...u32(entry.zip64 ? 0xffffffff : localOffset),
			...name,
			...centralExtra
		])
	}

	const directoryOffset = out.length

	for (const record of central) push(record)

	const directorySize = out.length - directoryOffset
	const comment = encoder.encode(options.comment ?? "")
	const count = options.totalEntries ?? entries.length

	if (options.zip64End) {
		const recordOffset = out.length

		push([...u32(0x06064b50), ...u64(44), ...u16(45), ...u16(45), ...u32(0), ...u32(0), ...u64(count), ...u64(count)])
		push([...u64(directorySize), ...u64(directoryOffset)])
		push([...u32(0x07064b50), ...u32(0), ...u64(recordOffset), ...u32(1)])
		push([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(0xffff), ...u16(0xffff), ...u32(0xffffffff), ...u32(0xffffffff)])
	} else {
		push([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(count), ...u16(count), ...u32(directorySize), ...u32(directoryOffset)])
	}

	push(u16(comment.length))
	push(comment)

	if (options.trailing) push(options.trailing)

	const bytes = new Uint8Array(out.length)
	let offset = 0

	for (const chunk of chunks) {
		bytes.set(chunk, offset)
		offset += chunk.length
	}

	return bytes
}

export function text(value: string): Uint8Array {
	return encoder.encode(value)
}

// Byte equality without vitest's per-element diffing, which takes seconds on megabyte arrays.
export function sameBytes(actual: Uint8Array | undefined, expected: Uint8Array | undefined): boolean {
	return (
		actual !== undefined && expected !== undefined && Buffer.from(actual.buffer, actual.byteOffset, actual.byteLength).equals(expected)
	)
}
