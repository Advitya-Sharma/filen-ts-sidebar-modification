// Hand-built container bytes for probe tests: just enough of each box/element for the parser.

export type Part = Uint8Array | readonly number[]

export function concat(...parts: Part[]): Uint8Array {
	const length = parts.reduce((sum, part) => sum + part.length, 0)
	const out = new Uint8Array(length)
	let at = 0

	for (const part of parts) {
		out.set(part, at)
		at += part.length
	}

	return out
}

export function text(value: string): number[] {
	return Array.from(new TextEncoder().encode(value))
}

export function u16(value: number): number[] {
	return [(value >>> 8) & 0xff, value & 0xff]
}

export function u32(value: number): number[] {
	return [(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff]
}

export function u64(value: number): number[] {
	return [...u32(Math.floor(value / 0x100000000)), ...u32(value % 0x100000000)]
}

export function zeros(length: number): Uint8Array {
	return new Uint8Array(length)
}

export function f64(value: number): number[] {
	const bytes = new Uint8Array(8)

	new DataView(bytes.buffer).setFloat64(0, value)

	return Array.from(bytes)
}

// ── ISO BMFF ──

export function box(type: string, ...payload: Part[]): Uint8Array {
	const body = concat(...payload)

	return concat(u32(8 + body.length), text(type), body)
}

// A box with a 64-bit size (size field 1, then largesize).
export function largeBox(type: string, ...payload: Part[]): Uint8Array {
	const body = concat(...payload)

	return concat(u32(1), text(type), u64(16 + body.length), body)
}

export function fullBox(type: string, version: number, flags: number, ...payload: Part[]): Uint8Array {
	return box(type, [version, (flags >>> 16) & 0xff, (flags >>> 8) & 0xff, flags & 0xff], ...payload)
}

export const FTYP = box("ftyp", text("isom"), u32(0x200), text("isomiso2mp41"))

// tkhd version 0 with a 16.16 display matrix (a, b, c, d).
export function tkhd({ id, enabled = true, matrix = [1, 0, 0, 1] }: { id: number; enabled?: boolean; matrix?: number[] }): Uint8Array {
	const [a = 1, b = 0, c = 0, d = 1] = matrix
	const fixed = (value: number): number[] => u32((value * 0x10000) >>> 0)

	return fullBox(
		"tkhd",
		0,
		enabled ? 3 : 0,
		zeros(8),
		u32(id),
		zeros(4),
		u32(0),
		zeros(8),
		zeros(8),
		fixed(a),
		fixed(b),
		u32(0),
		fixed(c),
		fixed(d),
		u32(0),
		zeros(8),
		u32(0x40000000),
		zeros(8)
	)
}

// Packed ISO 639-2/T, three 5-bit letters.
export function mdhd({
	timescale = 1000,
	duration = 0,
	language = "und"
}: {
	timescale?: number
	duration?: number
	language?: string
}): Uint8Array {
	const code = Array.from(language).reduce((packed, letter) => packed * 32 + (letter.charCodeAt(0) - 0x60), 0)

	return fullBox("mdhd", 0, 0, zeros(8), u32(timescale), u32(duration), u16(code), zeros(2))
}

export function hdlr(handler: string): Uint8Array {
	return fullBox("hdlr", 0, 0, zeros(4), text(handler), zeros(12), [0])
}

export function stsd(entry: Uint8Array): Uint8Array {
	return fullBox("stsd", 0, 0, u32(1), entry)
}

export function videoEntry(
	fourcc: string,
	{ width = 1280, height = 720 }: { width?: number; height?: number },
	...children: Uint8Array[]
): Uint8Array {
	return box(fourcc, zeros(6), u16(1), zeros(16), u16(width), u16(height), zeros(50), ...children)
}

export function audioEntry(
	fourcc: string,
	{ channels = 2, rate = 48000 }: { channels?: number; rate?: number },
	...children: Uint8Array[]
): Uint8Array {
	return box(fourcc, zeros(6), u16(1), zeros(8), u16(channels), u16(16), zeros(4), u32(rate * 0x10000), ...children)
}

export function trak(header: Uint8Array, mdiaChildren: Uint8Array[], entry: Uint8Array, ...extra: Uint8Array[]): Uint8Array {
	return box("trak", header, ...extra, box("mdia", ...mdiaChildren, box("minf", box("stbl", stsd(entry), box("stts", zeros(8))))))
}

// mvhd version 0: timescale, duration and next_track_ID.
export function mvhd({
	timescale = 1000,
	duration = 0,
	nextTrackId = 0
}: {
	timescale?: number
	duration?: number
	nextTrackId?: number
}): Uint8Array {
	return fullBox("mvhd", 0, 0, zeros(8), u32(timescale), u32(duration), zeros(76), u32(nextTrackId))
}

// esds: ES_Descriptor → DecoderConfigDescriptor (object type, 12 bytes) → DecoderSpecificInfo.
export function esds(objectType: number, config: number[] = []): Uint8Array {
	const specific = config.length === 0 ? [] : [5, config.length, ...config]
	const decoderConfig = [4, 13 + specific.length, objectType, ...Array.from(zeros(12)), ...specific]

	return fullBox("esds", 0, 0, [3, 3 + decoderConfig.length, 0, 1, 0], decoderConfig)
}

// ── EBML ──

export function ebmlId(id: number): number[] {
	const bytes: number[] = []

	for (let value = id; value > 0; value = Math.floor(value / 256)) {
		bytes.unshift(value & 0xff)
	}

	return bytes
}

// A 4-byte size: marker bit 0x10 in the first byte.
export function element(id: number, ...payload: Part[]): Uint8Array {
	const body = concat(...payload)

	return concat(
		ebmlId(id),
		[0x10 | ((body.length >>> 24) & 0x0f), (body.length >>> 16) & 0xff, (body.length >>> 8) & 0xff, body.length & 0xff],
		body
	)
}

// An element of unknown size (all-ones 8-byte size).
export function unknownSize(id: number, ...payload: Part[]): Uint8Array {
	return concat(ebmlId(id), [0x01, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff], ...payload)
}

export function uintElement(id: number, value: number): Uint8Array {
	const bytes: number[] = []

	for (let rest = value; bytes.length === 0 || rest > 0; rest = Math.floor(rest / 256)) {
		bytes.unshift(rest % 256)
	}

	return element(id, bytes)
}

export function stringElement(id: number, value: string): Uint8Array {
	return element(id, text(value))
}

export function ebmlHeader(docType = "matroska"): Uint8Array {
	return element(0x1a45dfa3, stringElement(0x4282, docType))
}

export function segment(...children: Uint8Array[]): Uint8Array {
	return unknownSize(0x18538067, ...children)
}
