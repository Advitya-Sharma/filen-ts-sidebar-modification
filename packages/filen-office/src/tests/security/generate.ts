import { bytesSource, crc32, type ByteSource } from "@filen/office/core/opc"
import type { FontExtents, TextMeasurer } from "@filen/office/core/text"
import type { DocxEngineServices } from "@filen/office/dom/docx"

// Hostile inputs built in memory: nothing hostile is committed in binary form.

function join(parts: readonly Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
	let at = 0

	for (const part of parts) {
		out.set(part, at)
		at += part.length
	}

	return out
}

// Raw deflate of `mebibytes` MiB of zeros as one fixed-code block: a literal zero, then length-258 copies
// at distance 1 (13 bits each), about 6.5 KiB per MiB. One block, so an inflater has no block boundary
// to stop at.
export function deflateRawZeros(mebibytes: number): Uint8Array {
	const total = mebibytes * 1024 * 1024
	const copies = Math.floor((total - 1) / 258)
	const rest = total - 1 - copies * 258
	const out = new Uint8Array(Math.ceil((copies * 13 + rest * 8 + 64) / 8))
	let bit = 0
	const put = (value: number, count: number): void => {
		for (let k = 0; k < count; k++, bit++) if (((value >>> k) & 1) !== 0) out[bit >>> 3] = (out[bit >>> 3] ?? 0) | (1 << (bit & 7))
	}
	// Huffman codes go most significant bit first.
	const code = (value: number, count: number): void => {
		for (let k = count - 1; k >= 0; k--, bit++) if (((value >>> k) & 1) !== 0) out[bit >>> 3] = (out[bit >>> 3] ?? 0) | (1 << (bit & 7))
	}

	put(1, 1)
	put(1, 2)
	code(0x30, 8)

	for (let k = 0; k < copies; k++) {
		code(0xc5, 8)
		code(0, 5)
	}

	for (let k = 0; k < rest; k++) code(0x30, 8)

	code(0, 7)

	return out.subarray(0, (bit + 7) >>> 3)
}

// The same in a gzip member. The trailer's CRC is left zero: the stream is refused long before it is checked.
export function gzipZeros(mebibytes: number): Uint8Array {
	const size = new Uint8Array(4)

	new DataView(size.buffer).setUint32(0, (mebibytes * 1024 * 1024) % 2 ** 32, true)

	return join([Uint8Array.of(0x1f, 0x8b, 8, 0, 0, 0, 0, 0, 2, 0xff), deflateRawZeros(mebibytes), new Uint8Array(4), size])
}

// The same in a zlib stream (Adler-32 left zero, as above).
export function zlibZeros(mebibytes: number): Uint8Array {
	return join([Uint8Array.of(0x78, 0xda), deflateRawZeros(mebibytes), new Uint8Array(4)])
}

// An element opened `depth` times around `inner`.
export function nest(open: string, close: string, inner: string, depth: number): string {
	return open.repeat(depth) + inner + close.repeat(depth)
}

export interface CountingSource extends ByteSource {
	readonly reads: { readonly offset: number; readonly length: number }[]
	// Whether any read touched [start, end).
	touched(start: number, end: number): boolean
}

export function countingSource(bytes: Uint8Array): CountingSource {
	const inner = bytesSource(bytes)
	const reads: { offset: number; length: number }[] = []

	return {
		size: inner.size,
		reads,
		read(offset, length) {
			reads.push({ offset, length })

			return inner.read(offset, length)
		},
		touched(start, end) {
			return reads.some(read => read.offset < end && read.offset + read.length > start)
		}
	}
}

// --- PNG ---

function pngChunk(type: string, data: Uint8Array): Uint8Array {
	const out = new Uint8Array(12 + data.length)
	const view = new DataView(out.buffer)

	view.setUint32(0, data.length)

	for (let k = 0; k < 4; k++) out[4 + k] = type.charCodeAt(k)

	out.set(data, 8)
	view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)))

	return out
}

// An 8-bit RGBA PNG declaring `width` x `height`, with `idat` as its compressed data.
export function png(width: number, height: number, idat: Uint8Array = new Uint8Array(0)): Uint8Array {
	const header = new Uint8Array(13)
	const view = new DataView(header.buffer)

	view.setUint32(0, width)
	view.setUint32(4, height)
	header[8] = 8
	header[9] = 6

	return join([
		Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10),
		pngChunk("IHDR", header),
		pngChunk("IDAT", idat),
		pngChunk("IEND", new Uint8Array(0))
	])
}

// --- Compound files (MS-CFB) ---

export interface CompoundFile {
	readonly bytes: Uint8Array
	readonly view: DataView
	readonly sectorSize: number
	// Byte offset of sector `index`.
	sector(index: number): number
}

// A compound file of `sectors` zeroed sectors after the header, every header DIFAT slot free.
export function compoundFile(shift: number, sectors: number): CompoundFile {
	const sectorSize = 2 ** shift
	const bytes = new Uint8Array(Math.max(512, sectorSize) * (sectors + 1))
	const view = new DataView(bytes.buffer)

	bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
	view.setUint16(0x1e, shift, true)

	for (let k = 0; k < 109; k++) view.setUint32(0x4c + k * 4, 0xffffffff, true)

	return { bytes, view, sectorSize, sector: index => (index + 1) * sectorSize }
}

// A directory entry (128 bytes) naming a stream.
export function directoryEntry(file: CompoundFile, at: number, name: string, type: number = 2): void {
	for (let k = 0; k < name.length; k++) file.view.setUint16(at + k * 2, name.charCodeAt(k), true)

	file.view.setUint16(at + 0x40, (name.length + 1) * 2, true)
	file.view.setUint8(at + 0x42, type)
}

// --- Engine ---

const measurer: TextMeasurer = {
	kerningControl: true,
	measure(_font, _kerning, texts, count, widths) {
		for (let i = 0; i < count; i++) widths[i] = (texts[i] ?? "").length * 8
	},
	fontExtents(): FontExtents {
		return { ascent: 12, descent: 4 }
	}
}

// Engine services without a browser: fonts are accepted and never loaded.
export function engineServices(): DocxEngineServices {
	return {
		fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
		replayFonts: false,
		fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
		measurer,
		yield: () => Promise.resolve(),
		sliceMs: Infinity
	}
}

// --- Network and execution traps ---

const TRAPPED = ["fetch", "XMLHttpRequest", "WebSocket", "EventSource", "Image", "importScripts", "open", "Worker", "eval"] as const

export interface Trap {
	readonly calls: string[]
	restore(): void
}

// Replaces every global that could reach the network or run code with one that records the call and
// throws, so a test can prove a document caused none.
export function trapGlobals(): Trap {
	const calls: string[] = []
	const saved = TRAPPED.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)

	for (const name of TRAPPED) {
		const trap = function (this: unknown, ...args: unknown[]): never {
			calls.push(`${name}(${args.map(arg => String(arg).slice(0, 80)).join(", ")})`)

			throw new Error(`${name} called`)
		}

		Object.defineProperty(globalThis, name, { value: trap, configurable: true, writable: true })
	}

	return {
		calls,
		restore() {
			for (const [name, descriptor] of saved) {
				if (descriptor === undefined) Reflect.deleteProperty(globalThis, name)
				else Object.defineProperty(globalThis, name, descriptor)
			}
		}
	}
}
