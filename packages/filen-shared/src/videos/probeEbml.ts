import { MalformedError, utf8 } from "./probeBytes"

// EBML (RFC 8794) element headers and values, hardened for hostile input: IDs of at most 4 bytes, sizes
// of at most 8 bytes and below 2^53, children that stay inside their parent, bounded counts and depth.

export const MAX_EBML_CHILDREN = 10_000
const MAX_EBML_DEPTH = 8
export const MAX_EBML_STRING_BYTES = 4096

export interface EbmlElement {
	id: number
	start: number
	dataStart: number
	// Null for an element of unknown size.
	end: number | null
}

// A variable-length integer's byte count: the position of its first set bit, 0 for a zero byte.
function vintLength(first: number): number {
	for (let i = 0; i < 8; i++) {
		if ((first & (0x80 >> i)) !== 0) {
			return i + 1
		}
	}

	return 0
}

// The element header at `at`, reading no further than `limit`: "short" when the bytes run out first,
// "invalid" when no element can start there.
export function parseElementHeader(bytes: Uint8Array, at: number, limit: number = bytes.length): EbmlElement | "short" | "invalid" {
	const first = bytes[at]

	if (first === undefined || at >= limit) {
		return "short"
	}

	const idLength = vintLength(first)

	if (idLength === 0 || idLength > 4) {
		return "invalid"
	}

	if (at + idLength > limit) {
		return "short"
	}

	let id = 0
	let idAllOnes = true

	for (let i = 0; i < idLength; i++) {
		const byte = bytes[at + i] ?? 0

		id = id * 256 + byte
		idAllOnes &&= (i === 0 ? (byte | (0xff << (8 - idLength))) & 0xff : byte) === 0xff
	}

	// All value bits set is reserved in every ID length.
	if (idAllOnes) {
		return "invalid"
	}

	const sizeAt = at + idLength
	const sizeFirst = bytes[sizeAt]

	if (sizeFirst === undefined || sizeAt >= limit) {
		return "short"
	}

	const sizeLength = vintLength(sizeFirst)

	if (sizeLength === 0) {
		return "invalid"
	}

	if (sizeAt + sizeLength > limit) {
		return "short"
	}

	const mask = 0xff >> sizeLength
	let size = sizeFirst & mask
	let unknown = size === mask

	for (let i = 1; i < sizeLength; i++) {
		const byte = bytes[sizeAt + i] ?? 0

		size = size * 256 + byte
		unknown &&= byte === 0xff
	}

	const dataStart = sizeAt + sizeLength

	if (unknown) {
		return { id, start: at, dataStart, end: null }
	}

	if (size > Number.MAX_SAFE_INTEGER) {
		return "invalid"
	}

	return { id, start: at, dataStart, end: dataStart + size }
}

// Shared by one parse: set once anything had to be skipped as malformed.
export interface EbmlContext {
	malformed: boolean
}

// The children of [start, end), in order. Stops at the first child that is unreadable, of unknown
// size, or reaches past `end`, flagging the context; throws only past the depth cap.
export function ebmlChildren(bytes: Uint8Array, start: number, end: number, depth: number, context: EbmlContext): EbmlElement[] {
	if (depth > MAX_EBML_DEPTH) {
		throw new MalformedError("EBML nested too deep")
	}

	const children: EbmlElement[] = []
	let at = start

	while (at < end) {
		const element = parseElementHeader(bytes, at, end)

		if (typeof element === "string" || element.end === null || element.end > end || children.length >= MAX_EBML_CHILDREN) {
			context.malformed = true

			break
		}

		children.push(element)
		at = element.end
	}

	return children
}

function elementEnd(element: EbmlElement): number {
	return element.end ?? element.dataStart
}

// An unsigned integer of up to 8 bytes; null past 2^53. An empty element takes the element's declared
// default (RFC 8794 §6.1), which is `fallback`, or 0 for an element that declares none.
export function ebmlUint(bytes: Uint8Array, element: EbmlElement, fallback = 0): number | null {
	const end = elementEnd(element)

	if (end === element.dataStart) {
		return fallback
	}

	if (end - element.dataStart > 8) {
		return null
	}

	let value = 0

	for (let i = element.dataStart; i < end; i++) {
		value = value * 256 + (bytes[i] ?? 0)
	}

	return value > Number.MAX_SAFE_INTEGER ? null : value
}

// A 4- or 8-byte float; an empty element takes `fallback`, its declared default, as ebmlUint does.
export function ebmlFloat(bytes: Uint8Array, element: EbmlElement, fallback = 0): number | null {
	const length = elementEnd(element) - element.dataStart

	if (length === 0) {
		return fallback
	}

	if ((length !== 4 && length !== 8) || elementEnd(element) > bytes.length) {
		return null
	}

	const view = new DataView(bytes.buffer, bytes.byteOffset + element.dataStart, length)
	const value = length === 4 ? view.getFloat32(0) : view.getFloat64(0)

	return Number.isFinite(value) ? value : null
}

// UTF-8 (a superset of the ASCII string type), capped; null when empty or too long to be a real one.
export function ebmlString(bytes: Uint8Array, element: EbmlElement): string | null {
	const end = elementEnd(element)

	return end - element.dataStart > MAX_EBML_STRING_BYTES ? null : utf8(bytes, element.dataStart, end)
}

export function ebmlBinary(bytes: Uint8Array, element: EbmlElement): Uint8Array {
	return bytes.subarray(element.dataStart, elementEnd(element))
}

// Compressed bytes handed to the inflater at a time. zlib expands at most about 1032:1, and a write
// waits for its output to be read, so a stream stopped at the cap never holds much more than it.
const INFLATE_INPUT_SLICE = 1024

// The Compression Streams API (browsers, Node, Deno), declared with the shape used here.
interface InflateStream {
	writable: { getWriter(): { write(chunk: Uint8Array): Promise<void>; close(): Promise<void> } }
	readable: {
		getReader(): { read(): Promise<{ done: true; value?: undefined } | { done: false; value: Uint8Array }>; cancel(): Promise<void> }
	}
}

type InflateStreamCtor = new (format: "deflate") => InflateStream

// Inflates zlib data with an output cap; null when it is corrupt or would exceed the cap, or where the runtime
// has no DecompressionStream (Hermes), which reads compressed data as unreadable.
export async function inflateCapped(bytes: Uint8Array, cap: number): Promise<Uint8Array | null> {
	const Decompression = (globalThis as unknown as { DecompressionStream?: InflateStreamCtor }).DecompressionStream

	if (Decompression === undefined) {
		return null
	}

	const stream = new Decompression("deflate")
	const writer = stream.writable.getWriter()
	const reader = stream.readable.getReader()
	const feeding = (async () => {
		for (let at = 0; at < bytes.length; at += INFLATE_INPUT_SLICE) {
			await writer.write(bytes.slice(at, at + INFLATE_INPUT_SLICE))
		}

		await writer.close()
	})().catch(() => undefined)
	const chunks: Uint8Array[] = []
	let length = 0

	try {
		for (;;) {
			const { done, value } = await reader.read()

			if (done) {
				break
			}

			length += value.length

			if (length > cap) {
				await reader.cancel()
				await feeding

				return null
			}

			chunks.push(value)
		}
	} catch {
		await reader.cancel().catch(() => undefined)
		await feeding

		return null
	}

	await feeding

	const out = new Uint8Array(length)
	let at = 0

	for (const chunk of chunks) {
		out.set(chunk, at)
		at += chunk.length
	}

	return out
}
