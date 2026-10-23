import { ByteWriter, writeHex } from "./bytes"

// PDF objects (ISO 32000-1 §7.3) as plain values, and their serialization. A JS string is a text string
// (§7.9.2.2: PDFDocEncoding when printable ASCII, else UTF-16BE with a byte order mark); byte strings are
// PdfString; numbers are written without exponents.

export class PdfName {
	readonly name: string

	constructor(name: string) {
		this.name = name
	}
}

export class PdfRef {
	readonly id: number

	constructor(id: number) {
		this.id = id
	}
}

// A byte string, written in hex (§7.3.4.3), so any bytes (ciphertext included) survive.
export class PdfString {
	readonly bytes: Uint8Array

	constructor(bytes: Uint8Array) {
		this.bytes = bytes
	}
}

// Pre-serialized PDF syntax, written as is (an array built in a ByteWriter, for example /W of a font).
export class PdfRaw {
	readonly bytes: Uint8Array

	constructor(bytes: Uint8Array) {
		this.bytes = bytes
	}
}

export type PdfValue = null | boolean | number | string | PdfName | PdfRef | PdfString | PdfRaw | readonly PdfValue[] | PdfDict

// Entries whose value is undefined are left out.
export interface PdfDict {
	readonly [key: string]: PdfValue | undefined
}

const names = new Map<string, PdfName>()

export function pdfName(name: string): PdfName {
	let value = names.get(name)

	if (value === undefined) {
		value = new PdfName(name)

		if (names.size < 4096) names.set(name, value)
	}

	return value
}

export function pdfHex(bytes: Uint8Array): PdfString {
	return new PdfString(bytes)
}

export function pdfRaw(syntax: string | Uint8Array): PdfRaw {
	return new PdfRaw(typeof syntax === "string" ? new ByteWriter(syntax.length).ascii(syntax).take() : syntax)
}

// §7.9.2.2: a text string's bytes.
export function textStringBytes(text: string): Uint8Array {
	if (isPrintableAscii(text)) return new ByteWriter(text.length).ascii(text).take()

	const out = new Uint8Array(2 + text.length * 2)

	out[0] = 0xfe
	out[1] = 0xff

	for (let index = 0; index < text.length; index++) {
		const unit = text.charCodeAt(index)

		out[2 + index * 2] = unit >>> 8
		out[3 + index * 2] = unit & 0xff
	}

	return out
}

function isPrintableAscii(text: string): boolean {
	for (let index = 0; index < text.length; index++) {
		const unit = text.charCodeAt(index)

		if (unit < 0x20 || unit > 0x7e) return false
	}

	return true
}

// §7.3.5: regular characters as is, everything else (delimiters, whitespace, '#', non-ASCII as UTF-8) as #xx.
export function writeName(out: ByteWriter, name: string): void {
	out.byte(0x2f)

	const bytes = new TextEncoderLike().encode(name)

	for (const byte of bytes) {
		if (byte > 0x20 && byte < 0x7f && !DELIMITERS.has(byte)) out.byte(byte)
		else out.byte(0x23).ascii(HEX2[byte] ?? "00")
	}
}

const DELIMITERS = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25, 0x23])
const HEX2 = Array.from({ length: 256 }, (_, byte) => byte.toString(16).toUpperCase().padStart(2, "0"))

// UTF-8 without TextEncoder (Hermes has one, but a name is short and this keeps the module dependency-free).
class TextEncoderLike {
	encode(text: string): number[] {
		const out: number[] = []

		for (const char of text) {
			const code = char.codePointAt(0) ?? 0

			if (code < 0x80) out.push(code)
			else if (code < 0x800) out.push(0xc0 | (code >> 6), 0x80 | (code & 63))
			else if (code < 0x10000) out.push(0xe0 | (code >> 12), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63))
			else out.push(0xf0 | (code >> 18), 0x80 | ((code >> 12) & 63), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63))
		}

		return out
	}
}

export function utf8(text: string): Uint8Array {
	return Uint8Array.from(new TextEncoderLike().encode(text))
}

// §7.3.4.2 literal string of printable ASCII, escaping only what must be.
function writeLiteral(out: ByteWriter, bytes: Uint8Array): void {
	out.byte(0x28)

	for (const byte of bytes) {
		if (byte === 0x28 || byte === 0x29 || byte === 0x5c) out.byte(0x5c)

		out.byte(byte)
	}

	out.byte(0x29)
}

export function writeValue(out: ByteWriter, value: PdfValue): void {
	if (value === null) {
		out.ascii("null")
	} else if (typeof value === "boolean") {
		out.ascii(value ? "true" : "false")
	} else if (typeof value === "number") {
		if (Number.isInteger(value)) out.int(value)
		else out.num(value, 5)
	} else if (typeof value === "string") {
		if (isPrintableAscii(value)) {
			writeLiteral(out, textStringBytes(value))
		} else {
			out.byte(0x3c)
			writeHex(out, textStringBytes(value))
			out.byte(0x3e)
		}
	} else if (value instanceof PdfName) {
		writeName(out, value.name)
	} else if (value instanceof PdfRef) {
		out.int(value.id).ascii(" 0 R")
	} else if (value instanceof PdfString) {
		out.byte(0x3c)
		writeHex(out, value.bytes)
		out.byte(0x3e)
	} else if (value instanceof PdfRaw) {
		out.bytes(value.bytes)
	} else if (isArray(value)) {
		out.byte(0x5b)

		let first = true

		for (const item of value) {
			if (!first) out.byte(0x20)

			writeValue(out, item)
			first = false
		}

		out.byte(0x5d)
	} else {
		out.ascii("<<")

		let first = true

		for (const key in value) {
			const item = value[key]

			if (item === undefined) continue
			if (!first) out.byte(0x20)

			first = false
			writeName(out, key)
			out.byte(0x20)
			writeValue(out, item)
		}

		out.ascii(">>")
	}
}

function isArray(value: PdfValue): value is readonly PdfValue[] {
	return Array.isArray(value)
}

export function serialize(value: PdfValue): Uint8Array {
	const out = new ByteWriter(64)

	writeValue(out, value)

	return out.take()
}

// A copy of `value` with every string (text or byte) passed through `transform`, for encrypting the strings
// of an object outside an object stream (§7.6.2).
export async function mapStrings(value: PdfValue, transform: (bytes: Uint8Array) => Promise<Uint8Array>): Promise<PdfValue> {
	if (typeof value === "string") return new PdfString(await transform(textStringBytes(value)))
	if (value instanceof PdfString) return new PdfString(await transform(value.bytes))
	if (value === null || typeof value !== "object" || value instanceof PdfName || value instanceof PdfRef || value instanceof PdfRaw) {
		return value
	}

	if (isArray(value)) {
		const out: PdfValue[] = []

		for (const item of value) out.push(await mapStrings(item, transform))

		return out
	}

	const out: Record<string, PdfValue> = {}

	for (const key in value) {
		const item = value[key]

		if (item !== undefined) out[key] = await mapStrings(item, transform)
	}

	return out
}
