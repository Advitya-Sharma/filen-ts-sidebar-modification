import { corrupt } from "./errors"

const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
const DECODE = new Int16Array(128).fill(-1)

for (let k = 0; k < ALPHABET.length; k++) DECODE[ALPHABET.charCodeAt(k)] = k

// xsd:base64Binary: whitespace is skipped, padding must be canonical, anything else is refused.
export function decodeBase64(text: string): Uint8Array {
	const values = new Uint8Array(text.length)
	let count = 0
	let padding = 0

	for (let k = 0; k < text.length; k++) {
		const code = text.charCodeAt(k)

		if (code === 32 || code === 9 || code === 10 || code === 13) continue

		if (code === 61) {
			padding++

			continue
		}

		const value = code < 128 ? (DECODE[code] ?? -1) : -1

		if (value < 0 || padding > 0) throw corrupt("Invalid base64 value")

		values[count++] = value
	}

	if (padding > 2 || (count + padding) % 4 !== 0 || (padding > 0 && count % 4 === 0)) throw corrupt("Invalid base64 length")

	const out = new Uint8Array(Math.floor((count * 3) / 4))
	let o = 0

	for (let k = 0; k < count; k += 4) {
		const a = values[k] ?? 0
		const b = values[k + 1] ?? 0
		const c = values[k + 2] ?? 0
		const d = values[k + 3] ?? 0

		out[o++] = (a << 2) | (b >> 4)

		if (k + 2 < count) out[o++] = ((b & 15) << 4) | (c >> 2)
		if (k + 3 < count) out[o++] = ((c & 3) << 6) | d
	}

	return out
}

export function encodeBase64(bytes: Uint8Array): string {
	let out = ""

	for (let k = 0; k < bytes.length; k += 3) {
		const a = bytes[k] ?? 0
		const b = bytes[k + 1] ?? 0
		const c = bytes[k + 2] ?? 0

		out += ALPHABET[a >> 2]
		out += ALPHABET[((a & 3) << 4) | (b >> 4)]
		out += k + 1 < bytes.length ? ALPHABET[((b & 15) << 2) | (c >> 6)] : "="
		out += k + 2 < bytes.length ? ALPHABET[c & 63] : "="
	}

	return out
}

export function concatBytes(...parts: readonly Uint8Array[]): Uint8Array {
	let length = 0

	for (const part of parts) length += part.length

	const out = new Uint8Array(length)
	let offset = 0

	for (const part of parts) {
		out.set(part, offset)
		offset += part.length
	}

	return out
}

export function uint32LE(value: number): Uint8Array {
	const out = new Uint8Array(4)

	new DataView(out.buffer).setUint32(0, value, true)

	return out
}

// The password as UTF-16LE code units, unnormalised ([MS-OFFCRYPTO] §2.3.4.7, §2.3.4.11).
export function utf16LE(text: string): Uint8Array {
	const out = new Uint8Array(text.length * 2)

	for (let k = 0; k < text.length; k++) {
		const unit = text.charCodeAt(k)

		out[2 * k] = unit & 255
		out[2 * k + 1] = unit >> 8
	}

	return out
}

// `bytes` cut or padded to `length` (§2.3.4.11 keys and §2.3.4.12 IVs pad with 0x36).
export function fitBytes(bytes: Uint8Array, length: number, pad: number = 0x36): Uint8Array {
	if (bytes.length >= length) return bytes.subarray(0, length)

	const out = new Uint8Array(length).fill(pad)

	out.set(bytes)

	return out
}

// `bytes` zero-padded to a multiple of `block` (a copy only when padding is needed).
export function padToBlock(bytes: Uint8Array, block: number): Uint8Array {
	const rest = bytes.length % block

	if (rest === 0) return bytes

	const out = new Uint8Array(bytes.length + block - rest)

	out.set(bytes)

	return out
}

// Compares without an early exit, so the time taken says nothing about where two values differ.
export function constantTimeEqual(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false

	let difference = 0

	for (let k = 0; k < a.length; k++) difference |= (a[k] ?? 0) ^ (b[k] ?? 0)

	return difference === 0
}
