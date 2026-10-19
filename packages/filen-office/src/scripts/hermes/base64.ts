// Base64 without atob/btoa or Buffer, so the Hermes entry and the Node runner share one codec.

const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
const VALUES = new Int16Array(128).fill(-1)

for (let i = 0; i < ALPHABET.length; i++) VALUES[ALPHABET.charCodeAt(i)] = i

export function encodeBase64(bytes: Uint8Array): string {
	const parts: string[] = []
	let chunk = ""

	for (let i = 0; i < bytes.length; i += 3) {
		const a = bytes[i] ?? 0
		const b = bytes[i + 1] ?? 0
		const c = bytes[i + 2] ?? 0
		const word = (a << 16) | (b << 8) | c

		chunk +=
			ALPHABET.charAt(word >>> 18) +
			ALPHABET.charAt((word >>> 12) & 63) +
			(i + 1 < bytes.length ? ALPHABET.charAt((word >>> 6) & 63) : "=") +
			(i + 2 < bytes.length ? ALPHABET.charAt(word & 63) : "=")

		if (chunk.length >= 4096) {
			parts.push(chunk)
			chunk = ""
		}
	}

	parts.push(chunk)

	return parts.join("")
}

export function decodeBase64(text: string): Uint8Array {
	const clean = text.replace(/[^A-Za-z0-9+/]/g, "")
	const out = new Uint8Array(Math.floor((clean.length * 3) / 4))
	let at = 0

	for (let i = 0; i < clean.length; i += 4) {
		const word =
			((VALUES[clean.charCodeAt(i)] ?? 0) << 18) |
			((VALUES[clean.charCodeAt(i + 1)] ?? 0) << 12) |
			(Math.max(0, VALUES[clean.charCodeAt(i + 2)] ?? 0) << 6) |
			Math.max(0, VALUES[clean.charCodeAt(i + 3)] ?? 0)

		if (at < out.length) out[at++] = word >>> 16
		if (at < out.length) out[at++] = (word >>> 8) & 255
		if (at < out.length) out[at++] = word & 255
	}

	return out
}
