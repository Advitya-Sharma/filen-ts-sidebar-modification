// MD5 (RFC 1321) and RC4, for reading Word 97–2003 files encrypted with [MS-OFFCRYPTO] §2.3.5/§2.3.6 RC4. WebCrypto
// has neither. Decrypt-only: nothing in Filen writes these ciphers; they exist to read files that already use them.
// Plain table-free implementations over typed arrays; no allocation per block. Callers wipe what they hand in.

// --- MD5, RFC 1321 §3 ---

// T[i] = floor(2^32 * abs(sin(i + 1))), §3.4.
const T = new Int32Array(64)

for (let i = 0; i < 64; i++) T[i] = Math.floor(Math.abs(Math.sin(i + 1)) * 0x100000000) | 0

// Per-round shift amounts, §3.4.
const S = Uint8Array.of(7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21)

function md5Block(state: Int32Array, x: Int32Array): void {
	let a = state[0] ?? 0
	let b = state[1] ?? 0
	let c = state[2] ?? 0
	let d = state[3] ?? 0

	for (let i = 0; i < 64; i++) {
		const round = i >> 4
		let f: number
		let g: number

		if (round === 0) {
			f = (b & c) | (~b & d)
			g = i
		} else if (round === 1) {
			f = (b & d) | (c & ~d)
			g = (5 * i + 1) & 15
		} else if (round === 2) {
			f = b ^ c ^ d
			g = (3 * i + 5) & 15
		} else {
			f = c ^ (b | ~d)
			g = (7 * i) & 15
		}

		const s = S[(round << 2) | (i & 3)] ?? 0
		const sum = (a + f + (T[i] ?? 0) + (x[g] ?? 0)) | 0
		const next = (b + ((sum << s) | (sum >>> (32 - s)))) | 0

		a = d
		d = c
		c = b
		b = next
	}

	state[0] = ((state[0] ?? 0) + a) | 0
	state[1] = ((state[1] ?? 0) + b) | 0
	state[2] = ((state[2] ?? 0) + c) | 0
	state[3] = ((state[3] ?? 0) + d) | 0
}

// The MD5 digest of `parts` concatenated (16 bytes). Inputs here are a few hundred bytes (a password, a salt);
// the working block is wiped before returning.
export function md5(...parts: readonly Uint8Array[]): Uint8Array {
	let length = 0

	for (const part of parts) length += part.length

	const state = Int32Array.of(0x67452301, 0xefcdab89 | 0, 0x98badcfe | 0, 0x10325476)
	const block = new Uint8Array(64)
	const words = new Int32Array(16)
	let fill = 0
	const flush = (): void => {
		for (let w = 0; w < 16; w++) {
			const at = w << 2

			words[w] = (block[at] ?? 0) | ((block[at + 1] ?? 0) << 8) | ((block[at + 2] ?? 0) << 16) | ((block[at + 3] ?? 0) << 24)
		}

		md5Block(state, words)
		fill = 0
	}
	const push = (byte: number): void => {
		block[fill++] = byte

		if (fill === 64) flush()
	}

	for (const part of parts) for (let k = 0; k < part.length; k++) push(part[k] ?? 0)

	// §3.1–§3.2: a 1 bit, zeros to 56 mod 64, the bit length as a 64-bit little-endian count.
	push(0x80)

	while (fill !== 56) push(0)

	const bitsLow = (length * 8) >>> 0
	const bitsHigh = Math.floor(length / 0x20000000) >>> 0

	for (let k = 0; k < 4; k++) push((bitsLow >>> (8 * k)) & 0xff)
	for (let k = 0; k < 4; k++) push((bitsHigh >>> (8 * k)) & 0xff)

	const out = new Uint8Array(16)

	for (let k = 0; k < 16; k++) out[k] = ((state[k >> 2] ?? 0) >>> (8 * (k & 3))) & 0xff

	block.fill(0)
	words.fill(0)
	state.fill(0)

	return out
}

// --- RC4 ---

// RC4 keyed once; apply() continues the key stream across calls. wipe() clears its state.
export class Rc4 {
	private readonly s = new Uint8Array(256)
	private i = 0
	private j = 0

	// Key of 1 to 256 bytes.
	constructor(key: Uint8Array) {
		if (key.length < 1 || key.length > 256) throw new RangeError("RC4 key must be 1 to 256 bytes")

		const s = this.s

		for (let k = 0; k < 256; k++) s[k] = k

		for (let k = 0, j = 0; k < 256; k++) {
			const value = s[k] ?? 0

			j = (j + value + (key[k % key.length] ?? 0)) & 0xff
			s[k] = s[j] ?? 0
			s[j] = value
		}
	}

	// XORs `length` bytes of `input` from `from` with the key stream into `output` at `to` (may be the same array).
	apply(input: Uint8Array, output: Uint8Array, from = 0, to = from, length = input.length - from): void {
		const s = this.s
		let i = this.i
		let j = this.j

		for (let k = 0; k < length; k++) {
			i = (i + 1) & 0xff

			const si = s[i] ?? 0

			j = (j + si) & 0xff

			const sj = s[j] ?? 0

			s[i] = sj
			s[j] = si
			output[to + k] = (input[from + k] ?? 0) ^ (s[(si + sj) & 0xff] ?? 0)
		}

		this.i = i
		this.j = j
	}

	wipe(): void {
		this.s.fill(0)
		this.i = 0
		this.j = 0
	}
}

// Equality whose time depends only on the lengths (a password verifier's hash against the stored one).
export function constantTimeEqual(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false

	let difference = 0

	for (let k = 0; k < a.length; k++) difference |= (a[k] ?? 0) ^ (b[k] ?? 0)

	return difference === 0
}
