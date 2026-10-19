import { concatBytes } from "./bytes"
import { OfficeCryptoError, unsupported } from "./errors"

// The platform's WebCrypto, typed with the minimal shape used here: the module type-checks without DOM or
// Node typings, and runs in a page, a worker or a WebView alike (Hermes has none: every call reports
// "unsupported" there).

export interface CryptoKeyHandle {
	readonly type: string
}

export type HashName = "SHA-1" | "SHA-256" | "SHA-384" | "SHA-512"

interface AesCbcParams {
	readonly name: "AES-CBC"
	readonly iv: Uint8Array
}

interface SubtleLike {
	digest(algorithm: HashName, data: Uint8Array): Promise<ArrayBuffer>
	importKey(
		format: "raw",
		keyData: Uint8Array,
		algorithm: "AES-CBC" | { readonly name: "HMAC"; readonly hash: HashName },
		extractable: false,
		usages: readonly ("encrypt" | "decrypt" | "sign" | "verify")[]
	): Promise<CryptoKeyHandle>
	encrypt(algorithm: AesCbcParams, key: CryptoKeyHandle, data: Uint8Array): Promise<ArrayBuffer>
	decrypt(algorithm: AesCbcParams, key: CryptoKeyHandle, data: Uint8Array): Promise<ArrayBuffer>
	sign(algorithm: "HMAC", key: CryptoKeyHandle, data: Uint8Array): Promise<ArrayBuffer>
	verify(algorithm: "HMAC", key: CryptoKeyHandle, signature: Uint8Array, data: Uint8Array): Promise<boolean>
}

interface CryptoLike {
	readonly subtle: SubtleLike
	getRandomValues(array: Uint8Array): Uint8Array
}

let cached: CryptoLike | undefined

function platform(): CryptoLike {
	if (cached !== undefined) return cached

	const crypto: unknown = Reflect.get(globalThis, "crypto")
	const subtle: unknown = typeof crypto === "object" && crypto !== null ? Reflect.get(crypto, "subtle") : undefined

	// A WebView outside a secure context has crypto.getRandomValues but no subtle.
	if (typeof subtle !== "object" || subtle === null || typeof Reflect.get(crypto as object, "getRandomValues") !== "function") {
		throw unsupported("WebCrypto is not available here")
	}

	cached = crypto as CryptoLike

	return cached
}

export function subtle(): SubtleLike {
	return platform().subtle
}

export function randomBytes(length: number): Uint8Array {
	const out = new Uint8Array(length)
	const crypto = platform()

	// getRandomValues fills at most 65,536 bytes per call.
	for (let offset = 0; offset < length; offset += 65536) crypto.getRandomValues(out.subarray(offset, Math.min(length, offset + 65536)))

	return out
}

// A concatenation of the parts is zero-filled once hashed: it may hold key material (Hn + blockKey).
export async function digest(hash: HashName, ...parts: readonly Uint8Array[]): Promise<Uint8Array> {
	if (parts.length === 1 && parts[0] !== undefined) return new Uint8Array(await subtle().digest(hash, parts[0]))

	const joined = concatBytes(...parts)

	try {
		return new Uint8Array(await subtle().digest(hash, joined))
	} finally {
		joined.fill(0)
	}
}

// An AES key for both directions. A key size the platform's AES lacks (Chromium has no AES-192) is unsupported.
export async function aesKey(raw: Uint8Array): Promise<CryptoKeyHandle> {
	try {
		return await subtle().importKey("raw", raw, "AES-CBC", false, ["encrypt", "decrypt"])
	} catch (error) {
		if (error instanceof OfficeCryptoError) throw error

		throw unsupported(`AES-${raw.length * 8} is not available here`)
	}
}

export async function hmacKey(hash: HashName, raw: Uint8Array): Promise<CryptoKeyHandle> {
	return subtle().importKey("raw", raw, { name: "HMAC", hash }, false, ["sign", "verify"])
}

const BLOCK = 16
const FULL_PAD = new Uint8Array(BLOCK).fill(BLOCK)

// AES-CBC decryption of whole blocks without padding. WebCrypto always strips PKCS#7 padding, so one block is
// appended that decrypts to a full padding block: E(0x10…10 XOR lastCipherBlock), the first block of a CBC
// encryption of 0x10…10 chained on the last ciphertext block.
export async function cbcDecryptRaw(key: CryptoKeyHandle, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
	if (data.length === 0) return new Uint8Array(0)
	if (data.length % BLOCK !== 0) throw new OfficeCryptoError("corrupt", "Ciphertext is not a whole number of blocks")

	const crafted = new Uint8Array(await subtle().encrypt({ name: "AES-CBC", iv: data.subarray(data.length - BLOCK) }, key, FULL_PAD))
	const input = new Uint8Array(data.length + BLOCK)

	input.set(data)
	input.set(crafted.subarray(0, BLOCK), data.length)

	return new Uint8Array(await subtle().decrypt({ name: "AES-CBC", iv }, key, input))
}

// AES-CBC encryption of whole blocks without padding: the padding block WebCrypto appends is dropped.
export async function cbcEncryptRaw(key: CryptoKeyHandle, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
	if (data.length % BLOCK !== 0) throw new RangeError("Plaintext is not a whole number of blocks")

	return new Uint8Array(await subtle().encrypt({ name: "AES-CBC", iv }, key, data), 0, data.length)
}

const ZERO_IV = new Uint8Array(BLOCK)

// AES-ECB decryption (Standard encryption) without an ECB mode: CBC with a zero IV gives D(C[i]) XOR C[i-1],
// so XOR each block but the first with the ciphertext block before it.
export async function ecbDecrypt(key: CryptoKeyHandle, data: Uint8Array): Promise<Uint8Array> {
	const out = await cbcDecryptRaw(key, ZERO_IV, data)

	for (let k = BLOCK; k < out.length; k++) out[k] = (out[k] ?? 0) ^ (data[k - BLOCK] ?? 0)

	return out
}
