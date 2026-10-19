import { concatBytes, constantTimeEqual, fitBytes, uint32LE, utf16LE } from "./bytes"
import { OfficeCryptoError, corrupt, type OfficeCryptoErrorCode } from "./errors"
import {
	MAX_SPIN_COUNT,
	parseEncryptionInfo,
	type AgileCipher,
	type CertificateKeyEncryptor,
	type DataIntegrity,
	type PasswordKeyEncryptor,
	type StandardEncryptionInfo
} from "./info"
import {
	aesKey,
	cbcDecryptRaw,
	cbcEncryptRaw,
	digest,
	ecbDecrypt,
	hmacKey,
	randomBytes,
	subtle,
	type CryptoKeyHandle,
	type HashName
} from "./subtle"
import { raceSignal, turn, turnsOf, type AbortSignalLike, type OfficeCryptoOptions, type Turns } from "./turns"

// [MS-OFFCRYPTO] §2.3.4 password encryption of OOXML packages on the platform's WebCrypto: Standard
// (AES-ECB, SHA-1) and Agile (AES-CBC) to open; seal.ts seals, always as Agile. Long work gives the host a turn
// every few milliseconds (turns.ts), so a caller's thread keeps handling input; the password spin (100,000
// sequential digests by default: 0.2 s in Chromium, 1.2 s in Firefox, 2.7 s in WebKit) runs in a worker.
//
// The caller holds the streams, never a key: unlock returns an opaque OfficeKey whose intermediate key is a
// non-extractable CryptoKey, which seal reuses (no spin at save for an Agile source). The password is dropped
// once its first hash is taken, and raw key bytes are zero-filled after import. A key moves to another thread
// as an OfficeKeyTransfer (CryptoKeys are structured-cloneable; their bytes stay unreadable).

// Excel's key for workbooks protected without an open password ("VelvetSweatshop"): try it before prompting.
export const DEFAULT_OFFICE_PASSWORD = "VelvetSweatshop"

// Word's defaults (§2.3.4.10), written when a Standard source is sealed as Agile.
export const AGILE_DEFAULT_SPIN_COUNT = 100_000
const UPGRADE_HASH: HashName = "SHA-512"
const UPGRADE_KEY_BITS = 256
const UPGRADE_SALT_SIZE = 16

const STANDARD_SPIN_COUNT = 50_000
export const SEGMENT = 4096
export const BLOCK = 16
// Segments per WebCrypto call when decrypting, and per batch of concurrent calls when encrypting: 1 MiB.
export const WINDOW_SEGMENTS = 256
// Spin iterations between looks at the clock and the abort signal.
const SPIN_CHECK = 64

const BLOCK_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79])
const BLOCK_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e])
const BLOCK_KEY_VALUE = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6])
export const BLOCK_HMAC_KEY = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6])
export const BLOCK_HMAC_VALUE = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33])

export type { AbortSignalLike, OfficeCryptoOptions }

export interface OfficeUnlockOptions extends OfficeCryptoOptions {
	// Spin count of the password key encryptor written when a Standard source is sealed as Agile.
	readonly upgradeSpinCount?: number
}

// Opaque: its keys live in this module, for this thread. Hand it back to seal, export and release.
export interface OfficeKey {
	readonly scheme: "standard" | "agile"
}

export type OfficeUnlockResult =
	| {
			readonly ok: true
			readonly key: OfficeKey
			// The decrypted package (StreamSize bytes): a zip.
			readonly plaintext: Uint8Array
			// verified: the Agile HMAC matched; unchecked: Standard has none, or the Agile source omitted it.
			readonly integrity: "verified" | "unchecked"
	  }
	| {
			readonly ok: false
			readonly reason: OfficeCryptoErrorCode
			readonly message: string
	  }

// The three password keys of §2.3.4.13 for an encryptor this module wrote.
export interface PasswordKeys {
	readonly input: CryptoKeyHandle
	readonly value: CryptoKeyHandle
	readonly key: CryptoKeyHandle
}

// What a seal writes besides fresh salts and the HMAC: the keyData parameters and the key encryptors.
export interface SealTemplate {
	readonly keyData: Omit<AgileCipher, "salt">
	readonly password: PasswordKeyEncryptor
	readonly certificates: readonly CertificateKeyEncryptor[]
	// For an encryptor this module wrote: its password-derived keys, so the post-seal check verifies it end to end.
	readonly check: PasswordKeys | undefined
}

// The intermediate key and the template: everything a seal needs.
export interface SealKeys {
	readonly secret: CryptoKeyHandle
	readonly template: SealTemplate
}

// A key moved to another thread (structured clone): importOfficeKey registers it there.
export interface OfficeKeyTransfer extends SealKeys {
	readonly scheme: OfficeKey["scheme"]
}

// A Standard source's upgrade to Agile, not derived yet: H0 of the new encryptor's salt and the password. The
// spin, the new intermediate key and the encryptor come with the first seal or export.
interface Upgrade {
	readonly h0: Uint8Array
	readonly salt: Uint8Array
	readonly spinCount: number
}

interface KeyState {
	readonly scheme: OfficeKey["scheme"]
	readonly released: { aborted: boolean }
	keys: SealKeys | undefined
	upgrade: Upgrade | undefined
	pending: Promise<SealKeys> | undefined
}

const states = new WeakMap<OfficeKey, KeyState>()

// Which scheme an EncryptionInfo stream uses, or why it cannot be opened: refuse before prompting.
export function probeEncryptionInfo(encryptionInfo: Uint8Array):
	| {
			readonly ok: true
			readonly scheme: "standard" | "agile"
			readonly spinCount: number
	  }
	| { readonly ok: false; readonly reason: OfficeCryptoErrorCode } {
	try {
		const info = parseEncryptionInfo(encryptionInfo)

		return {
			ok: true,
			scheme: info.kind,
			spinCount: info.kind === "agile" ? info.password.spinCount : STANDARD_SPIN_COUNT
		}
	} catch (error) {
		if (error instanceof OfficeCryptoError) return { ok: false, reason: error.code }

		throw error
	}
}

// Verifies the password, then (Agile) the HMAC over the encrypted package, then decrypts it. Wrong passwords,
// unsupported schemes and damaged streams come back as { ok: false }; nothing is logged.
export async function unlockOfficePackage(
	encryptionInfo: Uint8Array,
	encryptedPackage: Uint8Array,
	password: string,
	options: OfficeUnlockOptions = {}
): Promise<OfficeUnlockResult> {
	const upgradeSpinCount = options.upgradeSpinCount ?? AGILE_DEFAULT_SPIN_COUNT

	if (!Number.isInteger(upgradeSpinCount) || upgradeSpinCount < 0 || upgradeSpinCount > MAX_SPIN_COUNT)
		throw new RangeError("upgradeSpinCount is out of range")

	try {
		const info = parseEncryptionInfo(encryptionInfo)
		const size = packageSize(encryptedPackage)
		const turns = turnsOf(options)

		if (info.kind === "agile") {
			const keys = await passwordKeys(
				info.password,
				await iteratedHash(info.password.hash, info.password.salt, password, info.password.spinCount, turns)
			)
			const { secret, certificates } = await unwrapSecret(info.password, keys, info.keyData, info.certificates)
			let integrity: "verified" | "unchecked" = "unchecked"

			if (info.dataIntegrity !== undefined) {
				await verifyIntegrity(secret, info.keyData, info.dataIntegrity, encryptedPackage, turns)
				integrity = "verified"
			}

			const plaintext = new Uint8Array(size)

			await decryptAgile(secret, info.keyData, encryptedPackage, size, turns, (start, plain) => plaintext.set(plain, start))

			const template: SealTemplate = { keyData: info.keyData, password: info.password, certificates, check: undefined }

			return { ok: true, key: register("agile", { secret, template }, undefined), plaintext, integrity }
		}

		const key = await standardKey(info, password, turns)

		await verifyStandard(key, info)

		const plaintext = await decryptStandard(key, encryptedPackage, size, turns)
		// Only H0 of the Agile encryptor a seal will write is taken now: the password is not kept, and the spin
		// waits for a seal (a session that only views never pays it).
		const salt = randomBytes(UPGRADE_SALT_SIZE)
		const h0 = await firstHash(UPGRADE_HASH, salt, password)

		return {
			ok: true,
			key: register("standard", undefined, { h0, salt, spinCount: upgradeSpinCount }),
			plaintext,
			integrity: "unchecked"
		}
	} catch (error) {
		if (error instanceof OfficeCryptoError) return { ok: false, reason: error.code, message: error.message }

		throw error
	}
}

// Drops the key's state; an upgrade spin under way stops at its next turn.
export function releaseOfficeKey(key: OfficeKey): void {
	const state = states.get(key)

	if (state === undefined) return

	state.released.aborted = true
	state.upgrade?.h0.fill(0)
	state.upgrade = undefined
	states.delete(key)
}

// The key as another thread registers it (importOfficeKey): the intermediate key and the template, the Standard
// upgrade's spin run here first (with the caller's signal and slicing). The bytes inside stay unreadable.
export async function exportOfficeKey(key: OfficeKey, options: OfficeCryptoOptions = {}): Promise<OfficeKeyTransfer> {
	const state = stateOf(key)
	const keys = await sealKeys(key, options)

	return { scheme: state.scheme, secret: keys.secret, template: keys.template }
}

export function importOfficeKey(transfer: OfficeKeyTransfer): OfficeKey {
	return register(transfer.scheme, { secret: transfer.secret, template: transfer.template }, undefined)
}

function register(scheme: OfficeKey["scheme"], keys: SealKeys | undefined, upgrade: Upgrade | undefined): OfficeKey {
	const key: OfficeKey = Object.freeze({ scheme })

	states.set(key, { scheme, released: { aborted: false }, keys, upgrade, pending: undefined })

	return key
}

function stateOf(key: OfficeKey): KeyState {
	const state = states.get(key)

	if (state === undefined) throw new OfficeCryptoError("aborted", "The key was released")

	return state
}

// The intermediate key and the template a seal writes. A Standard source's upgrade starts at the first call, on
// that caller's slicing; each caller waits until it is done or its own signal fires.
export function sealKeys(key: OfficeKey, options: OfficeCryptoOptions = {}): Promise<SealKeys> {
	let state: KeyState

	try {
		state = stateOf(key)
	} catch (error) {
		return Promise.reject(error)
	}

	if (state.keys !== undefined) return Promise.resolve(state.keys)

	const upgrade = state.upgrade

	if (upgrade === undefined) return Promise.reject(new OfficeCryptoError("aborted", "The key was released"))

	// One upgrade for every caller, stopped only by release: a caller that gives up (its signal) stops waiting,
	// not the others, and the spin goes on for the next seal.
	const run = async (): Promise<SealKeys> => {
		try {
			const keys = await upgradeKeys(
				upgrade,
				turnsOf({ signal: state.released, ...(options.slicing === undefined ? {} : { slicing: options.slicing }) })
			)

			if (state.released.aborted) throw new OfficeCryptoError("aborted", "The key was released")

			state.keys = keys
			upgrade.h0.fill(0)
			state.upgrade = undefined

			return keys
		} finally {
			state.pending = undefined
		}
	}

	state.pending ??= run()
	// Nobody may be waiting when it fails (every caller gave up): that is no unhandled rejection.
	state.pending.catch(() => undefined)

	return raceSignal(state.pending, options.signal)
}

// StreamSize (UINT64) and that the ciphertext after it covers that many bytes in whole blocks.
export function packageSize(stream: Uint8Array): number {
	return streamSizeOf(stream.subarray(0, 8), stream.length)
}

// StreamSize from the stream's first 8 bytes, checked against the stream's length.
export function streamSizeOf(head: Uint8Array, streamLength: number): number {
	if (head.length < 8 || streamLength < 8) throw corrupt("EncryptedPackage is truncated")

	const size = new DataView(head.buffer, head.byteOffset, 8).getBigUint64(0, true)

	if (size > BigInt(streamLength - 8)) throw corrupt("EncryptedPackage is shorter than its StreamSize")

	const length = Number(size)

	if (streamLength - 8 < Math.ceil(length / BLOCK) * BLOCK) throw corrupt("EncryptedPackage ends inside a block")

	return length
}

// H0 = H(salt + UTF-16LE(password)) (§2.3.4.7, §2.3.4.11); every copy of the password's bytes is zero-filled.
async function firstHash(hash: HashName, salt: Uint8Array, password: string): Promise<Uint8Array> {
	const units = utf16LE(password)
	const first = concatBytes(salt, units)

	units.fill(0)

	try {
		return await digest(hash, first)
	} finally {
		first.fill(0)
	}
}

// H0, then Hn = H(LE32(i) + Hn-1) for i < spinCount.
async function iteratedHash(hash: HashName, salt: Uint8Array, password: string, spinCount: number, turns: Turns): Promise<Uint8Array> {
	return spin(hash, await firstHash(hash, salt, password), spinCount, turns)
}

// The spin from H0, which it zero-fills. One buffer is reused: digest copies its input before it resolves.
async function spin(hash: HashName, h0: Uint8Array, spinCount: number, turns: Turns): Promise<Uint8Array> {
	const crypto = subtle()
	const buffer = new Uint8Array(4 + h0.length)
	const view = new DataView(buffer.buffer)
	let h = h0

	try {
		for (let i = 0; i < spinCount; i++) {
			view.setUint32(0, i, true)
			buffer.set(h, 4)
			h.fill(0)
			h = new Uint8Array(await crypto.digest(hash, buffer))

			if (i % SPIN_CHECK === SPIN_CHECK - 1) await turn(turns)
		}
	} catch (error) {
		h.fill(0)

		throw error
	} finally {
		buffer.fill(0)
	}

	return h
}

// Hfinal = H(Hn + blockKey), cut or 0x36-padded to keyBits (§2.3.4.11), as an AES key.
async function blockKeyOf(encryptor: Omit<AgileCipher, "salt">, hn: Uint8Array, blockKey: Uint8Array): Promise<CryptoKeyHandle> {
	const full = await digest(encryptor.hash, hn, blockKey)
	const raw = fitBytes(full, encryptor.keyBits / 8)

	try {
		return await aesKey(raw)
	} finally {
		full.fill(0)
		raw.fill(0)
	}
}

// The three password keys of §2.3.4.13. Zero-fills `hn`.
async function passwordKeys(encryptor: Omit<AgileCipher, "salt">, hn: Uint8Array): Promise<PasswordKeys> {
	try {
		const [input, value, key] = await Promise.all([
			blockKeyOf(encryptor, hn, BLOCK_VERIFIER_INPUT),
			blockKeyOf(encryptor, hn, BLOCK_VERIFIER_VALUE),
			blockKeyOf(encryptor, hn, BLOCK_KEY_VALUE)
		])

		return { input, value, key }
	} finally {
		hn.fill(0)
	}
}

// Checks the verifier (H(input[0..saltSize]) equals value[0..hashSize]), then decrypts the intermediate key
// (§2.3.4.13). Each decryption's IV is the encryptor's salt fitted to the block (§2.3.4.12). Certificate key
// encryptors whose certVerifier is not HMAC(intermediate key, X509Certificate) are discarded (§2.3.4.10).
export async function unwrapSecret(
	encryptor: PasswordKeyEncryptor,
	keys: PasswordKeys,
	keyData: Omit<AgileCipher, "salt">,
	certificates: readonly CertificateKeyEncryptor[] = []
): Promise<{ readonly secret: CryptoKeyHandle; readonly certificates: readonly CertificateKeyEncryptor[] }> {
	const iv = fitBytes(encryptor.salt, BLOCK)
	const [input, value] = await Promise.all([
		cbcDecryptRaw(keys.input, iv, encryptor.encryptedVerifierHashInput),
		cbcDecryptRaw(keys.value, iv, encryptor.encryptedVerifierHashValue)
	])
	const expected = await digest(encryptor.hash, input.subarray(0, encryptor.saltSize))
	const matches = constantTimeEqual(expected, value.subarray(0, encryptor.hashSize))

	input.fill(0)
	value.fill(0)
	expected.fill(0)

	if (!matches) throw new OfficeCryptoError("wrongPassword", "The password is not correct")

	const decrypted = await cbcDecryptRaw(keys.key, iv, encryptor.encryptedKeyValue)
	const raw = decrypted.subarray(0, keyData.keyBits / 8)

	try {
		const secret = await aesKey(raw)

		if (certificates.length === 0) return { secret, certificates }

		const mac = await hmacKey(keyData.hash, raw)
		const verified = await Promise.all(
			certificates.map(certificate => subtle().verify("HMAC", mac, certificate.certVerifier, certificate.certificate))
		)

		return { secret, certificates: certificates.filter((_, k) => verified[k] === true) }
	} finally {
		decrypted.fill(0)
	}
}

// IV = H(keyData.saltValue + blockKey), fitted to the block (§2.3.4.12).
export async function blockIv(keyData: AgileCipher, blockKey: Uint8Array): Promise<Uint8Array> {
	return fitBytes(await digest(keyData.hash, keyData.salt, blockKey), BLOCK)
}

// §2.3.4.14: HMAC over the whole EncryptedPackage stream, StreamSize included. The HMAC is one platform call
// over the whole stream (about 0.8 ms per MB in Chromium, inside the calling task): a turn comes first.
export async function verifyIntegrity(
	secret: CryptoKeyHandle,
	keyData: AgileCipher,
	dataIntegrity: DataIntegrity,
	stream: Uint8Array,
	turns: Turns
): Promise<void> {
	const [keyIv, valueIv] = await Promise.all([blockIv(keyData, BLOCK_HMAC_KEY), blockIv(keyData, BLOCK_HMAC_VALUE)])
	const [keyBytes, valueBytes] = await Promise.all([
		cbcDecryptRaw(secret, keyIv, dataIntegrity.encryptedHmacKey),
		cbcDecryptRaw(secret, valueIv, dataIntegrity.encryptedHmacValue)
	])
	let mac: CryptoKeyHandle

	try {
		mac = await hmacKey(keyData.hash, keyBytes.subarray(0, Math.min(keyBytes.length, keyData.hashSize)))
	} finally {
		keyBytes.fill(0)
	}

	await turn(turns)

	if (!(await subtle().verify("HMAC", mac, valueBytes.subarray(0, keyData.hashSize), stream))) {
		throw new OfficeCryptoError("integrity", "The document was changed after it was encrypted")
	}
}

// Segment i is AES-CBC under IV = H(keyData.saltValue + LE32(i)) (§2.3.4.15). A window of segments is one
// CBC decryption, whose chaining is wrong only for each later segment's first block: there it used the
// previous segment's last ciphertext block instead of the segment's IV, so XOR both in. Each window's plaintext
// (up to `size`) goes to `sink` with its offset; the sink may not keep the buffer.
export async function decryptAgile(
	secret: CryptoKeyHandle,
	keyData: AgileCipher,
	stream: Uint8Array,
	size: number,
	turns: Turns,
	sink: (start: number, plain: Uint8Array) => void
): Promise<void> {
	const cipherLength = Math.ceil(size / BLOCK) * BLOCK
	const data = stream.subarray(8, 8 + cipherLength)

	for (let first = 0; first * SEGMENT < cipherLength; first += WINDOW_SEGMENTS) {
		await turn(turns)

		const start = first * SEGMENT
		const window = data.subarray(start, Math.min(cipherLength, start + WINDOW_SEGMENTS * SEGMENT))
		const segments = Math.ceil(window.length / SEGMENT)
		const ivs = await Promise.all(Array.from({ length: segments }, (_, k) => blockIv(keyData, uint32LE(first + k))))
		const plain = await cbcDecryptRaw(secret, ivs[0] ?? new Uint8Array(BLOCK), window)

		for (let k = 1; k < segments; k++) {
			const at = k * SEGMENT
			const iv = ivs[k] ?? new Uint8Array(BLOCK)

			for (let b = 0; b < BLOCK; b++) plain[at + b] = (plain[at + b] ?? 0) ^ (window[at - BLOCK + b] ?? 0) ^ (iv[b] ?? 0)
		}

		sink(start, plain.subarray(0, Math.min(plain.length, size - start)))
	}
}

// §2.3.4.7: Hfinal = SHA-1(Hn + LE32(0)) after 50,000 iterations, then CryptDeriveKey: X1 = SHA-1(0x36 x 64 XOR
// Hfinal), X2 = SHA-1(0x5C x 64 XOR Hfinal), key = the first KeySize/8 bytes of X1 + X2.
async function standardKey(info: StandardEncryptionInfo, password: string, turns: Turns): Promise<CryptoKeyHandle> {
	const hn = await iteratedHash("SHA-1", info.salt, password, STANDARD_SPIN_COUNT, turns)
	const final = await digest("SHA-1", hn, uint32LE(0))
	const inner = new Uint8Array(64).fill(0x36)
	const outer = new Uint8Array(64).fill(0x5c)

	for (let k = 0; k < final.length; k++) {
		inner[k] = (inner[k] ?? 0) ^ (final[k] ?? 0)
		outer[k] = (outer[k] ?? 0) ^ (final[k] ?? 0)
	}

	const [x1, x2] = await Promise.all([digest("SHA-1", inner), digest("SHA-1", outer)])
	const x3 = concatBytes(x1, x2)

	try {
		return await aesKey(x3.subarray(0, info.keyBits / 8))
	} finally {
		for (const bytes of [hn, final, inner, outer, x1, x2, x3]) bytes.fill(0)
	}
}

// §2.3.4.9: SHA-1 of the decrypted verifier equals the first 20 bytes of the decrypted verifier hash.
async function verifyStandard(key: CryptoKeyHandle, info: StandardEncryptionInfo): Promise<void> {
	const [verifier, hash] = await Promise.all([ecbDecrypt(key, info.encryptedVerifier), ecbDecrypt(key, info.encryptedVerifierHash)])
	const expected = await digest("SHA-1", verifier)
	const matches = constantTimeEqual(expected, hash.subarray(0, info.verifierHashSize))

	verifier.fill(0)
	hash.fill(0)
	expected.fill(0)

	if (!matches) throw new OfficeCryptoError("wrongPassword", "The password is not correct")
}

// AES-ECB over the stream after StreamSize, a window at a time (ECB blocks are independent) (§2.3.4.15 does
// not apply: Standard has no segments).
async function decryptStandard(key: CryptoKeyHandle, stream: Uint8Array, size: number, turns: Turns): Promise<Uint8Array> {
	const out = new Uint8Array(size)
	const cipherLength = Math.ceil(size / BLOCK) * BLOCK
	const windowBytes = WINDOW_SEGMENTS * SEGMENT

	for (let start = 0; start < cipherLength; start += windowBytes) {
		await turn(turns)

		const plain = await ecbDecrypt(key, stream.subarray(8 + start, 8 + Math.min(cipherLength, start + windowBytes)))

		out.set(plain.subarray(0, Math.min(plain.length, size - start)), start)
	}

	return out
}

// The Agile password key encryptor that replaces a Standard one (Word's defaults: AES-256, SHA-512) and the new
// intermediate key it wraps, from H0 (copied: an aborted spin may run again). The new key's bytes are zero-filled
// once wrapped, or on failure.
async function upgradeKeys(upgrade: Upgrade, turns: Turns): Promise<SealKeys> {
	const cipher: Omit<AgileCipher, "salt"> = {
		saltSize: UPGRADE_SALT_SIZE,
		blockSize: BLOCK,
		keyBits: UPGRADE_KEY_BITS,
		hashSize: 64,
		hash: UPGRADE_HASH
	}
	const keys = await passwordKeys(cipher, await spin(UPGRADE_HASH, upgrade.h0.slice(), upgrade.spinCount, turns))
	const secretRaw = randomBytes(UPGRADE_KEY_BITS / 8)
	const verifierInput = randomBytes(UPGRADE_SALT_SIZE)

	try {
		const iv = fitBytes(upgrade.salt, BLOCK)
		const verifierHash = await digest(UPGRADE_HASH, verifierInput)
		const [secret, encryptedVerifierHashInput, encryptedVerifierHashValue, encryptedKeyValue] = await Promise.all([
			aesKey(secretRaw),
			cbcEncryptRaw(keys.input, iv, verifierInput),
			cbcEncryptRaw(keys.value, iv, verifierHash),
			cbcEncryptRaw(keys.key, iv, secretRaw)
		])

		verifierHash.fill(0)

		return {
			secret,
			template: {
				keyData: cipher,
				password: {
					...cipher,
					salt: upgrade.salt,
					spinCount: upgrade.spinCount,
					encryptedVerifierHashInput,
					encryptedVerifierHashValue,
					encryptedKeyValue
				},
				certificates: [],
				check: keys
			}
		}
	} finally {
		secretRaw.fill(0)
		verifierInput.fill(0)
	}
}
