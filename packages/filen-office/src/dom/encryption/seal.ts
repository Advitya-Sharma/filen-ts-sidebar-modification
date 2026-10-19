import { CfbError, CfbFile } from "../../core/cfb/index"
import { isEncryptedContainer, isSealedOfficeFile, readEncryptedContainer, sealedContainer } from "../../core/encryption/index"
import { bytesSource } from "../../core/opc/index"
import { padToBlock, uint32LE } from "./bytes"
import { classifyDataSpaces } from "./dataSpaces"
import { OfficeCryptoError, corrupt } from "./errors"
import { buildAgileInfo, parseEncryptionInfo, type AgileCipher } from "./info"
import { cbcEncryptRaw, hmacKey, randomBytes, subtle, type CryptoKeyHandle } from "./subtle"
import { turn, turnsOf, type OfficeCryptoOptions, type Turns } from "./turns"
import {
	BLOCK,
	BLOCK_HMAC_KEY,
	BLOCK_HMAC_VALUE,
	SEGMENT,
	WINDOW_SEGMENTS,
	blockIv,
	decryptAgile,
	packageSize,
	releaseOfficeKey,
	sealKeys,
	unwrapSecret,
	verifyIntegrity,
	type OfficeKey,
	type SealKeys
} from "./webCrypto"

// Sealing a package again ([MS-OFFCRYPTO] §2.3.4.10-§2.3.4.15), always as Agile under the session's key, and
// the post-seal check. Only a save needs any of it (an editor, a workbook's worker): a viewer's chain never
// imports this module.

// Fewer salt bytes than this are never written, whatever the source declared.
const MIN_SALT = 16

export interface OfficeSealed {
	readonly encryptionInfo: Uint8Array
	// The whole EncryptedPackage stream: StreamSize, then the segments.
	readonly encryptedPackage: Uint8Array
}

// Seals a package as Agile under the session's intermediate key: new keyData salt (so new segment IVs), new
// HMAC key and value, the key encryptors kept. A Standard source's upgraded encryptor is derived at the first
// seal (its spin on `options`). Throws OfficeCryptoError("aborted") once released or aborted.
export async function sealOfficePackage(key: OfficeKey, plaintext: Uint8Array, options: OfficeCryptoOptions = {}): Promise<OfficeSealed> {
	const turns = turnsOf(options)
	const { secret, template } = await sealKeys(key, options)
	const saltSize = Math.max(MIN_SALT, template.keyData.saltSize)
	const keyData: AgileCipher = { ...template.keyData, saltSize, salt: randomBytes(saltSize) }
	const encryptedPackage = await encryptAgile(secret, keyData, plaintext, turns)
	const [hmacKeyIv, hmacValueIv] = await Promise.all([blockIv(keyData, BLOCK_HMAC_KEY), blockIv(keyData, BLOCK_HMAC_VALUE)])
	// hashSize bytes, as Word writes; a reader taking saltSize bytes of a zero-padded key gets the same HMAC.
	const hmacRaw = randomBytes(keyData.hashSize)
	const hmacPadded = padToBlock(hmacRaw, BLOCK)
	let encryptedHmacKey: Uint8Array
	let mac: CryptoKeyHandle

	try {
		encryptedHmacKey = await cbcEncryptRaw(secret, hmacKeyIv, hmacPadded)
		mac = await hmacKey(keyData.hash, hmacRaw)
	} finally {
		hmacRaw.fill(0)
		hmacPadded.fill(0)
	}

	// One platform call over the whole stream (about 0.8 ms per MB in Chromium, inside the calling task).
	await turn(turns)

	const hmacValue = new Uint8Array(await subtle().sign("HMAC", mac, encryptedPackage))

	await turn(turns)
	const encryptedHmacValue = await cbcEncryptRaw(secret, hmacValueIv, padToBlock(hmacValue, BLOCK))
	const encryptionInfo = buildAgileInfo({
		keyData,
		dataIntegrity: { encryptedHmacKey, encryptedHmacValue },
		password: template.password,
		certificates: template.certificates
	})

	return { encryptionInfo, encryptedPackage }
}

// The key a sealed package was encrypted under, read from the package as any reader would: an encryptor this
// module wrote is unwrapped with its derived keys, a source encryptor kept was verified at unlock.
async function sealedSecret(
	keys: SealKeys,
	encryptionInfo: Uint8Array
): Promise<{ secret: CryptoKeyHandle; info: ReturnType<typeof agileOf> }> {
	const info = agileOf(encryptionInfo)
	const secret =
		keys.template.check === undefined ? keys.secret : (await unwrapSecret(info.password, keys.template.check, info.keyData)).secret

	return { secret, info }
}

function agileOf(encryptionInfo: Uint8Array) {
	const info = parseEncryptionInfo(encryptionInfo)

	if (info.kind !== "agile" || info.dataIntegrity === undefined) throw corrupt("A sealed package is Agile with data integrity")

	return { ...info, dataIntegrity: info.dataIntegrity }
}

// Opens a package this session sealed, without the password.
export async function openSealedPackage(
	key: OfficeKey,
	encryptionInfo: Uint8Array,
	encryptedPackage: Uint8Array,
	options: OfficeCryptoOptions = {}
): Promise<Uint8Array> {
	const turns = turnsOf(options)
	const { secret, info } = await sealedSecret(await sealKeys(key, options), encryptionInfo)

	await verifyIntegrity(secret, info.keyData, info.dataIntegrity, encryptedPackage, turns)

	const size = packageSize(encryptedPackage)
	const out = new Uint8Array(size)

	await decryptAgile(secret, info.keyData, encryptedPackage, size, turns, (start, plain) => out.set(plain, start))

	return out
}

// The post-seal self-check without a copy of the package: what openSealedPackage verifies, each decrypted
// window compared with `expected` as it comes. False when the package decrypts to anything else.
export async function sealedPackageEquals(
	key: OfficeKey,
	encryptionInfo: Uint8Array,
	encryptedPackage: Uint8Array,
	expected: Uint8Array,
	options: OfficeCryptoOptions = {}
): Promise<boolean> {
	const turns = turnsOf(options)
	const { secret, info } = await sealedSecret(await sealKeys(key, options), encryptionInfo)

	await verifyIntegrity(secret, info.keyData, info.dataIntegrity, encryptedPackage, turns)

	const size = packageSize(encryptedPackage)
	let equal = size === expected.length

	if (equal)
		await decryptAgile(secret, info.keyData, encryptedPackage, size, turns, (start, plain) => {
			for (let k = 0; k < plain.length && equal; k++) if (plain[k] !== expected[start + k]) equal = false
		})

	return equal
}

// One CBC encryption per segment (each restarts at its own IV), a window of them at a time; the last segment
// is zero-padded to the block. The output is the EncryptedPackage stream itself, so the HMAC needs no copy.
async function encryptAgile(secret: CryptoKeyHandle, keyData: AgileCipher, plaintext: Uint8Array, turns: Turns): Promise<Uint8Array> {
	const size = plaintext.length
	const out = new Uint8Array(8 + Math.ceil(size / BLOCK) * BLOCK)
	const segments = Math.ceil(size / SEGMENT)

	new DataView(out.buffer).setBigUint64(0, BigInt(size), true)

	const encryptSegment = async (index: number): Promise<void> => {
		const start = index * SEGMENT
		const chunk = padToBlock(plaintext.subarray(start, Math.min(size, start + SEGMENT)), BLOCK)
		const cipher = await cbcEncryptRaw(secret, await blockIv(keyData, uint32LE(index)), chunk)

		out.set(cipher, 8 + start)
	}

	for (let first = 0; first < segments; first += WINDOW_SEGMENTS) {
		await turn(turns)

		const batch: Promise<void>[] = []

		for (let index = first; index < Math.min(segments, first + WINDOW_SEGMENTS); index++) batch.push(encryptSegment(index))

		await Promise.all(batch)
	}

	return out
}

// --- The sealer a save goes through ---

export type PackageSealOptions = OfficeCryptoOptions

export interface PackageSealer {
	// The package as a sealed compound file: `zip` encrypted as Agile under the session's key, then the post-seal
	// check (the file read back as a compound file, its package decrypted with the key and compared with `zip`)
	// before it is returned. Rejects with a PackageSealError, or the error `options.slicing`'s pause threw.
	seal(zip: Uint8Array, options?: PackageSealOptions): Promise<Uint8Array>
	readonly released: boolean
}

export type PackageSealFailure = "released" | "aborted" | "check"

export class PackageSealError extends Error {
	readonly reason: PackageSealFailure

	constructor(reason: PackageSealFailure, message: string) {
		super(message)

		this.name = "PackageSealError"
		this.reason = reason
	}
}

// The sealed compound file. Its own frame holds the EncryptedPackage stream, so that copy is garbage once the
// compound file holding it is built: it never coexists with the check.
async function sealedFile(key: OfficeKey, zip: Uint8Array, options: PackageSealOptions): Promise<Uint8Array> {
	const sealed = await sealOfficePackage(key, zip, options)

	return sealedContainer(sealed.encryptionInfo, sealed.encryptedPackage)
}

// The post-seal check reads what will be stored, as any reader would: the compound file, its DataSpaces, the
// package (a view of the file) with its HMAC verified, decrypted a window at a time and compared with what was
// sealed: no copy of the package is made.
async function checkSealed(key: OfficeKey, file: Uint8Array, zip: Uint8Array, options: PackageSealOptions): Promise<void> {
	if (!isSealedOfficeFile(file)) throw new PackageSealError("check", "The sealed file is not a compound file")

	const cfb = await CfbFile.open(bytesSource(file))
	const back = isEncryptedContainer(cfb) ? await readEncryptedContainer(cfb) : undefined

	if (back?.encryptionInfo === undefined || classifyDataSpaces(back.dataSpaces) !== "password")
		throw new PackageSealError("check", "The sealed file does not read back as an encrypted package")

	const stream = await back.encryptedPackage.read(0, back.encryptedPackage.size)

	if (!(await sealedPackageEquals(key, back.encryptionInfo, stream, zip, options)))
		throw new PackageSealError("check", "The sealed file does not decrypt to the package")
}

// The sealer of a key unlocked on this thread (or imported here); release() drops the key.
export function packageSealer(key: OfficeKey): { readonly sealer: PackageSealer; release(): void } {
	let released = false

	const sealer: PackageSealer = {
		get released() {
			return released
		},
		async seal(zip, options = {}) {
			if (released) throw new PackageSealError("released", "The package's key was released")

			try {
				const file = await sealedFile(key, zip, options)

				// The compound file was copied together in one go: a turn before the check reads it back.
				await turn(turnsOf(options))
				await checkSealed(key, file, zip, options)

				return file
			} catch (error) {
				if (error instanceof OfficeCryptoError)
					throw error.code === "aborted"
						? new PackageSealError(released ? "released" : "aborted", error.message)
						: new PackageSealError("check", error.message)
				if (error instanceof CfbError) throw new PackageSealError("check", error.message)

				throw error
			}
		}
	}

	return {
		sealer,
		release() {
			if (released) return

			released = true
			releaseOfficeKey(key)
		}
	}
}

export { isSealedOfficeFile }
