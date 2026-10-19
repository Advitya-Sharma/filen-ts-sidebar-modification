import { Namespaces, XmlError, encodeUtf8, parseXml, type XmlDocument } from "../../core/xml/index"
import { decodeBase64, encodeBase64 } from "./bytes"
import { corrupt, unsupported } from "./errors"
import type { HashName } from "./subtle"

// The \EncryptionInfo stream ([MS-OFFCRYPTO] §2.3.4.5 Standard, §2.3.4.10 Agile): parsed into what the
// backend uses, refusing what it does not implement. No crypto here.

export const ENCRYPTION_NAMESPACE = "http://schemas.microsoft.com/office/2006/encryption"
export const PASSWORD_ENCRYPTOR_URI = "http://schemas.microsoft.com/office/2006/keyEncryptor/password"
export const CERTIFICATE_ENCRYPTOR_URI = "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"

// Far above any real descriptor (a certificate encryptor carries a few KB).
export const MAX_ENCRYPTION_INFO_BYTES = 1024 * 1024
// ST_SpinCount maximum (§2.3.4.10).
export const MAX_SPIN_COUNT = 10_000_000

const AES_BLOCK = 16

export interface StandardEncryptionInfo {
	readonly kind: "standard"
	readonly keyBits: 128 | 192 | 256
	readonly salt: Uint8Array
	readonly encryptedVerifier: Uint8Array
	readonly verifierHashSize: number
	readonly encryptedVerifierHash: Uint8Array
}

// Cipher parameters shared by keyData and the password key encryptor. Only AES-CBC with a supported hash
// gets this far.
export interface AgileCipher {
	readonly saltSize: number
	readonly blockSize: 16
	readonly keyBits: 128 | 192 | 256
	readonly hashSize: number
	readonly hash: HashName
	readonly salt: Uint8Array
}

export interface PasswordKeyEncryptor extends AgileCipher {
	readonly spinCount: number
	readonly encryptedVerifierHashInput: Uint8Array
	readonly encryptedVerifierHashValue: Uint8Array
	readonly encryptedKeyValue: Uint8Array
}

// Kept on a seal when its certVerifier (an HMAC of the certificate under the intermediate key) verified at
// unlock, else discarded (§2.3.4.10): it encrypts the same intermediate key, so it stays valid while the key does.
export interface CertificateKeyEncryptor {
	readonly encryptedKeyValue: Uint8Array
	readonly certificate: Uint8Array
	readonly certVerifier: Uint8Array
}

export interface DataIntegrity {
	readonly encryptedHmacKey: Uint8Array
	readonly encryptedHmacValue: Uint8Array
}

export interface AgileEncryptionInfo {
	readonly kind: "agile"
	readonly keyData: AgileCipher
	// SHOULD be present (§2.3.4.10); absent means the package cannot be checked.
	readonly dataIntegrity: DataIntegrity | undefined
	readonly password: PasswordKeyEncryptor
	readonly certificates: readonly CertificateKeyEncryptor[]
}

export type EncryptionInfo = StandardEncryptionInfo | AgileEncryptionInfo

const HASHES: Readonly<Record<string, { readonly name: HashName; readonly size: number }>> = {
	"SHA-1": { name: "SHA-1", size: 20 },
	SHA1: { name: "SHA-1", size: 20 },
	SHA256: { name: "SHA-256", size: 32 },
	"SHA-256": { name: "SHA-256", size: 32 },
	SHA384: { name: "SHA-384", size: 48 },
	"SHA-384": { name: "SHA-384", size: 48 },
	SHA512: { name: "SHA-512", size: 64 },
	"SHA-512": { name: "SHA-512", size: 64 }
}

// The names written. §2.3.4.10 lists "SHA-1", but Office, POI and msoffcrypto-tool write and expect "SHA1"
// (POI refuses "SHA-1"); both are read.
const HASH_ATTRIBUTE: Readonly<Record<HashName, string>> = {
	"SHA-1": "SHA1",
	"SHA-256": "SHA256",
	"SHA-384": "SHA384",
	"SHA-512": "SHA512"
}

export function parseEncryptionInfo(bytes: Uint8Array): EncryptionInfo {
	if (bytes.length < 8) throw corrupt("EncryptionInfo is truncated")
	if (bytes.length > MAX_ENCRYPTION_INFO_BYTES) throw corrupt("EncryptionInfo is too large")

	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const major = view.getUint16(0, true)
	const minor = view.getUint16(2, true)

	if (major === 4 && minor === 4) {
		// Attributes are read lazily: damage inside one surfaces as an XmlError wherever it is read.
		try {
			return parseAgile(bytes.subarray(8))
		} catch (error) {
			if (error instanceof XmlError) throw corrupt("EncryptionInfo XML is malformed")

			throw error
		}
	}
	if ((major === 2 || major === 3 || major === 4) && minor === 2) return parseStandard(view)
	if ((major === 3 || major === 4) && minor === 3) throw unsupported("Extensible encryption is not supported")

	throw unsupported(`EncryptionInfo version ${major}.${minor} is not supported`)
}

// §2.3.4.5: Version, Flags, HeaderSize, EncryptionHeader (§2.3.4.2), EncryptionVerifier (§2.3.4.3).
function parseStandard(view: DataView): StandardEncryptionInfo {
	const length = view.byteLength

	if (length < 12) throw corrupt("Standard EncryptionInfo is truncated")

	const headerSize = view.getUint32(8, true)

	// The header is 32 bytes before the CSP name; the verifier is 4 + 16 + 16 + 4 + 32 bytes for AES.
	if (headerSize < 32 || headerSize > length - 12 - 72) throw corrupt("Standard EncryptionHeader size is out of range")

	const header = 12
	const flags = view.getUint32(header, true)
	const algId = view.getUint32(header + 8, true)
	const algIdHash = view.getUint32(header + 12, true)
	const keySize = view.getUint32(header + 16, true)
	const fCryptoAPI = (flags & 0x04) !== 0
	const fExternal = (flags & 0x10) !== 0
	const fAES = (flags & 0x20) !== 0

	if (fExternal) throw unsupported("Extensible encryption is not supported")
	if (!fCryptoAPI || !fAES) throw unsupported("Only AES Standard encryption is supported")
	if (algIdHash !== 0 && algIdHash !== 0x8004) throw unsupported("Standard encryption hash is not SHA-1")

	const algBits = algId === 0x660e || algId === 0 ? 128 : algId === 0x660f ? 192 : algId === 0x6610 ? 256 : 0

	if (algBits === 0) throw unsupported("Standard encryption cipher is not AES")

	const keyBits = keySize === 0 ? algBits : keySize

	if (keyBits !== algBits) throw corrupt("Standard encryption key size does not match the cipher")

	const verifier = header + headerSize
	const saltSize = view.getUint32(verifier, true)

	if (saltSize !== 16) throw corrupt("Standard encryption salt is not 16 bytes")

	const verifierHashSize = view.getUint32(verifier + 36, true)

	if (verifierHashSize !== 20) throw corrupt("Standard encryption verifier hash is not 20 bytes")

	const bytes = new Uint8Array(view.buffer, view.byteOffset, view.byteLength)

	return {
		kind: "standard",
		keyBits,
		salt: bytes.slice(verifier + 4, verifier + 20),
		encryptedVerifier: bytes.slice(verifier + 20, verifier + 36),
		verifierHashSize,
		encryptedVerifierHash: bytes.slice(verifier + 40, verifier + 72)
	}
}

function parseAgile(xml: Uint8Array): AgileEncryptionInfo {
	let document: XmlDocument

	try {
		document = parseXml(xml)
	} catch (error) {
		if (error instanceof XmlError) throw corrupt("EncryptionInfo XML is malformed")

		throw error
	}

	const namespaces = new Namespaces(document)
	const root = document.documentElement()

	if (root < 0 || !namespaces.is(root, ENCRYPTION_NAMESPACE, "encryption")) throw corrupt("EncryptionInfo has no encryption element")

	let keyData: AgileCipher | undefined
	let dataIntegrity: DataIntegrity | undefined
	let password: PasswordKeyEncryptor | undefined
	const certificates: CertificateKeyEncryptor[] = []
	let encryptors = 0

	for (let child = document.firstChildElement(root); child >= 0; child = document.nextSiblingElement(child)) {
		if (namespaces.is(child, ENCRYPTION_NAMESPACE, "keyData")) {
			keyData ??= cipherOf(document, child)
		} else if (namespaces.is(child, ENCRYPTION_NAMESPACE, "dataIntegrity")) {
			dataIntegrity ??= {
				encryptedHmacKey: base64Attribute(document, child, "encryptedHmacKey"),
				encryptedHmacValue: base64Attribute(document, child, "encryptedHmacValue")
			}
		} else if (namespaces.is(child, ENCRYPTION_NAMESPACE, "keyEncryptors")) {
			for (let encryptor = document.firstChildElement(child); encryptor >= 0; encryptor = document.nextSiblingElement(encryptor)) {
				if (!namespaces.is(encryptor, ENCRYPTION_NAMESPACE, "keyEncryptor")) continue

				encryptors++

				const uri = (document.attribute(encryptor, "uri") ?? "").trim()
				const key = document.firstChildElement(encryptor)

				if (key < 0) continue

				if (uri === PASSWORD_ENCRYPTOR_URI && namespaces.is(key, PASSWORD_ENCRYPTOR_URI, "encryptedKey")) {
					if (password !== undefined) throw corrupt("EncryptionInfo has more than one password key encryptor")

					password = {
						...cipherOf(document, key),
						spinCount: numberAttribute(document, key, "spinCount", 0, MAX_SPIN_COUNT),
						encryptedVerifierHashInput: base64Attribute(document, key, "encryptedVerifierHashInput"),
						encryptedVerifierHashValue: base64Attribute(document, key, "encryptedVerifierHashValue"),
						encryptedKeyValue: base64Attribute(document, key, "encryptedKeyValue")
					}
				} else if (uri === CERTIFICATE_ENCRYPTOR_URI && namespaces.is(key, CERTIFICATE_ENCRYPTOR_URI, "encryptedKey")) {
					certificates.push({
						encryptedKeyValue: base64Attribute(document, key, "encryptedKeyValue"),
						certificate: base64Attribute(document, key, "X509Certificate"),
						certVerifier: base64Attribute(document, key, "certVerifier")
					})
				}
			}
		}
	}

	if (keyData === undefined) throw corrupt("EncryptionInfo has no keyData")
	if (encryptors === 0) throw corrupt("EncryptionInfo has no key encryptor")
	if (password === undefined) throw unsupported("The document has no password key encryptor")

	// §2.3.4.10: the password key encryptor's hashAlgorithm and cipherAlgorithm MUST be keyData's (cipherOf
	// already allows AES-CBC alone).
	if (password.hash !== keyData.hash) throw corrupt("The password key encryptor's hash is not keyData's")

	checkSizes(keyData, password, dataIntegrity)

	return { kind: "agile", keyData, dataIntegrity, password, certificates }
}

// The encrypted values must hold what is decrypted from them (each zero-padded to the block size).
function checkSizes(keyData: AgileCipher, password: PasswordKeyEncryptor, dataIntegrity: DataIntegrity | undefined): void {
	const holds = (value: Uint8Array, needed: number): boolean => value.length % AES_BLOCK === 0 && value.length >= needed

	if (
		!holds(password.encryptedVerifierHashInput, password.saltSize) ||
		!holds(password.encryptedVerifierHashValue, password.hashSize) ||
		!holds(password.encryptedKeyValue, keyData.keyBits / 8)
	) {
		throw corrupt("Password key encryptor values are truncated")
	}

	if (
		dataIntegrity !== undefined &&
		(!holds(dataIntegrity.encryptedHmacKey, 1) || !holds(dataIntegrity.encryptedHmacValue, keyData.hashSize))
	) {
		throw corrupt("Data integrity values are truncated")
	}
}

function cipherOf(document: XmlDocument, node: number): AgileCipher {
	const cipherAlgorithm = document.attribute(node, "cipherAlgorithm")
	const cipherChaining = document.attribute(node, "cipherChaining")
	const hash = HASHES[document.attribute(node, "hashAlgorithm") ?? ""]

	if (cipherAlgorithm !== "AES") throw unsupported("Agile encryption cipher is not AES")
	if (cipherChaining !== "ChainingModeCBC") throw unsupported("Agile encryption chaining is not CBC")
	if (hash === undefined) throw unsupported("Agile encryption hash is not supported")

	const saltSize = numberAttribute(document, node, "saltSize", 1, 65536)
	const blockSize = numberAttribute(document, node, "blockSize", 2, 4096)
	const keyBits = numberAttribute(document, node, "keyBits", 8, 0xffffffff)
	const hashSize = numberAttribute(document, node, "hashSize", 1, 65536)
	const salt = base64Attribute(document, node, "saltValue")

	if (blockSize !== AES_BLOCK) throw corrupt("AES block size is not 16")
	if (keyBits !== 128 && keyBits !== 192 && keyBits !== 256) throw unsupported("AES key size is not 128, 192 or 256 bits")
	if (hashSize !== hash.size) throw corrupt("Hash size does not match the hash algorithm")
	// §2.3.4.10: saltValue decodes to exactly saltSize bytes (no verifier or HMAC covers saltSize itself).
	if (salt.length !== saltSize) throw corrupt("Salt size does not match the salt")

	return { saltSize, blockSize, keyBits, hashSize, hash: hash.name, salt }
}

function numberAttribute(document: XmlDocument, node: number, name: string, min: number, max: number): number {
	const text = (document.attribute(node, name) ?? "").trim()

	if (!/^[0-9]{1,10}$/u.test(text)) throw corrupt(`${name} is not a number`)

	const value = Number(text)

	if (value < min || value > max) throw corrupt(`${name} is out of range`)

	return value
}

function base64Attribute(document: XmlDocument, node: number, name: string): Uint8Array {
	const text = document.attribute(node, name)

	if (text === undefined) throw corrupt(`${name} is missing`)

	return decodeBase64(text)
}

// Agile EncryptionInfo (§2.3.4.10): version 4.4, reserved 0x40, then the descriptor as UTF-8.
export function buildAgileInfo(
	info: Omit<AgileEncryptionInfo, "kind" | "dataIntegrity"> & { readonly dataIntegrity: DataIntegrity }
): Uint8Array {
	const cipher = (c: AgileCipher): string =>
		`saltSize="${c.saltSize}" blockSize="${c.blockSize}" keyBits="${c.keyBits}" hashSize="${c.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${HASH_ATTRIBUTE[c.hash]}" saltValue="${encodeBase64(c.salt)}"`
	const password = info.password
	const certificates = info.certificates
		.map(
			c =>
				`<keyEncryptor uri="${CERTIFICATE_ENCRYPTOR_URI}"><c:encryptedKey encryptedKeyValue="${encodeBase64(c.encryptedKeyValue)}" X509Certificate="${encodeBase64(c.certificate)}" certVerifier="${encodeBase64(c.certVerifier)}"/></keyEncryptor>`
		)
		.join("")
	const xml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
		`<encryption xmlns="${ENCRYPTION_NAMESPACE}" xmlns:p="${PASSWORD_ENCRYPTOR_URI}"${certificates === "" ? "" : ` xmlns:c="${CERTIFICATE_ENCRYPTOR_URI}"`}>` +
		`<keyData ${cipher(info.keyData)}/>` +
		`<dataIntegrity encryptedHmacKey="${encodeBase64(info.dataIntegrity.encryptedHmacKey)}" encryptedHmacValue="${encodeBase64(info.dataIntegrity.encryptedHmacValue)}"/>` +
		`<keyEncryptors><keyEncryptor uri="${PASSWORD_ENCRYPTOR_URI}"><p:encryptedKey spinCount="${password.spinCount}" ${cipher(password)} ` +
		`encryptedVerifierHashInput="${encodeBase64(password.encryptedVerifierHashInput)}" encryptedVerifierHashValue="${encodeBase64(password.encryptedVerifierHashValue)}" ` +
		`encryptedKeyValue="${encodeBase64(password.encryptedKeyValue)}"/></keyEncryptor>${certificates}</keyEncryptors></encryption>`
	// Every value is a number, a fixed token or base64: nothing to escape.
	const body = encodeUtf8(xml)
	const out = new Uint8Array(8 + body.length)
	const view = new DataView(out.buffer)

	view.setUint16(0, 4, true)
	view.setUint16(2, 4, true)
	view.setUint32(4, 0x40, true)
	out.set(body, 8)

	return out
}
