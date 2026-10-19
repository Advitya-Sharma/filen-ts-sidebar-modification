import { createCipheriv, createDecipheriv, createHash, createHmac, randomBytes } from "node:crypto"

// Encrypted packages built with Node's crypto, an implementation independent of the WebCrypto backend, straight
// from [MS-OFFCRYPTO] §2.3.4: the backend must open them, and what it seals must open here too.

export interface EncryptedStreams {
	readonly encryptionInfo: Uint8Array
	readonly encryptedPackage: Uint8Array
}

type NodeHash = "sha1" | "sha256" | "sha384" | "sha512"

const HASH_SIZE: Record<NodeHash, number> = { sha1: 20, sha256: 32, sha384: 48, sha512: 64 }
const HASH_ATTRIBUTE: Record<NodeHash, string> = { sha1: "SHA1", sha256: "SHA256", sha384: "SHA384", sha512: "SHA512" }

const hash = (algorithm: NodeHash, ...parts: Uint8Array[]): Buffer => {
	const h = createHash(algorithm)

	for (const part of parts) h.update(part)

	return h.digest()
}

const le32 = (value: number): Buffer => {
	const out = Buffer.alloc(4)

	out.writeUInt32LE(value)

	return out
}

const fit = (bytes: Buffer, length: number, pad = 0x36): Buffer =>
	bytes.length >= length ? bytes.subarray(0, length) : Buffer.concat([bytes, Buffer.alloc(length - bytes.length, pad)])

const padZero = (bytes: Buffer): Buffer =>
	bytes.length % 16 === 0 ? bytes : Buffer.concat([bytes, Buffer.alloc(16 - (bytes.length % 16))])

function aes(mode: "cbc" | "ecb", key: Buffer, iv: Buffer | null, data: Buffer, decrypt = false): Buffer {
	const name = `aes-${key.length * 8}-${mode}`
	const cipher = decrypt ? createDecipheriv(name, key, iv) : createCipheriv(name, key, iv)

	cipher.setAutoPadding(false)

	return Buffer.concat([cipher.update(data), cipher.final()])
}

function iterate(algorithm: NodeHash, salt: Buffer, password: string, spinCount: number): Buffer {
	let h = hash(algorithm, salt, Buffer.from(password, "utf16le"))

	for (let i = 0; i < spinCount; i++) h = hash(algorithm, le32(i), h)

	return h
}

function packageHeader(size: number): Buffer {
	const out = Buffer.alloc(8)

	out.writeBigUInt64LE(BigInt(size))

	return out
}

export interface AgileOptions {
	readonly hash?: NodeHash
	readonly keyBits?: 128 | 192 | 256
	readonly spinCount?: number
	readonly saltSize?: number
	readonly integrity?: boolean
	// Extra keyEncryptor elements, as XML.
	readonly extraEncryptors?: string
	// Certificate key encryptors (§2.3.4.10): certVerifier = HMAC(keyData hash, intermediate key, certificate), or a
	// wrong one.
	readonly certificates?: readonly { readonly certificate: Uint8Array; readonly validVerifier: boolean }[]
}

export function agileStreams(plaintext: Uint8Array, password: string, options: AgileOptions = {}): EncryptedStreams {
	const algorithm = options.hash ?? "sha512"
	const keyBits = options.keyBits ?? 256
	const spinCount = options.spinCount ?? 1000
	const saltSize = options.saltSize ?? 16
	const hashSize = HASH_SIZE[algorithm]
	const keyBytes = keyBits / 8
	const secret = randomBytes(keyBytes)
	const keySalt = randomBytes(saltSize)
	const passwordSalt = randomBytes(saltSize)
	const hn = iterate(algorithm, passwordSalt, password, spinCount)
	const passwordKey = (blockKey: number[]): Buffer => fit(hash(algorithm, hn, Buffer.from(blockKey)), keyBytes)
	const passwordIv = fit(passwordSalt, 16)
	const verifierInput = randomBytes(saltSize)
	const encryptedVerifierHashInput = aes(
		"cbc",
		passwordKey([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
		passwordIv,
		padZero(verifierInput)
	)
	const encryptedVerifierHashValue = aes(
		"cbc",
		passwordKey([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
		passwordIv,
		padZero(hash(algorithm, verifierInput))
	)
	const encryptedKeyValue = aes("cbc", passwordKey([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]), passwordIv, padZero(secret))
	const segments: Buffer[] = [packageHeader(plaintext.length)]

	for (let start = 0, index = 0; start < plaintext.length; start += 4096, index++) {
		const iv = fit(hash(algorithm, keySalt, le32(index)), 16)

		segments.push(aes("cbc", secret, iv, padZero(Buffer.from(plaintext.subarray(start, start + 4096)))))
	}

	const encryptedPackage = Buffer.concat(segments)
	const blockIv = (blockKey: number[]): Buffer => fit(hash(algorithm, keySalt, Buffer.from(blockKey)), 16)
	const hmacKey = randomBytes(hashSize)
	const hmacValue = createHmac(algorithm, hmacKey).update(encryptedPackage).digest()
	const cipher = (salt: Buffer): string =>
		`saltSize="${saltSize}" blockSize="16" keyBits="${keyBits}" hashSize="${hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${HASH_ATTRIBUTE[algorithm]}" saltValue="${salt.toString("base64")}"`
	const integrity =
		options.integrity === false
			? ""
			: `<dataIntegrity encryptedHmacKey="${aes("cbc", secret, blockIv([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]), padZero(hmacKey)).toString("base64")}" encryptedHmacValue="${aes("cbc", secret, blockIv([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]), padZero(hmacValue)).toString("base64")}"/>`
	const certificates = (options.certificates ?? [])
		.map(({ certificate, validVerifier }) => {
			const verifier = createHmac(algorithm, secret).update(certificate).digest()

			if (!validVerifier) verifier[0] = (verifier[0] ?? 0) ^ 1

			return `<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><c:encryptedKey encryptedKeyValue="${padZero(secret).toString("base64")}" X509Certificate="${Buffer.from(certificate).toString("base64")}" certVerifier="${verifier.toString("base64")}"/></keyEncryptor>`
		})
		.join("")
	const xml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
		'<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
		`<keyData ${cipher(keySalt)}/>${integrity}<keyEncryptors>` +
		`<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="${spinCount}" ${cipher(passwordSalt)} ` +
		`encryptedVerifierHashInput="${encryptedVerifierHashInput.toString("base64")}" encryptedVerifierHashValue="${encryptedVerifierHashValue.toString("base64")}" encryptedKeyValue="${encryptedKeyValue.toString("base64")}"/></keyEncryptor>` +
		`${certificates}${options.extraEncryptors ?? ""}</keyEncryptors></encryption>`

	return { encryptionInfo: agileInfo(xml), encryptedPackage: new Uint8Array(encryptedPackage) }
}

export function agileInfo(xml: string): Uint8Array {
	return new Uint8Array(Buffer.concat([Buffer.from([4, 0, 4, 0, 0x40, 0, 0, 0]), Buffer.from(xml, "utf8")]))
}

// Standard (§2.3.4.5-§2.3.4.9): AES-ECB, SHA-1, 50,000 iterations, CryptDeriveKey.
export function standardStreams(plaintext: Uint8Array, password: string, keyBits: 128 | 192 | 256 = 128): EncryptedStreams {
	const salt = randomBytes(16)
	const hn = iterate("sha1", salt, password, 50_000)
	const final = hash("sha1", hn, le32(0))
	const x = (fill: number): Buffer =>
		hash(
			"sha1",
			Buffer.alloc(64, fill).map((b, k) => (k < 20 ? b ^ (final[k] ?? 0) : b))
		)
	const key = Buffer.concat([x(0x36), x(0x5c)]).subarray(0, keyBits / 8)
	const verifier = randomBytes(16)
	const csp = Buffer.from("Microsoft Enhanced RSA and AES Cryptographic Provider\0", "utf16le")
	const header = Buffer.alloc(32)
	const algId = keyBits === 128 ? 0x660e : keyBits === 192 ? 0x660f : 0x6610

	header.writeUInt32LE(0x24, 0)
	header.writeUInt32LE(algId, 8)
	header.writeUInt32LE(0x8004, 12)
	header.writeUInt32LE(keyBits, 16)
	header.writeUInt32LE(0x18, 20)

	const head = Buffer.alloc(12)

	head.writeUInt16LE(4, 0)
	head.writeUInt16LE(2, 2)
	head.writeUInt32LE(0x24, 4)
	head.writeUInt32LE(header.length + csp.length, 8)

	const encryptionInfo = Buffer.concat([
		head,
		header,
		csp,
		le32(16),
		salt,
		aes("ecb", key, null, verifier),
		le32(20),
		aes("ecb", key, null, padZero(hash("sha1", verifier)))
	])
	const encryptedPackage = Buffer.concat([packageHeader(plaintext.length), aes("ecb", key, null, padZero(Buffer.from(plaintext)))])

	return { encryptionInfo: new Uint8Array(encryptionInfo), encryptedPackage: new Uint8Array(encryptedPackage) }
}

// Decrypts an Agile package with Node's crypto: the independent check of what the backend seals.
export function decryptAgileStreams(streams: EncryptedStreams, password: string): Buffer {
	const xml = Buffer.from(streams.encryptionInfo.subarray(8)).toString("utf8")
	const attribute = (element: string, name: string): string => {
		const match = new RegExp(`<${element}\\b[^>]*\\s${name}="([^"]*)"`, "u").exec(xml)

		if (match?.[1] === undefined) throw new Error(`${element}@${name} missing`)

		return match[1]
	}
	const nodeHash = (name: string): NodeHash =>
		({ SHA1: "sha1", "SHA-1": "sha1", SHA256: "sha256", SHA384: "sha384", SHA512: "sha512" })[name] as NodeHash
	const b64 = (value: string): Buffer => Buffer.from(value, "base64")
	const encryptor = "p:encryptedKey"
	const algorithm = nodeHash(attribute(encryptor, "hashAlgorithm"))
	const keyBytes = Number(attribute(encryptor, "keyBits")) / 8
	const salt = b64(attribute(encryptor, "saltValue"))
	const hn = iterate(algorithm, salt, password, Number(attribute(encryptor, "spinCount")))
	const passwordKey = (blockKey: number[]): Buffer => fit(hash(algorithm, hn, Buffer.from(blockKey)), keyBytes)
	const iv = fit(salt, 16)
	const input = aes(
		"cbc",
		passwordKey([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
		iv,
		b64(attribute(encryptor, "encryptedVerifierHashInput")),
		true
	)
	const value = aes(
		"cbc",
		passwordKey([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
		iv,
		b64(attribute(encryptor, "encryptedVerifierHashValue")),
		true
	)
	const saltSize = Number(attribute(encryptor, "saltSize"))
	const hashSize = Number(attribute(encryptor, "hashSize"))

	if (!hash(algorithm, input.subarray(0, saltSize)).equals(value.subarray(0, hashSize))) throw new Error("wrong password")

	const keyDataBits = Number(attribute("keyData", "keyBits"))
	const secret = aes(
		"cbc",
		passwordKey([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
		iv,
		b64(attribute(encryptor, "encryptedKeyValue")),
		true
	).subarray(0, keyDataBits / 8)
	const keyAlgorithm = nodeHash(attribute("keyData", "hashAlgorithm"))
	const keySalt = b64(attribute("keyData", "saltValue"))
	const blockIv = (blockKey: number[]): Buffer => fit(hash(keyAlgorithm, keySalt, Buffer.from(blockKey)), 16)
	const stream = Buffer.from(streams.encryptedPackage)
	const keyHashSize = Number(attribute("keyData", "hashSize"))
	const hmacKey = aes(
		"cbc",
		secret,
		blockIv([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
		b64(attribute("dataIntegrity", "encryptedHmacKey")),
		true
	)
	const hmacValue = aes(
		"cbc",
		secret,
		blockIv([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
		b64(attribute("dataIntegrity", "encryptedHmacValue")),
		true
	)

	if (!createHmac(keyAlgorithm, hmacKey.subarray(0, keyHashSize)).update(stream).digest().equals(hmacValue.subarray(0, keyHashSize))) {
		throw new Error("integrity")
	}

	const size = Number(stream.readBigUInt64LE(0))
	const out: Buffer[] = []

	for (let start = 0, index = 0; start < size; start += 4096, index++) {
		const end = Math.min(start + 4096, Math.ceil(size / 16) * 16)

		out.push(aes("cbc", secret, fit(hash(keyAlgorithm, keySalt, le32(index)), 16), stream.subarray(8 + start, 8 + end), true))
	}

	return Buffer.concat(out).subarray(0, size)
}
