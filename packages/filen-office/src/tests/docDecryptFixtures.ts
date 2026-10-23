import { createHash } from "node:crypto"
import { CfbFile, cfbBytes, type CfbWriteEntry } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import { Rc4 } from "../core/crypto/legacy"
import { legacyPasswordVerifier } from "../core/docx/protection/hash"
import { buildDoc, type DocSpec } from "./docFixtures"

// Encrypts documents built by buildDoc the way Word writes a password to open ([MS-DOC] §2.2.6): XOR obfuscation
// ([MS-OFFCRYPTO] §2.3.7), RC4 (§2.3.6) or CryptoAPI RC4 (§2.3.5). Hashes come from node:crypto, not the reader's MD5.

export type FixtureScheme = "xor" | "rc4" | "cryptoApiRc4"

export interface EncryptOptions {
	readonly scheme: FixtureScheme
	readonly password: string
	// CryptoAPI KeySize in bits (0 stands for 40).
	readonly keySize?: number
	// The single-byte password XOR is keyed with (default: each UTF-16 unit's low byte).
	readonly xorBytes?: Uint8Array
}

export interface EncryptedFixture {
	readonly bytes: Uint8Array
	// The plain streams the encrypted file was made from, by name.
	readonly plain: ReadonlyMap<string, Uint8Array>
	readonly lKey: number
}

const SALT = Uint8Array.from({ length: 16 }, (_, k) => (k * 37 + 11) & 0xff)
const VERIFIER = Uint8Array.from({ length: 16 }, (_, k) => (k * 91 + 5) & 0xff)
const CSP = "Microsoft Enhanced Cryptographic Provider v1.0\0"

function hash(algorithm: "md5" | "sha1", ...parts: Uint8Array[]): Uint8Array {
	const h = createHash(algorithm)

	for (const part of parts) h.update(part)

	return new Uint8Array(h.digest())
}

function utf16(text: string): Uint8Array {
	return new Uint8Array(Buffer.from(text, "utf16le"))
}

function u32(value: number): Uint8Array {
	return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff)
}

function concat(...parts: Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
	let at = 0

	for (const part of parts) {
		out.set(part, at)
		at += part.length
	}

	return out
}

function rc4Keys(password: string): (block: number) => Uint8Array {
	const h0 = hash("md5", utf16(password)).subarray(0, 5)
	const h1 = hash("md5", ...Array.from({ length: 16 }, () => concat(h0, SALT))).subarray(0, 5)

	return block => hash("md5", h1, u32(block))
}

function cryptoApiKeys(password: string, keySize: number): (block: number) => Uint8Array {
	const h0 = hash("sha1", SALT, utf16(password))
	const bits = keySize === 0 ? 40 : keySize

	return block => {
		const full = hash("sha1", h0, u32(block))

		return bits === 40 ? concat(full.subarray(0, 5), new Uint8Array(11)) : full.subarray(0, bits / 8)
	}
}

function blocks(input: Uint8Array, keyOf: (block: number) => Uint8Array): Uint8Array {
	const out = new Uint8Array(input.length)

	for (let at = 0; at < input.length; at += 512) new Rc4(keyOf(at / 512)).apply(input, out, at, at, Math.min(512, input.length - at))

	return out
}

// [MS-OFFCRYPTO] §2.3.7.2, transcribed for the encrypting side.
function xorArray(password: Uint8Array): Uint8Array {
	const pad = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00]
	const key = legacyPasswordVerifier(password) >>> 16
	const ror = (a: number, b: number): number => {
		const v = (a ^ b) & 0xff

		return ((v >> 1) | (v << 7)) & 0xff
	}
	const out = new Uint8Array(16)
	let index = password.length

	if (index % 2 === 1) {
		out[index] = ror(pad[0] ?? 0, key >> 8)
		index -= 1
		out[index] = ror(password[password.length - 1] ?? 0, key & 0xff)
	}

	while (index > 0) {
		index -= 1
		out[index] = ror(password[index] ?? 0, key >> 8)
		index -= 1
		out[index] = ror(password[index] ?? 0, key & 0xff)
	}

	let at = 15

	for (let padIndex = 15 - password.length; padIndex > 0;) {
		out[at] = ror(pad[padIndex] ?? 0, key >> 8)
		at -= 1
		padIndex -= 1
		out[at] = ror(pad[padIndex] ?? 0, key & 0xff)
		at -= 1
		padIndex -= 1
	}

	return out
}

function xorStream(input: Uint8Array, array: Uint8Array): Uint8Array {
	return input.map((byte, k) => {
		const mask = array[k % 16] ?? 0

		return byte === 0 || (byte ^ mask) === 0 ? byte : byte ^ mask
	})
}

// The EncryptionHeader written at the start of the Table stream, its verifier encrypted with block 0's key.
function header(options: EncryptOptions): Uint8Array {
	if (options.scheme === "rc4") {
		const cipher = new Rc4(rc4Keys(options.password)(0))
		const sealed = new Uint8Array(32)

		cipher.apply(concat(VERIFIER, hash("md5", VERIFIER)), sealed)

		return concat(Uint8Array.of(1, 0, 1, 0), SALT, sealed)
	}

	const keySize = options.keySize ?? 128
	const csp = utf16(CSP)
	const encryptionHeader = concat(u32(0x04), u32(0), u32(0x6801), u32(0x8004), u32(keySize), u32(1), u32(0), u32(0), csp)
	const cipher = new Rc4(cryptoApiKeys(options.password, keySize)(0))
	const verifier = new Uint8Array(16)
	const verifierHash = new Uint8Array(20)

	cipher.apply(VERIFIER, verifier)
	cipher.apply(hash("sha1", VERIFIER), verifierHash)

	return concat(
		Uint8Array.of(4, 0, 2, 0),
		u32(0x04),
		u32(encryptionHeader.length),
		encryptionHeader,
		u32(16),
		SALT,
		verifier,
		u32(20),
		verifierHash
	)
}

export function headerLength(options: EncryptOptions): number {
	return options.scheme === "xor" ? 0 : header(options).length
}

// Changes made to the encrypted WordDocument and Table streams before they are written (hostile-input tests).
export type Tamper = (word: Uint8Array, table: Uint8Array) => void

// Builds `spec` (its Table stream starts with room for the EncryptionHeader) and encrypts it as `options` says.
export async function encryptedDoc(spec: DocSpec, options: EncryptOptions, tamper?: Tamper): Promise<EncryptedFixture> {
	const prefix = options.scheme === "xor" ? new Uint8Array(0) : header(options)
	const plainFile = await buildDoc({ ...spec, tablePrefix: new Uint8Array(prefix.length) })
	const cfb = await CfbFile.open(bytesSource(plainFile))
	const plain = new Map<string, Uint8Array>()

	for (const entry of cfb.walk())
		if (entry.kind === "stream" && entry.path.length === 1) plain.set(entry.name, (await cfb.read(entry)).slice())

	const word = plain.get("WordDocument") ?? new Uint8Array(0)
	const tableName = plain.has("1Table") ? "1Table" : "0Table"
	const xorBytes = options.xorBytes ?? Uint8Array.from(options.password, char => char.charCodeAt(0) & 0xff)
	const encrypt = (input: Uint8Array): Uint8Array =>
		options.scheme === "xor"
			? xorStream(input, xorArray(xorBytes))
			: blocks(input, options.scheme === "rc4" ? rc4Keys(options.password) : cryptoApiKeys(options.password, options.keySize ?? 128))
	const lKey = options.scheme === "xor" ? legacyPasswordVerifier(xorBytes) : prefix.length
	const sealedWord = encrypt(word)

	sealedWord.set(word.subarray(0, 68))
	sealedWord[11] = (sealedWord[11] ?? 0) | 0x01 | (options.scheme === "xor" ? 0x80 : 0)
	sealedWord.set(u32(lKey), 14)

	const table = plain.get(tableName) ?? new Uint8Array(0)
	const sealedTable = encrypt(table)

	sealedTable.set(prefix)
	tamper?.(sealedWord, sealedTable)

	const entries: CfbWriteEntry[] = [...plain].map(([name, bytes]) => ({
		path: [name],
		bytes: name === "WordDocument" ? sealedWord : name === tableName ? sealedTable : name === "Data" ? encrypt(bytes) : bytes
	}))

	return { bytes: await cfbBytes(entries), plain, lKey }
}
