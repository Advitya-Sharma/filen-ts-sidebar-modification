import { createHash } from "node:crypto"
import { describe, expect, it } from "vitest"
import { DocError, docText } from "@filen/office/core/doc"
import { bytesSource } from "@filen/office/core/opc"
import { CLEAR_WORD_BYTES, type DocDecryption } from "../core/doc/decrypt"
import { readDocFile } from "../core/doc/read"
import { convertDocToDocx } from "../core/doc/toDocx/index"
import { legacyPasswordVerifier } from "../core/docx/protection/hash"
import { encryptedDoc, type EncryptOptions, type FixtureScheme, type Tamper } from "./docDecryptFixtures"
import type { DocSpec } from "./docFixtures"

function decryption(password: string): DocDecryption {
	return { password, sha1: data => Promise.resolve(new Uint8Array(createHash("sha1").update(data).digest())) }
}

async function failure(run: () => Promise<unknown>): Promise<DocError> {
	try {
		await run()
	} catch (error) {
		if (error instanceof DocError) return error

		throw error
	}

	throw new Error("expected a DocError")
}

// Long enough for the WordDocument stream to span several 512-byte rekey blocks, with a Data stream of its own.
const TEXT = `${"The quick brown fox jumps over the lazy dog. ".repeat(60)}\r`
const DATA = Uint8Array.from({ length: 1500 }, (_, k) => (k * 7 + 3) & 0xff)
const SPEC: DocSpec = { pieces: [{ text: TEXT }], streams: [{ path: ["Data"], bytes: DATA }] }

const SCHEMES: readonly EncryptOptions[] = [
	{ scheme: "xor", password: "Example" },
	{ scheme: "rc4", password: "password1" },
	{ scheme: "cryptoApiRc4", password: "Password1234_", keySize: 128 },
	{ scheme: "cryptoApiRc4", password: "Password1234_", keySize: 56 },
	{ scheme: "cryptoApiRc4", password: "Password1234_", keySize: 40 },
	{ scheme: "cryptoApiRc4", password: "Password1234_", keySize: 0 }
]

const label = (options: EncryptOptions): string => `${options.scheme}${options.keySize === undefined ? "" : ` ${options.keySize}-bit`}`

describe("encrypted Word 97–2003 documents", () => {
	it.each(SCHEMES.map(options => [label(options), options] as const))("opens %s with its password", async (_, options) => {
		const fixture = await encryptedDoc(SPEC, options)
		const file = await readDocFile(bytesSource(fixture.bytes), { decryption: decryption(options.password) })
		const plainWord = fixture.plain.get("WordDocument") ?? new Uint8Array(0)
		const plainTable = fixture.plain.get("1Table") ?? new Uint8Array(0)
		const clearTable = options.scheme === "xor" ? 0 : fixture.lKey

		expect({
			encryption: file.encryption,
			text: docText(file.model).trim(),
			word: Buffer.from(file.word.subarray(CLEAR_WORD_BYTES)).equals(plainWord.subarray(CLEAR_WORD_BYTES)),
			table: Buffer.from(file.table.subarray(clearTable)).equals(plainTable.subarray(clearTable)),
			data: file.data !== undefined && Buffer.from(file.data).equals(DATA)
		}).toEqual({ encryption: options.scheme, text: TEXT.trim(), word: true, table: true, data: true })
	})

	it.each(SCHEMES.map(options => [label(options), options] as const))("refuses %s with a wrong password", async (_, options) => {
		const fixture = await encryptedDoc(SPEC, options)

		for (const password of ["wrong", "", options.password.toUpperCase(), `${options.password}x`]) {
			const error = await failure(() => readDocFile(bytesSource(fixture.bytes), { decryption: decryption(password) }))

			expect([password, error.code]).toEqual([password, "wrongPassword"])
		}
	})

	it("refuses an encrypted file opened without a password, naming its scheme", async () => {
		const found: [FixtureScheme | undefined, string][] = []

		for (const options of SCHEMES.slice(0, 3)) {
			const error = await failure(async () => readDocFile(bytesSource((await encryptedDoc(SPEC, options)).bytes)))

			found.push([error.encryption as FixtureScheme | undefined, error.code])
		}

		expect(found).toEqual([
			["xor", "encrypted"],
			["rc4", "encrypted"],
			["cryptoApiRc4", "encrypted"]
		])
	})

	it("refuses a password longer than 255 characters before hashing it", async () => {
		const fixture = await encryptedDoc(SPEC, { scheme: "cryptoApiRc4", password: "p" })
		let hashed = 0
		const error = await failure(() =>
			readDocFile(bytesSource(fixture.bytes), {
				decryption: {
					password: "p".repeat(256),
					sha1: data => {
						hashed++

						return decryption("").sha1(data)
					}
				}
			})
		)

		expect([error.code, hashed]).toEqual(["wrongPassword", 0])
	})

	it("accepts either single-byte form of an XOR password (§2.3.7.4)", async () => {
		// "€" is 0x80 in Windows-1252 and 0xAC as the low byte of U+20AC: a file keyed with either opens.
		const opened: string[] = []

		for (const xorBytes of [Uint8Array.of(0x80, 0x31), Uint8Array.of(0xac, 0x31)]) {
			const fixture = await encryptedDoc(SPEC, { scheme: "xor", password: "€1", xorBytes })
			const file = await readDocFile(bytesSource(fixture.bytes), { decryption: decryption("€1") })

			opened.push(docText(file.model).trim().slice(0, 9))
		}

		expect(opened).toEqual(["The quick", "The quick"])
	})

	it("keys XOR with the first 15 characters, as Word does", async () => {
		const fixture = await encryptedDoc(SPEC, { scheme: "xor", password: "abcdefghijklmnopq" })
		const file = await readDocFile(bytesSource(fixture.bytes), { decryption: decryption("abcdefghijklmnoXYZ") })

		expect(file.encryption).toBe("xor")
	})

	it("matches the published XOR verifier of the default password", () => {
		// [MS-XLS] / msoffcrypto: VelvetSweatshop → CreatePasswordVerifier_Method1 0x9A0A.
		const bytes = Uint8Array.from("VelvetSweatshop", char => char.charCodeAt(0))

		expect((legacyPasswordVerifier(bytes) & 0xffff).toString(16)).toBe("9a0a")
	})

	it("leaves the source's bytes as they were", async () => {
		const fixture = await encryptedDoc(SPEC, { scheme: "rc4", password: "pw" })
		const before = fixture.bytes.slice()

		await readDocFile(bytesSource(fixture.bytes), { decryption: decryption("pw") })

		expect(Buffer.from(fixture.bytes).equals(before)).toBe(true)
	})

	it("converts a decrypted document to an unencrypted package and says it was encrypted", async () => {
		const fixture = await encryptedDoc(SPEC, { scheme: "cryptoApiRc4", password: "pw", keySize: 128 })
		const result = await convertDocToDocx(bytesSource(fixture.bytes), { decryption: decryption("pw") })
		const document = new TextDecoder().decode(result.parts.get("word/document.xml"))

		expect([result.encryption, document.includes("The quick brown fox")]).toEqual(["cryptoApiRc4", true])
	})
})

describe("hostile encryption headers", () => {
	// A seeded generator: the same mutations every run.
	function random(seed: number): () => number {
		let state = seed >>> 0

		return () => {
			state = (Math.imul(state ^ (state >>> 15), 0x2c1b3c6d) + 0x9e3779b9) >>> 0

			return state / 0x100000000
		}
	}

	async function outcome(bytes: Uint8Array, password: string): Promise<string> {
		try {
			await readDocFile(bytesSource(bytes), { decryption: decryption(password) })

			return "read"
		} catch (error) {
			if (error instanceof DocError) return error.code

			return `untyped: ${error instanceof Error ? error.message : String(error)}`
		}
	}

	const u32At = (bytes: Uint8Array, offset: number, value: number): void =>
		new DataView(bytes.buffer, bytes.byteOffset).setUint32(offset, value, true)
	const tampered = async (options: EncryptOptions, tamper: Tamper): Promise<string> =>
		outcome((await encryptedDoc(SPEC, options, tamper)).bytes, options.password)

	it("refuses malformed CryptoAPI headers with typed reasons", async () => {
		const options: EncryptOptions = { scheme: "cryptoApiRc4", password: "pw" }

		expect([
			await tampered(options, (_, table) => u32At(table, 8, 0xffffffff)), // EncryptionHeaderSize past the header
			await tampered(options, (_, table) => u32At(table, 8, 4)), // EncryptionHeaderSize too small
			await tampered(options, (_, table) => u32At(table, 12 + 8, 0x660e)), // AlgID AES-128
			await tampered(options, (_, table) => u32At(table, 12 + 16, 200)), // KeySize past 128 bits
			await tampered(options, (_, table) => u32At(table, 12 + 16, 44)), // KeySize not a whole number of bytes
			await tampered(options, (_, table) => u32At(table, 12 + 32 + 94, 15)) // SaltSize not 16
		]).toEqual(["corrupt", "corrupt", "encrypted", "corrupt", "corrupt", "corrupt"])
	})

	it("refuses an lKey short of the header or past the Table stream", async () => {
		const rc4: EncryptOptions = { scheme: "rc4", password: "pw" }
		const cryptoApi: EncryptOptions = { scheme: "cryptoApiRc4", password: "pw" }

		expect([
			await tampered(rc4, word => u32At(word, 14, 10)),
			await tampered(rc4, word => u32At(word, 14, 0)),
			await tampered(cryptoApi, word => u32At(word, 14, 40)),
			await tampered(cryptoApi, word => u32At(word, 14, 0x7fffffff))
		]).toEqual(["corrupt", "corrupt", "corrupt", "corrupt"])
	})

	it.each(["xor", "rc4", "cryptoApiRc4"] as const)("answers every mutated %s header and FibBase with a typed outcome", async scheme => {
		const next = random(scheme.length * 7919)
		const seen = new Set<string>()

		for (let round = 0; round < 120; round++) {
			const edits = Array.from({ length: 1 + Math.floor(next() * 4) }, () => ({
				inTable: next() < 0.7,
				at: Math.floor(next() * 260),
				value: Math.floor(next() * 256)
			}))
			const result = await tampered({ scheme, password: "pw" }, (word, table) => {
				// The Table stream's header region, or FibBase's flags and lKey.
				for (const edit of edits) {
					if (edit.inTable) table[edit.at] = edit.value
					else word[10 + (edit.at % 8)] = edit.value
				}
			})

			expect(result.startsWith("untyped")).toBe(false)
			seen.add(result)
		}

		expect(
			[...seen].every(result =>
				["read", "corrupt", "encrypted", "wrongPassword", "badSignature", "legacyVersion", "tooLarge"].includes(result)
			)
		).toBe(true)
	})
})
