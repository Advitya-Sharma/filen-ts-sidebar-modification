import { spawnSync } from "node:child_process"
import { createHash } from "node:crypto"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, describe, expect, it } from "vitest"
import { CfbFile } from "@filen/office/core/cfb"
import { DocError } from "@filen/office/core/doc"
import { bytesSource } from "@filen/office/core/opc"
import { CLEAR_WORD_BYTES, type DocDecryption } from "../core/doc/decrypt"
import { readDocFile } from "../core/doc/read"
import { CORPUS_DIR } from "./corpusDir"
import { encryptedDoc, type EncryptOptions } from "./docDecryptFixtures"
import { msoffcryptoPython } from "./encryption/oracles/msoffcrypto"

// Our decrypted streams against msoffcrypto-tool's, byte for byte: the WordDocument stream past FibBase, the Table
// stream past its EncryptionHeader and the whole Data stream (msoffcrypto rewrites FibBase and leaves the header
// region as decrypted noise). Opt-in through FILEN_OFFICE_MSOFFCRYPTO; corpus files when the corpus is fetched.

const python = msoffcryptoPython()
const SCRIPT = join(import.meta.dirname, "docDecryptOracle.py")
const scratch = python === undefined ? "" : mkdtempSync(join(tmpdir(), "doc-decrypt-oracle-"))

afterAll(() => {
	if (scratch !== "") rmSync(scratch, { recursive: true, force: true })
})

function decryption(password: string): DocDecryption {
	return { password, sha1: data => Promise.resolve(new Uint8Array(createHash("sha1").update(data).digest())) }
}

async function streams(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
	const cfb = await CfbFile.open(bytesSource(bytes))
	const out = new Map<string, Uint8Array>()

	for (const entry of cfb.walk()) if (entry.kind === "stream" && entry.path.length === 1) out.set(entry.name, await cfb.read(entry))

	return out
}

function oracle(file: string, password: string): Uint8Array | string {
	const out = join(scratch, `${Math.random().toString(36).slice(2)}.doc`)
	const result = spawnSync(python ?? "", ["-I", SCRIPT, file, password, out], { encoding: "utf8", timeout: 120_000 })
	const answer = JSON.parse(result.stdout.trim().split("\n").at(-1) ?? "{}") as { ok?: boolean; error?: string }

	return answer.ok === true ? readFileSync(out) : (answer.error ?? `exit ${result.status}`)
}

interface Compared {
	readonly word: boolean
	readonly table: boolean
	readonly data: boolean | "none"
}

async function compare(path: string, password: string): Promise<Compared> {
	const ours = await readDocFile(bytesSource(readFileSync(path)), { decryption: decryption(password) })
	const theirs = oracle(path, password)

	if (typeof theirs === "string") throw new Error(`msoffcrypto refused: ${theirs}`)

	const decrypted = await streams(theirs)
	const tableName = decrypted.has("1Table") ? "1Table" : "0Table"
	const lKey = await lKeyOf(path)
	const equal = (a: Uint8Array | undefined, b: Uint8Array | undefined, from: number): boolean =>
		a !== undefined && b !== undefined && Buffer.from(a.subarray(from)).equals(b.subarray(from))

	return {
		word: equal(ours.word, decrypted.get("WordDocument"), CLEAR_WORD_BYTES),
		table: equal(ours.table, decrypted.get(tableName), lKey),
		data: decrypted.has("Data") ? equal(ours.data, decrypted.get("Data"), 0) : "none"
	}
}

async function lKeyOf(path: string): Promise<number> {
	const word = (await streams(readFileSync(path))).get("WordDocument")

	return word === undefined ? 0 : new DataView(word.buffer, word.byteOffset).getUint32(14, true)
}

// The encrypted .doc files in the corpus and their passwords (Apache POI test data, named for their password).
const CORPUS: readonly (readonly [file: string, password: string, data: boolean | "none"])[] = [
	["doc/poi__password_tika_binaryrc4.doc", "tika", "none"],
	["doc/poi__password_password_cryptoapi.doc", "password", true]
]

describe.skipIf(python === undefined)("encrypted .doc against msoffcrypto-tool", () => {
	const corpus = existsSync(join(CORPUS_DIR, "doc"))

	it.skipIf(!corpus).each(CORPUS)("decrypts %s as msoffcrypto does", async (file, password, data) => {
		expect(await compare(join(CORPUS_DIR, file), password)).toEqual({ word: true, table: true, data })
	})

	it.skipIf(!corpus)("refuses the corpus file whose password is not known, as msoffcrypto does", async () => {
		const path = join(CORPUS_DIR, "doc/poi__PasswordProtected.doc")
		const codes: string[] = []

		for (const guess of ["password", "VelvetSweatshop", "test"]) {
			try {
				await readDocFile(bytesSource(readFileSync(path)), { decryption: decryption(guess) })
				codes.push("read")
			} catch (error) {
				codes.push(error instanceof DocError ? error.code : "untyped")
			}

			codes.push(typeof oracle(path, guess) === "string" ? "refused" : "decrypted")
		}

		expect(codes).toEqual(["wrongPassword", "refused", "wrongPassword", "refused", "wrongPassword", "refused"])
	})

	const SYNTHETIC: readonly EncryptOptions[] = [
		{ scheme: "rc4", password: "password1" },
		{ scheme: "cryptoApiRc4", password: "Pässwort", keySize: 128 },
		{ scheme: "cryptoApiRc4", password: "Pässwort", keySize: 56 },
		{ scheme: "cryptoApiRc4", password: "Pässwort", keySize: 40 },
		{ scheme: "cryptoApiRc4", password: "Pässwort", keySize: 0 }
	]

	it.each(SYNTHETIC.map(options => [`${options.scheme} ${options.keySize ?? ""}`, options] as const))(
		"decrypts a %s fixture as msoffcrypto does",
		async (_, options) => {
			const text = `${"Line of text for several rekey blocks. ".repeat(80)}\r`
			const data = Uint8Array.from({ length: 2000 }, (_, k) => (k * 13) & 0xff)
			const fixture = await encryptedDoc({ pieces: [{ text }], streams: [{ path: ["Data"], bytes: data }] }, options)
			const path = join(scratch, `${options.scheme}-${options.keySize ?? 0}.doc`)

			writeFileSync(path, fixture.bytes)

			expect(await compare(path, options.password)).toEqual({ word: true, table: true, data: true })
		}
	)
})
