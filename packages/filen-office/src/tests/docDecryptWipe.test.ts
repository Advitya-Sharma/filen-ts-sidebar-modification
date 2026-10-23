import { createHash } from "node:crypto"
import { describe, expect, it, vi } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { readDocFile } from "../core/doc/read"
import { encryptedDoc, type EncryptOptions } from "./docDecryptFixtures"

// Every key a decryption derives (MD5 and SHA-1 results, RC4 keys) is zero once the file is read: the verifier check's
// block-0 key as much as the per-block keys.

const seen = vi.hoisted(() => ({ recording: false, arrays: [] as Uint8Array[] }))

vi.mock("../core/crypto/legacy", async importOriginal => {
	const original = await importOriginal<typeof import("../core/crypto/legacy")>()

	class RecordedRc4 extends original.Rc4 {
		constructor(key: Uint8Array) {
			super(key)

			if (seen.recording) seen.arrays.push(key)
		}
	}

	return {
		...original,
		Rc4: RecordedRc4,
		md5: (...parts: readonly Uint8Array[]): Uint8Array => {
			const out = original.md5(...parts)

			if (seen.recording) seen.arrays.push(out)

			return out
		}
	}
})

const SPEC = { pieces: [{ text: `${"Keys are wiped. ".repeat(80)}\r` }] }

describe("encrypted Word 97–2003 documents", () => {
	it.each([
		{ scheme: "rc4", password: "password1" },
		{ scheme: "cryptoApiRc4", password: "Password1234_", keySize: 128 },
		{ scheme: "cryptoApiRc4", password: "Password1234_", keySize: 40 }
	] satisfies EncryptOptions[])("wipes every derived key once the file is read ($scheme)", async options => {
		const fixture = await encryptedDoc(SPEC, options)
		const sha1 = (data: Uint8Array): Promise<Uint8Array> => {
			const out = new Uint8Array(createHash("sha1").update(data).digest())

			seen.arrays.push(out)

			return Promise.resolve(out)
		}

		seen.arrays.length = 0
		seen.recording = true

		try {
			await readDocFile(bytesSource(fixture.bytes), { decryption: { password: options.password, sha1 } })
		} finally {
			seen.recording = false
		}

		expect(seen.arrays.length).toBeGreaterThan(2)
		expect(seen.arrays.filter(array => array.some(byte => byte !== 0)).map(array => array.length)).toEqual([])

		// A wrong password leaves nothing behind either.
		seen.arrays.length = 0
		seen.recording = true

		try {
			await expect(readDocFile(bytesSource(fixture.bytes), { decryption: { password: "wrong", sha1 } })).rejects.toMatchObject({
				code: "wrongPassword"
			})
		} finally {
			seen.recording = false
		}

		expect(seen.arrays.length).toBeGreaterThan(1)
		expect(seen.arrays.filter(array => array.some(byte => byte !== 0)).map(array => array.length)).toEqual([])
	})
})
