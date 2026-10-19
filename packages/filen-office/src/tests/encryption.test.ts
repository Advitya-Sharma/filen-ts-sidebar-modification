import { createHash, randomBytes } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import {
	AGILE_DEFAULT_SPIN_COUNT,
	DEFAULT_OFFICE_PASSWORD,
	OfficeCryptoError,
	classifyDataSpaces,
	parseEncryptionInfo,
	probeEncryptionInfo,
	releaseOfficeKey,
	unlockOfficePackage,
	type AgileEncryptionInfo,
	type OfficeKey,
	type OfficeUnlockResult
} from "../dom/encryption/index"
import { openSealedPackage, sealOfficePackage } from "../dom/encryption/seal"
import { decodeBase64, encodeBase64 } from "../dom/encryption/bytes"
import { CORPUS_DIR } from "./corpusDir"
import { readCfbStreams } from "./encryption/cfbStreams"
import { agileInfo, agileStreams, decryptAgileStreams, standardStreams, type EncryptedStreams } from "./encryption/generate"

const sha256 = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex")

function unlocked(result: OfficeUnlockResult): Extract<OfficeUnlockResult, { ok: true }> {
	if (!result.ok) throw new Error(`unlock failed: ${result.reason} ${result.message}`)

	return result
}

const unlock = (streams: EncryptedStreams, password: string, upgradeSpinCount = 100): Promise<OfficeUnlockResult> =>
	unlockOfficePackage(streams.encryptionInfo, streams.encryptedPackage, password, { upgradeSpinCount })

const reason = async (promise: Promise<OfficeUnlockResult>): Promise<string> => {
	const result = await promise

	return result.ok ? "ok" : result.reason
}

async function rejection(promise: Promise<unknown>): Promise<string> {
	try {
		await promise
	} catch (error) {
		if (error instanceof OfficeCryptoError) return error.code

		throw error
	}

	return "resolved"
}

const agileOf = (bytes: Uint8Array): AgileEncryptionInfo => {
	const info = parseEncryptionInfo(bytes)

	if (info.kind !== "agile") throw new Error("not agile")

	return info
}

// Sizes around the block, the 4,096-byte segment and the 256-segment window.
const SIZES = [0, 1, 15, 16, 17, 4095, 4096, 4097, 3 * 4096 + 5, 256 * 4096, 256 * 4096 + 17, 600_000]

describe("office encryption: corpus known answers", () => {
	// Expected SHA-256 of the decrypted package, matched by Apache POI 5.4.1 for every sample and by
	// msoffcrypto-tool 6.0.0 for all but the Agile SHA-1 ones (it neither pads a SHA-1 key to 256 bits nor cuts a
	// SHA-1 verifier hash to 20 bytes). oxsdk's password is not published with its tests; "P@ssword" was
	// recovered by dictionary and confirmed by both tools.
	const samples = [
		{
			path: "docx/lo__Encrypted_LO_Standard_abc.docx",
			password: "abc",
			scheme: "standard",
			dataSpaces: "absent",
			size: 4104,
			sha: "e3f8eee962c7796c25ffa78b53ed90aae93034661c024f3ae3b9be4f91458cc5"
		},
		{
			path: "docx/poi__bug53475-password-is-pass.docx",
			password: "pass",
			scheme: "agile",
			dataSpaces: "password",
			size: 23162,
			sha: "8c9e00cbff38ca95b1f1fb2927a18d3de2087e64410f48759239cde34c1d0566"
		},
		{
			path: "docx/poi__bug53475-password-is-solrcell.docx",
			password: "solrcell",
			scheme: "standard",
			dataSpaces: "password",
			size: 24950,
			sha: "2ff6f70927616e8e39d81ffddc17ad5980827a0e8f26034d394ea3326287e123"
		},
		{
			path: "docx/oxsdk__document_with_password.docx",
			password: "P@ssword",
			scheme: "standard",
			dataSpaces: "password",
			size: 11373,
			sha: "94cff7504a2246b909d005c9f7f7b92b9cd9a26bc3f2585ff4aca84eba8da06c"
		},
		{
			path: "xlsx/poi__protected_passtika.xlsx",
			password: "tika",
			scheme: "standard",
			dataSpaces: "password",
			size: 8230,
			sha: "2bf260af678ac1b93b5dc618e0892aff98062f4566b3b6f5f3c9cba4e02b8d68"
		},
		{
			path: "xlsx/poi__58616.xlsx",
			password: DEFAULT_OFFICE_PASSWORD,
			scheme: "standard",
			dataSpaces: "password",
			size: 120332,
			sha: "2f5bc342ad84b8c49f53f1417d541f336cdf398e49c7d6725602108573cf560b"
		},
		{
			path: "docx/lo__Encrypted_MSO2007_abc.docx",
			password: "abc",
			scheme: "standard",
			dataSpaces: "password",
			size: 9983,
			sha: "4d5cb0b74843e055f000a7a1185afb3ad69e6c5f817038fd48c1ed10c7f4c6da"
		},
		{
			path: "docx/lo__Encrypted_MSO2010_abc.docx",
			password: "abc",
			scheme: "agile",
			dataSpaces: "password",
			size: 12880,
			sha: "cc7f265bd0f5c58ddf8cc1ada04035dc9f694f0472d14c305c35819c663c7db1"
		},
		{
			path: "docx/lo__Encrypted_MSO2013_abc.docx",
			password: "abc",
			scheme: "agile",
			dataSpaces: "password",
			size: 11339,
			sha: "78cf32719ce59290d1bac1bfff636267286f61298138cde320ad71435a7b092a"
		},
		{
			path: "docx/mso__example_password.docx",
			password: "Password1234_",
			scheme: "agile",
			dataSpaces: "password",
			size: 11995,
			sha: "8c8212db6e624bfc69286e94d09b7e68c753ee86b6826e51427a33c841f133d1"
		},
		{
			path: "pptx/tika__testPPT_protected_passtika.pptx",
			password: "tika",
			scheme: "standard",
			dataSpaces: "password",
			size: 37358,
			sha: "8fbe6db15e8a1db6a633bfe6272ca8a9822a505263c3f9f369d53ec959c66a20"
		},
		{
			path: "xlsx/poi__protected_sha512.xlsx",
			password: "this is a test",
			scheme: "agile",
			dataSpaces: "password",
			size: 8078,
			sha: "57e6892ba449ce51749df42210b396435000a2fea8215314912219e015318775"
		}
	] as const

	const streamsOf = (path: string): EncryptedStreams & { readonly dataSpaces: string } => {
		const cfb = readCfbStreams(new Uint8Array(readFileSync(join(CORPUS_DIR, path))))
		const encryptionInfo = cfb.stream("EncryptionInfo")
		const encryptedPackage = cfb.stream("EncryptedPackage")

		if (encryptionInfo === undefined || encryptedPackage === undefined) throw new Error("not an encrypted package")

		const dataSpaces = classifyDataSpaces({
			dataSpaceMap: cfb.stream("\u0006DataSpaces", "DataSpaceMap"),
			dataSpaceDefinition: name => cfb.stream("\u0006DataSpaces", "DataSpaceInfo", name),
			transformPrimary: name => cfb.stream("\u0006DataSpaces", "TransformInfo", name, "\u0006Primary")
		})

		return { encryptionInfo, encryptedPackage, dataSpaces }
	}

	for (const sample of samples) {
		describe.skipIf(!existsSync(join(CORPUS_DIR, sample.path)))(sample.path, () => {
			it("decrypts to the known package, verifying the HMAC where there is one", { timeout: 30_000 }, async () => {
				const streams = streamsOf(sample.path)

				expect(streams.dataSpaces).toBe(sample.dataSpaces)
				expect(probeEncryptionInfo(streams.encryptionInfo)).toMatchObject({ ok: true, scheme: sample.scheme })

				const result = unlocked(await unlock(streams, sample.password))

				expect(result.key.scheme).toBe(sample.scheme)
				expect(result.integrity).toBe(sample.scheme === "agile" ? "verified" : "unchecked")
				expect(result.plaintext.length).toBe(sample.size)
				expect(sha256(result.plaintext)).toBe(sample.sha)
				releaseOfficeKey(result.key)
			})

			it("refuses a wrong password", { timeout: 30_000 }, async () => {
				const streams = streamsOf(sample.path)

				expect(await reason(unlock(streams, `${sample.password}x`))).toBe("wrongPassword")
				expect(await reason(unlock(streams, ""))).toBe("wrongPassword")
			})

			it("seals as Agile and reopens with the same password", { timeout: 30_000 }, async () => {
				const streams = streamsOf(sample.path)
				const source = unlocked(await unlock(streams, sample.password))
				const sealed = await sealOfficePackage(source.key, source.plaintext)
				const reopened = unlocked(await unlockOfficePackage(sealed.encryptionInfo, sealed.encryptedPackage, sample.password))

				expect(reopened.integrity).toBe("verified")
				expect(sha256(reopened.plaintext)).toBe(sample.sha)
				expect(sha256(await openSealedPackage(source.key, sealed.encryptionInfo, sealed.encryptedPackage))).toBe(sample.sha)
				expect(sha256(decryptAgileStreams(sealed, sample.password))).toBe(sample.sha)

				const info = agileOf(sealed.encryptionInfo)

				if (sample.scheme === "agile") {
					// The source's parameters and password key encryptor are kept; only keyData's salt is new.
					const original = agileOf(streams.encryptionInfo)

					expect(info.password).toEqual(original.password)
					expect({ ...info.keyData, salt: null }).toEqual({ ...original.keyData, salt: null })
					expect(info.keyData.salt).not.toEqual(original.keyData.salt)
				} else {
					expect(info.keyData).toMatchObject({ keyBits: 256, hash: "SHA-512", hashSize: 64, saltSize: 16 })
					expect(info.password).toMatchObject({ keyBits: 256, hash: "SHA-512", spinCount: 100 })
				}

				releaseOfficeKey(source.key)
				releaseOfficeKey(reopened.key)
			})
		})
	}

	const agileSample = samples[1]

	describe.skipIf(!existsSync(join(CORPUS_DIR, agileSample.path)))("data integrity", () => {
		it("fails when any one byte of the EncryptedPackage stream changes, StreamSize included", { timeout: 60_000 }, async () => {
			const streams = streamsOf(agileSample.path)

			for (const at of [0, 8, 8 + 4096, streams.encryptedPackage.length - 1]) {
				const tampered = streams.encryptedPackage.slice()

				tampered[at] = (tampered[at] ?? 0) ^ 1

				expect(await reason(unlockOfficePackage(streams.encryptionInfo, tampered, agileSample.password))).toBe("integrity")
			}
		})
	})
})

describe("office encryption: independent generator round trips", () => {
	const plaintext = (size: number): Uint8Array => new Uint8Array(randomBytes(size))

	it("opens Agile packages at every boundary size and seals them back", { timeout: 60_000 }, async () => {
		for (const size of SIZES) {
			const data = plaintext(size)
			const streams = agileStreams(data, "secret", { hash: "sha256", keyBits: 128, spinCount: 10 })
			const source = unlocked(await unlock(streams, "secret"))

			expect(source.integrity).toBe("verified")
			expect(Buffer.from(source.plaintext).equals(Buffer.from(data))).toBe(true)

			const sealed = await sealOfficePackage(source.key, source.plaintext)

			expect(sealed.encryptedPackage.length).toBe(8 + Math.ceil(size / 16) * 16)
			expect(Buffer.from(decryptAgileStreams(sealed, "secret")).equals(Buffer.from(data))).toBe(true)
			expect(
				Buffer.from(await openSealedPackage(source.key, sealed.encryptionInfo, sealed.encryptedPackage)).equals(Buffer.from(data))
			).toBe(true)
			releaseOfficeKey(source.key)
		}
	})

	it("opens Standard packages and seals them as Agile AES-256 / SHA-512", { timeout: 60_000 }, async () => {
		for (const size of [0, 5, 4096, 256 * 4096 + 33]) {
			const data = plaintext(size)
			const source = unlocked(await unlock(standardStreams(data, "pw"), "pw"))

			expect(Buffer.from(source.plaintext).equals(Buffer.from(data))).toBe(true)

			const sealed = await sealOfficePackage(source.key, source.plaintext)

			expect(Buffer.from(decryptAgileStreams(sealed, "pw")).equals(Buffer.from(data))).toBe(true)
			releaseOfficeKey(source.key)
		}
	})

	it("opens AES-256 Standard packages", { timeout: 30_000 }, async () => {
		const data = plaintext(1000)
		const source = unlocked(await unlock(standardStreams(data, "pw", 256), "pw"))

		expect(Buffer.from(source.plaintext).equals(Buffer.from(data))).toBe(true)
	})

	it("opens every supported hash and key size", { timeout: 60_000 }, async () => {
		for (const hash of ["sha1", "sha256", "sha384", "sha512"] as const) {
			for (const keyBits of [128, 256] as const) {
				const data = plaintext(5000)
				const source = unlocked(await unlock(agileStreams(data, "k", { hash, keyBits, spinCount: 3, saltSize: 24 }), "k"))

				expect(Buffer.from(source.plaintext).equals(Buffer.from(data))).toBe(true)

				const sealed = await sealOfficePackage(source.key, source.plaintext)

				expect(agileOf(sealed.encryptionInfo).keyData).toMatchObject({ keyBits, saltSize: 24 })
				expect(Buffer.from(decryptAgileStreams(sealed, "k")).equals(Buffer.from(data))).toBe(true)
			}
		}
	})

	it("writes Word's defaults for a Standard source: spin 100,000", { timeout: 60_000 }, async () => {
		const data = plaintext(100)
		const streams = standardStreams(data, "pw")
		const source = unlocked(await unlockOfficePackage(streams.encryptionInfo, streams.encryptedPackage, "pw"))
		const sealed = await sealOfficePackage(source.key, source.plaintext)
		const info = agileOf(sealed.encryptionInfo)

		expect(info.password.spinCount).toBe(AGILE_DEFAULT_SPIN_COUNT)
		expect(info.keyData).toMatchObject({ keyBits: 256, hash: "SHA-512", saltSize: 16, hashSize: 64 })
		expect(Buffer.from(decryptAgileStreams(sealed, "pw")).equals(Buffer.from(data))).toBe(true)
	})

	it("uses fresh salts, IVs and HMAC keys on every seal, keeping the key encryptor", { timeout: 30_000 }, async () => {
		const data = plaintext(10_000)
		const source = unlocked(await unlock(agileStreams(data, "pw", { spinCount: 5 }), "pw"))
		const a = await sealOfficePackage(source.key, data)
		const b = await sealOfficePackage(source.key, data)
		const infoA = agileOf(a.encryptionInfo)
		const infoB = agileOf(b.encryptionInfo)

		expect(infoA.password).toEqual(infoB.password)
		expect(infoA.keyData.salt).not.toEqual(infoB.keyData.salt)
		expect(infoA.dataIntegrity?.encryptedHmacKey).not.toEqual(infoB.dataIntegrity?.encryptedHmacKey)
		expect(Buffer.from(a.encryptedPackage.subarray(8, 24)).equals(Buffer.from(b.encryptedPackage.subarray(8, 24)))).toBe(false)
	})

	it("keeps certificate key encryptors whose certVerifier is right and discards the rest (§2.3.4.10)", { timeout: 30_000 }, async () => {
		const data = plaintext(100)
		const streams = agileStreams(data, "pw", {
			spinCount: 5,
			certificates: [
				{ certificate: Uint8Array.of(1, 2, 3), validVerifier: true },
				{ certificate: Uint8Array.of(4, 5, 6), validVerifier: false }
			]
		})
		const source = unlocked(await unlock(streams, "pw"))
		const sealed = await sealOfficePackage(source.key, data)
		const kept = agileOf(sealed.encryptionInfo).certificates

		expect(agileOf(streams.encryptionInfo).certificates).toHaveLength(2)
		expect(kept).toEqual([agileOf(streams.encryptionInfo).certificates[0]])
	})

	it("treats passwords as UTF-16 code units without normalising them", { timeout: 30_000 }, async () => {
		const nfc = "pässwörd"
		const nfd = nfc.normalize("NFD")

		for (const password of [nfc, "\u{1f512}", "x".repeat(255), "a"]) {
			const streams = agileStreams(plaintext(50), password, { spinCount: 5 })

			expect(await reason(unlock(streams, password))).toBe("ok")
		}

		expect(await reason(unlock(agileStreams(plaintext(50), nfc, { spinCount: 5 }), nfd))).toBe("wrongPassword")
		expect(await reason(unlock(standardStreams(plaintext(50), "\u{1f512}"), "\u{1f512}"))).toBe("ok")
	})

	it("opens an Agile package without data integrity, reporting it unchecked", async () => {
		const data = plaintext(3000)
		const result = unlocked(await unlock(agileStreams(data, "pw", { spinCount: 5, integrity: false }), "pw"))

		expect(result.integrity).toBe("unchecked")
		expect(Buffer.from(result.plaintext).equals(Buffer.from(data))).toBe(true)
	})
})

describe("office encryption: keys and cancellation", () => {
	it("refuses to seal with a released key", async () => {
		const source = unlocked(await unlock(agileStreams(new Uint8Array(10), "pw", { spinCount: 5 }), "pw"))

		releaseOfficeKey(source.key)
		expect(await rejection(sealOfficePackage(source.key, new Uint8Array(10)))).toBe("aborted")
	})

	it("stops a Standard upgrade's background spin on release", { timeout: 30_000 }, async () => {
		const source = unlocked(await unlock(standardStreams(new Uint8Array(10), "pw"), "pw", 10_000_000))
		const seal = sealOfficePackage(source.key, source.plaintext)

		releaseOfficeKey(source.key)
		expect(await rejection(seal)).toBe("aborted")
	})

	it("aborts the spin when the caller's signal fires", { timeout: 30_000 }, async () => {
		const streams = agileStreams(new Uint8Array(10), "pw", { spinCount: 5 })
		const info = Buffer.from(streams.encryptionInfo).toString("latin1").replace('spinCount="5"', 'spinCount="10000000"')
		const signal = { aborted: false }
		const pending = unlockOfficePackage(Buffer.from(info, "latin1"), streams.encryptedPackage, "pw", { signal })

		signal.aborted = true
		expect(await reason(pending)).toBe("aborted")
	})

	it("gives the host turns through its slicing during the spin and the segments", { timeout: 30_000 }, async () => {
		let pauses = 0
		const slicing = { sliceMs: 0, now: () => 0, pause: () => (pauses++, Promise.resolve()) }
		const streams = agileStreams(new Uint8Array(3 * 256 * 4096), "pw", { spinCount: 1000 })
		const source = unlocked(await unlockOfficePackage(streams.encryptionInfo, streams.encryptedPackage, "pw", { slicing }))

		// 1,000 / 64 spin turns plus one per 1 MiB window.
		expect(pauses).toBeGreaterThanOrEqual(15 + 3)

		const before = pauses

		await sealOfficePackage(source.key, source.plaintext, { slicing })
		expect(pauses - before).toBeGreaterThanOrEqual(3)
	})

	it("returns opaque keys whose state never leaves this module", async () => {
		const source = unlocked(await unlock(agileStreams(new Uint8Array(10), "pw", { spinCount: 5 }), "pw"))
		const key: OfficeKey = source.key

		expect(Object.keys(key)).toEqual(["scheme"])
		expect(Object.isFrozen(key)).toBe(true)
		expect(JSON.stringify(key)).toBe('{"scheme":"agile"}')
	})
})

describe("office encryption: refusals", () => {
	const valid = agileStreams(new Uint8Array(100), "pw", { spinCount: 5 })
	const xml = Buffer.from(valid.encryptionInfo.subarray(8)).toString("utf8")
	const withXml = (edit: (text: string) => string): Uint8Array => agileInfo(edit(xml))
	const probe = (info: Uint8Array): string => {
		const result = probeEncryptionInfo(info)

		return result.ok ? result.scheme : result.reason
	}
	const versioned = (major: number, minor: number, flags: number): Uint8Array => {
		const standard = standardStreams(new Uint8Array(1), "x").encryptionInfo.slice()
		const view = new DataView(standard.buffer)

		view.setUint16(0, major, true)
		view.setUint16(2, minor, true)
		view.setUint32(12, flags, true)

		return standard
	}

	it("classifies schemes by version and flags", () => {
		expect(probe(valid.encryptionInfo)).toBe("agile")
		expect(probe(versioned(3, 2, 0x24))).toBe("standard")
		expect(probe(versioned(4, 3, 0x10))).toBe("unsupported")
		expect(probe(versioned(4, 2, 0x34))).toBe("unsupported")
		expect(probe(versioned(4, 2, 0x04))).toBe("unsupported")
		expect(probe(versioned(1, 1, 0x24))).toBe("unsupported")
		expect(probe(new Uint8Array([4, 0]))).toBe("corrupt")
	})

	it("refuses Agile parameters outside AES-CBC with a known hash", () => {
		expect(probe(withXml(t => t.replaceAll("ChainingModeCBC", "ChainingModeCFB")))).toBe("unsupported")
		expect(probe(withXml(t => t.replaceAll('cipherAlgorithm="AES"', 'cipherAlgorithm="3DES"')))).toBe("unsupported")
		expect(probe(withXml(t => t.replaceAll('hashAlgorithm="SHA512"', 'hashAlgorithm="MD5"')))).toBe("unsupported")
		expect(probe(withXml(t => t.replaceAll('keyBits="256"', 'keyBits="64"')))).toBe("unsupported")
		expect(probe(withXml(t => t.replaceAll('hashSize="64"', 'hashSize="32"')))).toBe("corrupt")
		expect(probe(withXml(t => t.replaceAll('blockSize="16"', 'blockSize="8"')))).toBe("corrupt")
		expect(probe(withXml(t => t.replace('spinCount="5"', 'spinCount="10000001"')))).toBe("corrupt")
		expect(probe(withXml(t => t.replace('spinCount="5"', 'spinCount="10000000"')))).toBe("agile")
		expect(probe(withXml(t => t.replace('spinCount="5"', 'spinCount="-1"')))).toBe("corrupt")
		expect(probe(withXml(t => t.replace(/saltValue="[^"]*"/u, 'saltValue="!!!!"')))).toBe("corrupt")
	})

	it("reads both SHA-1 spellings and writes SHA1, the one Office, POI and msoffcrypto-tool read", { timeout: 30_000 }, async () => {
		const data = new Uint8Array(randomBytes(100))
		const streams = agileStreams(data, "pw", { hash: "sha1", spinCount: 5 })
		const hyphenated = agileInfo(
			Buffer.from(streams.encryptionInfo.subarray(8)).toString("utf8").replaceAll('hashAlgorithm="SHA1"', 'hashAlgorithm="SHA-1"')
		)
		const source = unlocked(await unlockOfficePackage(hyphenated, streams.encryptedPackage, "pw"))
		const sealed = await sealOfficePackage(source.key, data)
		const xml = Buffer.from(sealed.encryptionInfo.subarray(8)).toString("utf8")

		expect(xml).toContain('hashAlgorithm="SHA1"')
		expect(xml).not.toContain("SHA-1")
	})

	it("refuses a certificate-only document as unsupported", () => {
		const certificateOnly = withXml(t =>
			t.replace(
				/<keyEncryptor uri="http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/password">.*<\/keyEncryptor>/u,
				'<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><c:encryptedKey encryptedKeyValue="AAAA" X509Certificate="AQID" certVerifier="BAUG"/></keyEncryptor>'
			)
		)

		expect(probe(certificateOnly)).toBe("unsupported")
	})

	it("refuses malformed XML and DTD subsets without expanding anything", () => {
		expect(probe(withXml(t => t.replace("</encryption>", "")))).toBe("corrupt")
		expect(probe(withXml(t => t.replace("<encryption", '<!DOCTYPE encryption [<!ENTITY a "aaaaaaaaaa">]><encryption')))).toBe("corrupt")
		expect(
			probe(withXml(t => t.replace("<encryption", '<!DOCTYPE encryption SYSTEM "http://example.invalid/x.dtd"><encryption')))
		).toBe("agile")
		expect(probe(agileInfo("<other/>"))).toBe("corrupt")
		expect(probe(new Uint8Array(2 * 1024 * 1024).fill(4))).toBe("corrupt")
	})

	it("refuses truncated or oversized package streams as corrupt", async () => {
		const header = (size: bigint, body: number): Uint8Array => {
			const out = new Uint8Array(8 + body)

			new DataView(out.buffer).setBigUint64(0, size, true)

			return out
		}

		for (const stream of [new Uint8Array(4), header(100n, 99), header(100n, 100), header(2n ** 63n, 16)]) {
			expect(await reason(unlockOfficePackage(valid.encryptionInfo, stream, "pw"))).toBe("corrupt")
		}
	})
})

describe("office encryption: data spaces", () => {
	const lp = (text: string): Buffer => {
		const data = Buffer.from(text, "utf16le")
		const out = Buffer.alloc(4 + data.length + ((4 - (data.length % 4)) % 4))

		out.writeUInt32LE(data.length)
		data.copy(out, 4)

		return out
	}
	const u32 = (value: number): Buffer => {
		const out = Buffer.alloc(4)

		out.writeUInt32LE(value)

		return out
	}
	const map = (entries: readonly (readonly [string, string])[]): Uint8Array => {
		const body = entries.map(([stream, space]) => {
			const rest = Buffer.concat([u32(1), u32(0), lp(stream), lp(space)])

			return Buffer.concat([u32(rest.length + 4), rest])
		})

		return Buffer.concat([u32(8), u32(entries.length), ...body])
	}
	const definition = (...transforms: string[]): Uint8Array => Buffer.concat([u32(8), u32(transforms.length), ...transforms.map(lp)])
	const primary = (id: string): Uint8Array =>
		Buffer.concat([u32(0), u32(1), lp(id), lp("name"), u32(0x10001), u32(0x10001), u32(0x10001)])
	const source = (dataSpaceMap: Uint8Array | undefined, spaces: Record<string, Uint8Array>, transforms: Record<string, Uint8Array>) => ({
		dataSpaceMap,
		dataSpaceDefinition: (name: string) => spaces[name],
		transformPrimary: (name: string) => transforms[name]
	})
	const strong = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"
	const drm = "{C73DFACD-061F-43B0-8B64-0C620D2A8B50}"

	it("tells password encryption from IRM and unknown transforms", () => {
		expect(
			classifyDataSpaces(
				source(
					map([["EncryptedPackage", "StrongEncryptionDataSpace"]]),
					{ StrongEncryptionDataSpace: definition("StrongEncryptionTransform") },
					{ StrongEncryptionTransform: primary(strong.toLowerCase()) }
				)
			)
		).toBe("password")
		expect(classifyDataSpaces(source(map([["EncryptedPackage", "Space"]]), { Space: definition("T") }, { T: primary(drm) }))).toBe(
			"rightsManaged"
		)
		expect(
			classifyDataSpaces(
				source(
					map([
						["EncryptedPackage", "StrongEncryptionDataSpace"],
						["\u0009DRMContent", "DRMEncryptedDataSpace"]
					]),
					{ StrongEncryptionDataSpace: definition("StrongEncryptionTransform") },
					{ StrongEncryptionTransform: primary(strong) }
				)
			)
		).toBe("rightsManaged")
		expect(
			classifyDataSpaces(
				source(
					map([["EncryptedPackage", "Space"]]),
					{ Space: definition("T") },
					{ T: primary("{00000000-0000-0000-0000-000000000000}") }
				)
			)
		).toBe("unsupported")
		expect(classifyDataSpaces(source(map([["EncryptedPackage", "Missing"]]), {}, {}))).toBe("unsupported")
		expect(classifyDataSpaces(source(undefined, {}, {}))).toBe("absent")
	})

	it("refuses truncated and oversized structures with typed errors", () => {
		const whole = map([["EncryptedPackage", "S"]])

		for (const cut of [2, 6, 10, 20, whole.length - 2]) {
			expect(() => classifyDataSpaces(source(whole.subarray(0, cut), {}, {}))).toThrow(OfficeCryptoError)
		}

		expect(() => classifyDataSpaces(source(Buffer.concat([u32(8), u32(1_000_000)]), {}, {}))).toThrow(OfficeCryptoError)
		expect(() => classifyDataSpaces(source(Buffer.concat([u32(8), u32(1), u32(16), u32(1), u32(0), u32(0xfffffff0)]), {}, {}))).toThrow(
			OfficeCryptoError
		)
	})
})

describe("office encryption: base64", () => {
	it("round-trips and refuses malformed input", () => {
		for (let length = 0; length < 40; length++) {
			const bytes = new Uint8Array(randomBytes(length))
			const text = encodeBase64(bytes)

			expect(text).toBe(Buffer.from(bytes).toString("base64"))
			expect(decodeBase64(text)).toEqual(bytes)
			expect(decodeBase64(` ${text.slice(0, 2)}\r\n${text.slice(2)} `)).toEqual(bytes)
		}

		for (const bad of ["A", "AB=", "A===", "AB=C", "AB*=", "éAAA", "AAAA="]) expect(() => decodeBase64(bad)).toThrow(OfficeCryptoError)
	})
})
