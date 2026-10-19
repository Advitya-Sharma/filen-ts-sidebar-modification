import { randomBytes } from "node:crypto"
import { afterEach, describe, expect, it } from "vitest"
import { cfbBytes } from "@filen/office/core/cfb"
import { bytesSource, type ByteSource } from "@filen/office/core/opc"
import { passwordDataSpaces } from "../core/encryption/index"
import {
	encryptedPackageVerdict,
	openEncryptedPackage,
	parseEncryptionInfo,
	probeEncryptionInfo,
	releaseOfficeKey,
	unlockOfficePackage,
	type OfficeUnlockResult
} from "../dom/encryption/index"
import { packageSealer, sealOfficePackage } from "../dom/encryption/seal"
import { agileInfo, agileStreams, decryptAgileStreams, standardStreams, type EncryptedStreams } from "./encryption/generate"

// Agile and Standard encryption edge cases: parameter checks, integrity, key encryptors, zero-filled secrets,
// when spins run, and what classifying reads ([MS-OFFCRYPTO] §2.3.4.10-§2.3.4.15).

function ok(result: OfficeUnlockResult): Extract<OfficeUnlockResult, { ok: true }> {
	if (!result.ok) throw new Error(`unlock failed: ${result.reason} ${result.message}`)

	return result
}

const xmlOf = (info: Uint8Array): string => Buffer.from(info.subarray(8)).toString("utf8")

const fileOf = (streams: EncryptedStreams): Promise<Uint8Array> =>
	cfbBytes([
		...passwordDataSpaces(),
		{ path: ["EncryptionInfo"], bytes: streams.encryptionInfo },
		{ path: ["EncryptedPackage"], bytes: streams.encryptedPackage }
	])

// Counts subtle.digest calls by algorithm, and records each input.
function watchDigest(): { readonly counts: Map<string, number>; readonly inputs: Uint8Array[] } {
	const subtle = globalThis.crypto.subtle
	const original = subtle.digest.bind(subtle)
	const counts = new Map<string, number>()
	const inputs: Uint8Array[] = []

	Object.defineProperty(subtle, "digest", {
		configurable: true,
		writable: true,
		value: (algorithm: string | { readonly name: string }, data: Uint8Array) => {
			const name = typeof algorithm === "string" ? algorithm : algorithm.name

			counts.set(name, (counts.get(name) ?? 0) + 1)
			inputs.push(data)

			return original(algorithm, data)
		}
	})

	return { counts, inputs }
}

afterEach(() => {
	Reflect.deleteProperty(globalThis.crypto.subtle, "digest")
	Reflect.deleteProperty(globalThis.crypto.subtle, "sign")
	Reflect.deleteProperty(globalThis.crypto.subtle, "verify")
})

describe("saltSize and the password encryptor's algorithms are checked against keyData", () => {
	it("refuses a keyData saltSize that does not match its saltValue (the only unauthenticated seal parameter)", () => {
		const streams = agileStreams(new Uint8Array(100), "pw", { spinCount: 5 })

		for (const size of ["1", "65536"]) {
			const tampered = agileInfo(xmlOf(streams.encryptionInfo).replace('saltSize="16"', `saltSize="${size}"`))

			expect(probeEncryptionInfo(tampered)).toEqual({ ok: false, reason: "corrupt" })
		}
	})

	it("refuses a password key encryptor whose hash is not keyData's", () => {
		const streams = agileStreams(new Uint8Array(100), "pw", { spinCount: 5, hash: "sha512" })
		const xml = xmlOf(streams.encryptionInfo).replace(
			/(<keyData [^>]*?)hashSize="64"([^>]*?)hashAlgorithm="SHA512"/u,
			'$1hashSize="32"$2hashAlgorithm="SHA256"'
		)

		expect(probeEncryptionInfo(agileInfo(xml))).toEqual({ ok: false, reason: "corrupt" })
	})

	it("seals with at least 16 salt bytes, so the IVs of two saves never repeat", { timeout: 30_000 }, async () => {
		const data = new Uint8Array(randomBytes(8192))
		const source = ok(await unlock(agileStreams(data, "pw", { spinCount: 5, saltSize: 1 }), "pw"))
		const sealed = await sealOfficePackage(source.key, data)
		const info = parseEncryptionInfo(sealed.encryptionInfo)

		expect(info.kind === "agile" && [info.keyData.saltSize, info.keyData.salt.length]).toEqual([16, 16])
		expect(Buffer.from(decryptAgileStreams(sealed, "pw")).equals(Buffer.from(data))).toBe(true)
		releaseOfficeKey(source.key)
	})
})

describe("a package without dataIntegrity opens, reported unchecked", () => {
	it("opens a stripped, changed package as unchecked (the app warns), and refuses the change while the HMAC is there", async () => {
		const data = new Uint8Array(randomBytes(5000))
		const streams = agileStreams(data, "pw", { spinCount: 5 })
		const flipped = streams.encryptedPackage.slice()

		flipped[8 + 100] = (flipped[8 + 100] ?? 0) ^ 0x01

		const caught = await unlockOfficePackage(streams.encryptionInfo, flipped, "pw")

		expect(caught.ok ? "ok" : caught.reason).toBe("integrity")

		const stripped = agileInfo(xmlOf(streams.encryptionInfo).replace(/<dataIntegrity [^>]*\/>/u, ""))
		const opened = await openEncryptedPackage(bytesSource(await fileOf({ encryptionInfo: stripped, encryptedPackage: flipped })))

		if (opened.kind !== "needsPassword") throw new Error(opened.kind)

		const result = await opened.locked.unlock("pw")

		expect(result.ok && result.document.integrity).toBe("unchecked")
		if (result.ok) result.document.release()
	})
})

describe("certificate key encryptors are discarded unless their certVerifier is right", () => {
	it("drops one whose certVerifier is not an HMAC of its certificate", { timeout: 30_000 }, async () => {
		const certificate =
			'<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><c:encryptedKey encryptedKeyValue="AAAA" X509Certificate="AQID" certVerifier="BAUG"/></keyEncryptor>'
		const streams = agileStreams(new Uint8Array(100), "pw", { spinCount: 5, extraEncryptors: certificate })
		const source = ok(await unlock(streams, "pw"))
		const info = parseEncryptionInfo((await sealOfficePackage(source.key, source.plaintext)).encryptionInfo)

		expect(info.kind === "agile" && info.certificates.length).toBe(0)
		releaseOfficeKey(source.key)
	})
})

describe("password-derived temporaries are zero-filled", () => {
	it("leaves no Hn + blockKey, Standard Hn + LE32(0) or password bytes behind once unlocked", { timeout: 60_000 }, async () => {
		const agile = agileStreams(new Uint8Array(10), "pw", { spinCount: 3, hash: "sha512" })
		const standard = standardStreams(new Uint8Array(10), "pw")
		const { inputs } = watchDigest()
		const a = ok(await unlock(agile, "pw"))
		const s = ok(await unlockOfficePackage(standard.encryptionInfo, standard.encryptedPackage, "pw", { upgradeSpinCount: 1 }))

		releaseOfficeKey(a.key)
		releaseOfficeKey(s.key)

		// Every input whose bytes could be password-derived (all but the ciphertext-sized IV inputs, salt + LE32
		// and salt + blockKey) is zero now.
		const live = inputs.filter(
			bytes => bytes.length !== 16 + 4 && bytes.length !== 16 + 8 && bytes.length !== 16 && bytes.some(byte => byte !== 0)
		)

		expect(live.map(bytes => bytes.length)).toEqual([])
	})

	it("validates upgradeSpinCount before any key work", async () => {
		const standard = standardStreams(new Uint8Array(10), "pw")
		const { counts } = watchDigest()

		await expect(
			unlockOfficePackage(standard.encryptionInfo, standard.encryptedPackage, "pw", { upgradeSpinCount: -1 })
		).rejects.toThrow(RangeError)
		expect(counts.size).toBe(0)
	})
})

describe("a Standard source's upgrade spin waits for a seal", () => {
	it("takes only H0 at unlock: a session that never seals runs no second spin", { timeout: 60_000 }, async () => {
		const opened = await openEncryptedPackage(bytesSource(await fileOf(standardStreams(new Uint8Array(5000).fill(7), "pw"))))

		if (opened.kind !== "needsPassword") throw new Error(opened.kind)

		const result = await opened.locked.unlock("pw")

		if (!result.ok) throw new Error(result.reason)

		const { counts } = watchDigest()

		await new Promise(resolve => setTimeout(resolve, 300))
		expect(counts.get("SHA-512") ?? 0).toBe(0)

		// The first seal derives the encryptor (100,000 spins).
		const { sealer, release } = packageSealer(result.document.key)

		await sealer.seal(result.document.bytes)
		expect(counts.get("SHA-512") ?? 0).toBeGreaterThanOrEqual(100_000)
		release()
	})

	it("stops a seal waiting on the spin at its signal, and a later seal derives it again", { timeout: 60_000 }, async () => {
		const data = new Uint8Array(3000).fill(3)
		const source = ok(await unlockOfficePackage(...streamsOf(standardStreams(data, "pw")), "pw", { upgradeSpinCount: 300_000 }))
		const signal = { aborted: false }
		const pending = sealOfficePackage(source.key, data, { signal })

		setTimeout(() => (signal.aborted = true), 20)
		await expect(pending).rejects.toMatchObject({ code: "aborted" })

		const sealed = await sealOfficePackage(source.key, data)

		expect(Buffer.from(decryptAgileStreams(sealed, "pw")).equals(Buffer.from(data))).toBe(true)
		releaseOfficeKey(source.key)
	})

	it("runs the spin on the caller's slicing clock", { timeout: 60_000 }, async () => {
		const data = new Uint8Array(100)
		const source = ok(await unlockOfficePackage(...streamsOf(standardStreams(data, "pw")), "pw", { upgradeSpinCount: 2000 }))
		let pauses = 0

		await sealOfficePackage(source.key, data, {
			slicing: { sliceMs: 0, now: () => performance.now(), pause: () => (pauses++, Promise.resolve()) }
		})
		// A turn every 64 spins at least.
		expect(pauses).toBeGreaterThanOrEqual(2000 / 64)
		releaseOfficeKey(source.key)
	})
})

describe("concurrent first seals of a Standard key wait on one upgrade, each on its own signal", () => {
	it("aborting one caller stops only that caller; the other seals", { timeout: 60_000 }, async () => {
		const data = new Uint8Array(10)
		const source = ok(await unlockOfficePackage(...streamsOf(standardStreams(data, "pw")), "pw", { upgradeSpinCount: 200_000 }))
		const first = { aborted: false }
		const a = sealOfficePackage(source.key, data, { signal: first })
		const b = sealOfficePackage(source.key, data)

		setTimeout(() => (first.aborted = true), 20)
		await expect(a).rejects.toMatchObject({ code: "aborted" })

		const sealed = await b

		expect(Buffer.from(decryptAgileStreams(sealed, "pw")).equals(Buffer.from(data))).toBe(true)
		releaseOfficeKey(source.key)
	})

	it("stops the upgrade itself only at release", { timeout: 60_000 }, async () => {
		const data = new Uint8Array(10)
		const source = ok(await unlockOfficePackage(...streamsOf(standardStreams(data, "pw")), "pw", { upgradeSpinCount: 10_000_000 }))
		const pending = sealOfficePackage(source.key, data)

		setTimeout(() => releaseOfficeKey(source.key), 20)
		await expect(pending).rejects.toMatchObject({ code: "aborted" })
	})
})

describe("the HMAC over the whole package gets a turn before it", () => {
	it("pauses right before HMAC sign in the seal and verify in the check", { timeout: 30_000 }, async () => {
		const data = new Uint8Array(randomBytes(3 * 4096))
		const source = ok(await unlock(agileStreams(data, "pw", { spinCount: 5 }), "pw"))
		const events: string[] = []
		const subtle = globalThis.crypto.subtle

		for (const name of ["sign", "verify"] as const) {
			const original = subtle[name].bind(subtle) as (...args: unknown[]) => Promise<unknown>

			Object.defineProperty(subtle, name, {
				configurable: true,
				writable: true,
				value: (...args: unknown[]) => {
					events.push(name)

					return original(...args)
				}
			})
		}

		const { sealer, release } = packageSealer(source.key)

		await sealer.seal(data, {
			slicing: { sliceMs: 0, now: () => performance.now(), pause: () => (events.push("turn"), Promise.resolve()) }
		})

		expect(events.filter(event => event === "sign")).toHaveLength(1)
		expect(events.filter(event => event === "verify")).toHaveLength(1)
		// Every whole-package HMAC call comes straight after a turn.
		for (const [k, event] of events.entries())
			if (event === "sign" || event === "verify") expect(events[k - 1], `${event} at ${k}`).toBe("turn")
		release()
	})
})

describe("Excel's default password is tried only at Excel's spin", () => {
	it("asks at once for a file whose spin is above 100,000, and tries the default below", { timeout: 60_000 }, async () => {
		const slow = await fileOf(agileStreams(new Uint8Array(100), "VelvetSweatshop", { spinCount: 100_001 }))
		const { counts } = watchDigest()
		const opened = await openEncryptedPackage(bytesSource(slow), { tryDefault: true })

		expect(opened.kind).toBe("needsPassword")
		expect(counts.get("SHA-512") ?? 0).toBe(0)

		const fast = await openEncryptedPackage(
			bytesSource(await fileOf(agileStreams(new Uint8Array(100), "VelvetSweatshop", { spinCount: 1000 }))),
			{
				tryDefault: true
			}
		)

		expect(fast.kind).toBe("unlocked")
		if (fast.kind === "unlocked") fast.document.release()
	})
})

describe("damage inside lazily read Agile XML attributes is refused as damaged", () => {
	const streams = agileStreams(new Uint8Array(100), "pw", { spinCount: 1 })
	const xml = xmlOf(streams.encryptionInfo)
	const cases: Record<string, string> = {
		"namespace declaration without a value": xml.replace(
			'xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"',
			"xmlns:"
		),
		"unterminated reference in an attribute": xml.replace('cipherAlgorithm="AES"', 'cipherAlgorithm="AES&amp"')
	}

	for (const [name, mutated] of Object.entries(cases)) {
		it(`${name}: probe says corrupt, open refuses as damaged`, async () => {
			expect(mutated).not.toBe(xml)

			const info = agileInfo(mutated)
			const file = await fileOf({ encryptionInfo: info, encryptedPackage: streams.encryptedPackage })

			expect(probeEncryptionInfo(info)).toEqual({ ok: false, reason: "corrupt" })
			await expect(encryptedPackageVerdict(bytesSource(file))).resolves.toBe("damaged")
			await expect(openEncryptedPackage(bytesSource(file))).resolves.toEqual({ kind: "refused", reason: "damaged" })
		})
	}
})

describe("classifying reads only the package's StreamSize", () => {
	it("refuses or prompts without reading the encrypted package, which the unlock reads once", { timeout: 60_000 }, async () => {
		const data = new Uint8Array(randomBytes(2 * 1024 * 1024))
		const streams = agileStreams(data, "pw", { spinCount: 5 })
		const file = await fileOf(streams)
		const reads: number[] = []
		const source: ByteSource = {
			size: file.length,
			read(offset, length) {
				reads.push(length)

				return Promise.resolve(file.subarray(offset, offset + length))
			}
		}
		const opened = await openEncryptedPackage(source)
		const classified = reads.reduce((sum, length) => sum + length, 0)

		// Header, FAT, directory, EncryptionInfo, the DataSpaces streams and 8 bytes: a few KiB of 2 MiB.
		expect(classified).toBeLessThan(64 * 1024)

		if (opened.kind !== "needsPassword") throw new Error(opened.kind)

		expect(await opened.locked.unlock("wrong")).toEqual({ ok: false, reason: "wrongPassword" })

		const result = await opened.locked.unlock("pw")
		const total = reads.reduce((sum, length) => sum + length, 0)

		// The stream was read once for both tries.
		expect(total - classified).toBeLessThan(streams.encryptedPackage.length * 1.1)
		expect(result.ok && Buffer.from(result.document.bytes).equals(Buffer.from(data))).toBe(true)
		if (result.ok) result.document.release()
	})

	it("says unsupported for an unknown EncryptionInfo version, whatever its package holds (§2.3.4.1)", async () => {
		// Version 0x0404.0x0404: no scheme of [MS-OFFCRYPTO] (Agile is 4.4, Standard x.2, Extensible x.3).
		const info = new Uint8Array(4096).fill(0x04)
		const damagedPackage = new Uint8Array(8192).fill(0xa5)

		expect(probeEncryptionInfo(info)).toEqual({ ok: false, reason: "unsupported" })
		expect(await openEncryptedPackage(bytesSource(await fileOf({ encryptionInfo: info, encryptedPackage: damagedPackage })))).toEqual({
			kind: "refused",
			reason: "unsupported"
		})
	})

	it("refuses a StreamSize larger than the stream as damaged before any prompt", async () => {
		const streams = agileStreams(new Uint8Array(100), "pw", { spinCount: 5 })
		const lying = streams.encryptedPackage.slice()

		new DataView(lying.buffer).setBigUint64(0, 1n << 40n, true)
		expect(await openEncryptedPackage(bytesSource(await fileOf({ ...streams, encryptedPackage: lying })))).toEqual({
			kind: "refused",
			reason: "damaged"
		})
	})
})

function unlock(streams: EncryptedStreams, password: string): Promise<OfficeUnlockResult> {
	return unlockOfficePackage(streams.encryptionInfo, streams.encryptedPackage, password)
}

function streamsOf(streams: EncryptedStreams): [Uint8Array, Uint8Array] {
	return [streams.encryptionInfo, streams.encryptedPackage]
}
