import { describe, expect, it } from "vitest"
import { CfbFile } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import {
	dataSpaceDefinitionStream,
	dataSpaceMapStream,
	isEncryptedContainer,
	readEncryptedContainer,
	transformPrimaryStream
} from "@filen/office/core/encryption"
import { DocxEngine, type DocxErrorReason, type DocxEvent } from "@filen/office/dom/docx"
import { DocxSealError, docxSealer, isSealedOfficeFile } from "@filen/office/dom/docx/editing"
import { classifyDataSpaces, parseEncryptionInfo } from "@filen/office/dom/encryption"
import { buildDocx, p } from "./docxSynthetic"
import {
	PROTECTION,
	documentXml,
	editAndSave,
	encryptedFile,
	expectSealedUpload,
	lockedOf,
	plaintextOf,
	services,
	unlockFile
} from "./docxEncryptedFixtures"
import { agileInfo, agileStreams, standardStreams } from "@filen/office/testing/encryption"
import { budget } from "./timing"

const openProtected: typeof PROTECTION.open = (source, options) => PROTECTION.open(source, options)

const PLAINTEXT = buildDocx({ body: p("Quarterly figures") + p("Second paragraph") })
const PASSWORD = "pässwörd 🔒"

// Agile at a low spin count: what changes with the count is only the time.
const agileFile = (password = PASSWORD, plaintext = PLAINTEXT): Promise<Uint8Array> =>
	encryptedFile(agileStreams(plaintext, password, { spinCount: 1000 }))

async function engineError(file: Uint8Array): Promise<DocxErrorReason | undefined> {
	const events: DocxEvent[] = []

	await new DocxEngine(bytesSource(file), services(), event => events.push(event)).run()

	const error = events.find(event => event.type === "error")

	return error?.type === "error" ? error.reason : undefined
}

// A Standard EncryptionInfo with the given flags and cipher (§2.3.4.5), verifier bytes arbitrary: refused from
// its header, before any password.
function standardHeader(flags: number, algId: number, keyBits: number): Uint8Array {
	const csp = Uint8Array.from(Array.from("Microsoft Base Cryptographic Provider v1.0\0").flatMap(char => [char.charCodeAt(0), 0]))
	const info = new Uint8Array(12 + 32 + csp.length + 4 + 16 + 16 + 4 + 32)
	const view = new DataView(info.buffer)

	view.setUint16(0, 4, true)
	view.setUint16(2, 2, true)
	view.setUint32(4, flags, true)
	view.setUint32(8, 32 + csp.length, true)
	view.setUint32(12, flags, true)
	view.setUint32(20, algId, true)
	view.setUint32(24, 0x8004, true)
	view.setUint32(28, keyBits, true)
	view.setUint32(32, 1, true)
	info.set(csp, 44)
	view.setUint32(44 + csp.length, 16, true)
	view.setUint32(44 + csp.length + 36, 20, true)

	return info
}

// The IRM layout ([MS-OFFCRYPTO] §2.2.10, §2.2.11): EncryptedPackage through the DRM transform, its licence in
// \x09DRMContent's place, no password key encryptor.
function irmDataSpaces(): { readonly path: readonly string[]; readonly bytes: Uint8Array }[] {
	return [
		{
			path: ["\u0006DataSpaces", "DataSpaceMap"],
			bytes: dataSpaceMapStream([{ components: ["EncryptedPackage"], dataSpace: "DRMEncryptedDataSpace" }])
		},
		{
			path: ["\u0006DataSpaces", "DataSpaceInfo", "DRMEncryptedDataSpace"],
			bytes: dataSpaceDefinitionStream(["DRMEncryptedTransform"])
		},
		{
			path: ["\u0006DataSpaces", "TransformInfo", "DRMEncryptedTransform", "\u0006Primary"],
			bytes: transformPrimaryStream("{C73DFACD-061F-43B0-8B64-0C620D2A8B50}", "Microsoft.Metadata.DRMTransform", false)
		}
	]
}

describe("encrypted docx: classification before any prompt", () => {
	it("says plain for a zip and for a compound file that is not an encrypted package", async () => {
		expect(await openProtected(bytesSource(PLAINTEXT))).toEqual({ kind: "plain" })
		expect(await openProtected(bytesSource(await encryptedFile({ encryptedPackage: new Uint8Array(16) }, [], [])))).toEqual({
			kind: "plain"
		})
		expect(await openProtected(bytesSource(Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)))).toEqual({
			kind: "plain"
		})
	})

	it("asks for the password of an Agile or Standard package, with its spin count", async () => {
		const agile = await lockedOf(await agileFile())

		expect([agile.scheme, agile.spinCount]).toEqual(["agile", 1000])

		const standard = await lockedOf(await encryptedFile(standardStreams(PLAINTEXT, PASSWORD)))

		expect([standard.scheme, standard.spinCount]).toEqual(["standard", 50_000])
	})

	it("refuses RC4 and other non-AES Standard, Extensible and certificate-only packages without a prompt", async () => {
		const refused = async (encryptionInfo: Uint8Array): Promise<unknown> =>
			openProtected(bytesSource(await encryptedFile({ encryptionInfo, encryptedPackage: new Uint8Array(24) })))
		// fCryptoAPI without fAES: CryptoAPI RC4 (AlgID 0x6801), as Office 2007 wrote with a weak policy.
		const rc4 = standardHeader(0x04, 0x6801, 128)
		const des = standardHeader(0x24, 0x6603, 192)
		const extensible = standardHeader(0x14, 0x660e, 128)
		const certificateOnly = agileInfo(
			'<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
				'<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>' +
				'<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><c:encryptedKey encryptedKeyValue="AA==" X509Certificate="AA==" certVerifier="AA=="/></keyEncryptor></keyEncryptors></encryption>'
		)
		const cfbChaining = agileInfo(
			new TextDecoder()
				.decode(agileStreams(PLAINTEXT, PASSWORD, { spinCount: 1 }).encryptionInfo.subarray(8))
				.replaceAll("ChainingModeCBC", "ChainingModeCFB")
		)

		for (const info of [rc4, des, extensible, certificateOnly, cfbChaining])
			expect(await refused(info)).toEqual({ kind: "refused", reason: "unsupported" })

		const version33 = standardHeader(0x24, 0x660e, 128)

		new DataView(version33.buffer).setUint16(2, 3, true)
		expect(await refused(version33)).toEqual({ kind: "refused", reason: "unsupported" })
	})

	it("refuses an IRM package as rights-managed, with or without an EncryptionInfo, and a damaged container", async () => {
		const irm = await encryptedFile(
			{ encryptedPackage: new Uint8Array(64) },
			[{ path: ["\u0009DRMContent"], bytes: new Uint8Array(32) }],
			irmDataSpaces()
		)

		expect(await openProtected(bytesSource(irm))).toEqual({ kind: "refused", reason: "rightsManaged" })

		const withInfo = await encryptedFile(agileStreams(PLAINTEXT, PASSWORD, { spinCount: 1 }), [], irmDataSpaces())

		expect(await openProtected(bytesSource(withInfo))).toEqual({ kind: "refused", reason: "rightsManaged" })

		const noInfo = await encryptedFile({ encryptedPackage: new Uint8Array(64) })

		expect(await openProtected(bytesSource(noInfo))).toEqual({ kind: "refused", reason: "damaged" })

		const badMap = await encryptedFile(
			agileStreams(PLAINTEXT, PASSWORD, { spinCount: 1 }),
			[],
			[{ path: ["\u0006DataSpaces", "DataSpaceMap"], bytes: Uint8Array.of(9, 0, 0, 0, 1, 0, 0, 0) }]
		)

		expect(await openProtected(bytesSource(badMap))).toEqual({ kind: "refused", reason: "damaged" })
	})

	it("reports each kind from the engine with a typed reason, never as an open failure", async () => {
		expect(await engineError(await agileFile())).toBe("needsPassword")
		expect(await engineError(await encryptedFile({ encryptedPackage: new Uint8Array(64) }, [], irmDataSpaces()))).toBe("rightsManaged")
		expect(
			await engineError(
				await encryptedFile({ encryptionInfo: standardHeader(0x04, 0x6801, 128), encryptedPackage: new Uint8Array(24) })
			)
		).toBe("unsupportedEncryption")
		expect(await engineError(await encryptedFile({ encryptedPackage: new Uint8Array(64) }))).toBe("damagedEncryption")
	})
})

describe("encrypted docx: unlock", () => {
	it("asks for a document's password without trying Excel's default first", async () => {
		const file = await encryptedFile(agileStreams(PLAINTEXT, "VelvetSweatshop", { spinCount: 1000 }))
		const opened = await openProtected(bytesSource(file))

		expect(opened.kind).toBe("needsPassword")
	})

	it("returns wrongPassword for a wrong password and opens on a retry", async () => {
		const locked = await lockedOf(await agileFile())

		expect(await locked.unlock("pässwörd")).toEqual({ ok: false, reason: "wrongPassword" })
		// NFC and NFD are different passwords: the code units are hashed as they are.
		expect(await locked.unlock(PASSWORD.normalize("NFD"))).toEqual({ ok: false, reason: "wrongPassword" })

		const result = await locked.unlock(PASSWORD)

		if (!result.ok) throw new Error(result.reason)

		expect((await result.document.take()).bytes).toEqual(PLAINTEXT)
		expect(result.document.integrity).toBe("verified")
		expect(result.document.bytes.size).toBe(PLAINTEXT.length)
		result.document.release()
	})

	it("stops an unlock at the caller's signal, the locked document left to retry", { timeout: 30_000 }, async () => {
		const locked = await lockedOf(await encryptedFile(agileStreams(PLAINTEXT, PASSWORD, { spinCount: 200_000 })))
		const signal = { aborted: false }
		const started = performance.now()
		const pending = locked.unlock(PASSWORD, { signal })

		setTimeout(() => (signal.aborted = true), 20)

		expect(await pending).toEqual({ ok: false, reason: "aborted" })
		expect(performance.now() - started).toBeLessThan(budget(1000))

		const aborted = { aborted: true }

		expect(await openProtected(bytesSource(await agileFile()), { signal: aborted })).toEqual({ kind: "aborted" })

		const result = await locked.unlock(PASSWORD)

		expect(result.ok).toBe(true)
		if (result.ok) result.document.release()
	})

	it("refuses a package changed after it was sealed (its HMAC) once the password verified", async () => {
		const streams = agileStreams(PLAINTEXT, PASSWORD, { spinCount: 1000 })
		const tampered = streams.encryptedPackage.slice()

		tampered[tampered.length - 1] = (tampered[tampered.length - 1] ?? 0) ^ 1

		const locked = await lockedOf(await encryptedFile({ ...streams, encryptedPackage: tampered }))

		expect(await locked.unlock(PASSWORD)).toEqual({ ok: false, reason: "damaged" })
	})
})

describe("encrypted docx: save", () => {
	it("seals every save as an Agile compound file that decrypts to what was written, unedited byte for byte", async () => {
		const document = await unlockFile(await agileFile(), PASSWORD)
		const { unedited, edited } = await editAndSave(document, "Sealed edit ")

		for (const written of [unedited, edited]) {
			if (written === undefined) throw new Error("no edited save")

			expect(written.encrypted).toBe(true)
			expectSealedUpload(written.bytes)
			expect(isSealedOfficeFile(written.bytes)).toBe(true)
			expect(written.timings.sealMs).toBeGreaterThanOrEqual(0)
		}

		// Ciphertext differs by design (fresh salts); the package inside is the source's, byte for byte.
		expect(unedited.bytes).not.toEqual(await agileFile())
		expect(await plaintextOf(unedited.bytes, PASSWORD)).toEqual(PLAINTEXT)
		expect(documentXml(await plaintextOf(edited?.bytes ?? new Uint8Array(0), PASSWORD))).toContain("Sealed edit Quarterly figures")
		document.release()
	})

	it("writes the compound file Word writes: v3, the password DataSpaces, EncryptionInfo and EncryptedPackage only", async () => {
		const source = await encryptedFile(agileStreams(PLAINTEXT, PASSWORD, { spinCount: 1000, hash: "sha256", keyBits: 128 }), [
			{ path: ["\u0005SummaryInformation"], bytes: new Uint8Array(200) }
		])
		const document = await unlockFile(source, PASSWORD)
		const { unedited } = await editAndSave(document, "x")
		const cfb = await CfbFile.open(bytesSource(unedited.bytes))
		const names = [...cfb.walk()].map(entry => entry.path.join("/"))

		expect(cfb.version).toBe(3)
		expect(names.sort()).toEqual(
			[
				"",
				"EncryptedPackage",
				"EncryptionInfo",
				"\u0006DataSpaces",
				"\u0006DataSpaces/DataSpaceInfo",
				"\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace",
				"\u0006DataSpaces/DataSpaceMap",
				"\u0006DataSpaces/TransformInfo",
				"\u0006DataSpaces/TransformInfo/StrongEncryptionTransform",
				"\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary",
				"\u0006DataSpaces/Version"
			].sort()
		)
		expect(isEncryptedContainer(cfb)).toBe(true)

		const container = await readEncryptedContainer(cfb)

		expect(classifyDataSpaces(container.dataSpaces)).toBe("password")

		// An Agile source keeps its parameters; every save writes Agile.
		const info = parseEncryptionInfo(container.encryptionInfo ?? new Uint8Array(0))

		expect(info.kind === "agile" && [info.keyData.hash, info.keyData.keyBits, info.password.spinCount]).toEqual(["SHA-256", 128, 1000])
		document.release()
	})

	it("upgrades a Standard source to Agile AES-256 / SHA-512 / 100,000 on save", { timeout: 60_000 }, async () => {
		const document = await unlockFile(await encryptedFile(standardStreams(PLAINTEXT, PASSWORD)), PASSWORD)
		const { unedited } = await editAndSave(document, "x")
		const container = await readEncryptedContainer(await CfbFile.open(bytesSource(unedited.bytes)))
		const info = parseEncryptionInfo(container.encryptionInfo ?? new Uint8Array(0))

		expect(info.kind === "agile" && [info.keyData.hash, info.keyData.keyBits, info.password.spinCount]).toEqual([
			"SHA-512",
			256,
			100_000
		])
		expect(await plaintextOf(unedited.bytes, PASSWORD)).toEqual(PLAINTEXT)
		document.release()
	})

	it("asks for the key only at the first seal, once, and again after a failed ask", async () => {
		const document = await unlockFile(await agileFile(), PASSWORD)
		const taken = await document.take()
		let asks = 0
		let fail = true
		const { sealer, release } = docxSealer({
			key: options => {
				asks++

				if (fail) {
					fail = false

					return Promise.reject(new Error("the worker went"))
				}

				return taken.key(options)
			}
		})

		expect(asks).toBe(0)
		await expect(sealer.seal(PLAINTEXT)).rejects.toSatisfy(
			(error: unknown) => error instanceof DocxSealError && error.reason === "check"
		)
		expectSealedUpload(await sealer.seal(PLAINTEXT))
		expectSealedUpload(await sealer.seal(PLAINTEXT))
		expect(asks).toBe(2)
		release()
		document.release()
	})

	it("refuses to save once the key is released: nothing is handed out", async () => {
		const document = await unlockFile(await agileFile(), PASSWORD)
		const { sealer, release } = docxSealer(await document.take())

		release()
		expect(sealer.released).toBe(true)
		await expect(sealer.seal(PLAINTEXT)).rejects.toSatisfy(
			(error: unknown) => error instanceof DocxSealError && error.reason === "released"
		)

		// The holder's release ends the document: nothing more can be taken.
		document.release()
		await expect(document.take()).rejects.toThrow()
	})
})
