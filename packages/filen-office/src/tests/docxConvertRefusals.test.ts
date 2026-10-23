import { mkdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { encryptedEntry } from "./odsFixtures"
import { odtBytes } from "./odtFixtures"
import { strFromU8, unzipSync } from "fflate"
import { convertText } from "../core/convert/text/index"
import { encryptedDoc } from "./docDecryptFixtures"
import { buildDoc } from "./docFixtures"
import { workerHarness } from "./docxEncryptedFixtures"

// What a conversion tells about a file it cannot convert as named: a Word 97–2003 document named .docx is sniffed and
// converted as the viewer opens it; an encrypted Word 97–2003 document's refusal says which cipher it uses, so a
// password is asked only where one can open it (a CryptoAPI cipher other than RC4 cannot).

// FILEN_OFFICE_CONVERT_REFUSALS_OUT names a directory the refused files are written to (a browser harness's). Opt-in.
const OUT = process.env["FILEN_OFFICE_CONVERT_REFUSALS_OUT"] || undefined

// A CryptoAPI RC4 document whose EncryptionHeader names AES (AlgID 0x660E): a cipher this reader does not decrypt.
function aesHeader(_word: Uint8Array, table: Uint8Array): void {
	new DataView(table.buffer, table.byteOffset, table.byteLength).setUint32(12 + 8, 0x660e, true)
}

describe("conversion refusals and sniffing", () => {
	it("converts a Word 97–2003 document named .docx to .odt, as the viewer opens it", async () => {
		const doc = await buildDoc({ pieces: [{ text: "Binary under a docx name\r" }] })
		const result = await convertText(doc, "docx", "odt")

		if (!result.ok) throw new Error(result.refusal.message)

		expect(strFromU8(unzipSync(result.bytes)["content.xml"] ?? new Uint8Array())).toContain("Binary under a docx name")
	})

	it("names the cipher of an encrypted Word 97–2003 document it refuses", async () => {
		const rc4 = await encryptedDoc({ pieces: [{ text: "x\r" }] }, { scheme: "rc4", password: "pw" })
		const aes = await encryptedDoc({ pieces: [{ text: "x\r" }] }, { scheme: "cryptoApiRc4", password: "pw" }, aesHeader)

		expect(await convertText(rc4.bytes, "doc", "docx")).toMatchObject({
			ok: false,
			refusal: { reason: "encrypted", encryption: "rc4" }
		})
		// Told apart before any password is asked for.
		expect(await convertText(aes.bytes, "doc", "docx")).toMatchObject({
			ok: false,
			refusal: { reason: "encrypted", encryption: "unknown" }
		})
	})

	it("reads an AlgID of 0 as the cipher the header's flags name (RC4 without fAES, AES with it)", async () => {
		// EncryptionHeader at table offset 12: Flags, SizeExtra, AlgID, AlgIDHash ([MS-OFFCRYPTO] §2.3.2).
		const header = (flags: number, algId: number, algIdHash: number) => (_word: Uint8Array, table: Uint8Array) => {
			const view = new DataView(table.buffer, table.byteOffset, table.byteLength)

			view.setUint32(12, flags, true)
			view.setUint32(12 + 8, algId, true)
			view.setUint32(12 + 12, algIdHash, true)
		}
		const rc4 = await encryptedDoc(
			{ pieces: [{ text: "Default cipher\r" }] },
			{ scheme: "cryptoApiRc4", password: "pw" },
			header(0x04, 0, 0)
		)
		const aes = await encryptedDoc({ pieces: [{ text: "x\r" }] }, { scheme: "cryptoApiRc4", password: "pw" }, header(0x24, 0, 0x8004))
		const aes128 = await encryptedDoc(
			{ pieces: [{ text: "x\r" }] },
			{ scheme: "cryptoApiRc4", password: "pw" },
			header(0x24, 0x660e, 0x8004)
		)

		expect(await convertText(rc4.bytes, "doc", "docx")).toMatchObject({
			ok: false,
			refusal: { reason: "encrypted", encryption: "cryptoApiRc4" }
		})

		const opened = await convertText(rc4.bytes, "doc", "docx", {
			decryption: { password: "pw", sha1: async data => new Uint8Array(await crypto.subtle.digest("SHA-1", data)) }
		})

		if (!opened.ok) throw new Error(opened.refusal.message)

		expect(strFromU8(unzipSync(opened.bytes)["word/document.xml"] ?? new Uint8Array())).toContain("Default cipher")

		for (const file of [aes, aes128])
			expect(await convertText(file.bytes, "doc", "docx")).toMatchObject({
				ok: false,
				refusal: { reason: "encrypted", encryption: "unknown" }
			})
	})

	it("refuses a CryptoAPI document whose EncryptionHeader is damaged as corrupt, before any password", async () => {
		// EncryptionHeaderSize below the fixed 32 bytes of the header ([MS-OFFCRYPTO] §2.3.5.1).
		const damaged = await encryptedDoc({ pieces: [{ text: "x\r" }] }, { scheme: "cryptoApiRc4", password: "pw" }, (_word, table) => {
			new DataView(table.buffer, table.byteOffset, table.byteLength).setUint32(8, 8, true)
		})

		expect(await convertText(damaged.bytes, "doc", "docx")).toMatchObject({ ok: false, refusal: { reason: "corrupt" } })
	})

	it("carries the cipher through the docx worker's answer", async () => {
		const harness = workerHarness()

		try {
			const aes = await encryptedDoc({ pieces: [{ text: "x\r" }] }, { scheme: "cryptoApiRc4", password: "pw" }, aesHeader)

			expect(await harness.pool.convert.convert(aes.bytes, "doc", "docx", { password: "pw" })).toMatchObject({
				kind: "refused",
				reason: "encrypted",
				encryption: "unknown"
			})
		} finally {
			harness.dispose()
		}
	})

	it("refuses a password-protected OpenDocument text document as encrypted", async () => {
		const bytes = odtBytes({ body: "<text:p>Secret</text:p>", manifestExtra: encryptedEntry("Secret.xml") })
		const aes = await encryptedDoc({ pieces: [{ text: "x\r" }] }, { scheme: "cryptoApiRc4", password: "pw" }, aesHeader)

		expect(await convertText(bytes, "odt", "docx")).toMatchObject({ ok: false, refusal: { reason: "encrypted" } })

		if (OUT !== undefined) {
			mkdirSync(OUT, { recursive: true })
			writeFileSync(join(OUT, "secret.odt"), bytes)
			writeFileSync(join(OUT, "aes.doc"), aes.bytes)
		}
	})
})
