import { mkdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { strFromU8, unzipSync } from "fflate"
import { isSealedOfficeFile } from "../core/encryption/index"
import { encryptedDoc } from "./docDecryptFixtures"
import { plaintextOf, workerHarness } from "./docxEncryptedFixtures"

// A password-protected Word 97–2003 document converted to .docx in the docx worker, as Download as converts it: refused
// as encrypted without its password, refused as wrong with another, and with the right one converted and sealed as an
// Agile-encrypted package under that same password, which opens with it. A protected file never becomes an .odt (no
// ODF encryption is written). FILEN_OFFICE_CONVERT_PROTECTED_OUT names a directory each sealed .docx is written to,
// with its source, for msoffcrypto-tool and LibreOffice to open with the password. Opt-in.

const OUT = process.env["FILEN_OFFICE_CONVERT_PROTECTED_OUT"] || undefined

describe("protected Word 97–2003 documents converted in the docx worker", () => {
	it.each(["xor", "rc4", "cryptoApiRc4"] as const)(
		"seals the %s document's .docx under the same password",
		{ timeout: 60_000 },
		async scheme => {
			const harness = workerHarness()

			try {
				const { bytes } = await encryptedDoc({ pieces: [{ text: "Behind a password\r" }] }, { scheme, password: "Geheim" })
				const convert = harness.pool.convert

				const without = await convert.convert(bytes, "doc", "docx")
				const wrong = await convert.convert(bytes, "doc", "docx", { password: "geheim" })
				const right = await convert.convert(bytes, "doc", "docx", { password: "Geheim" })

				expect(without.kind === "refused" ? without.reason : without.kind).toBe("encrypted")
				expect(wrong.kind === "refused" ? wrong.reason : wrong.kind).toBe("wrongPassword")

				if (right.kind !== "converted") throw new Error(right.kind)

				expect(isSealedOfficeFile(right.bytes)).toBe(true)

				if (OUT !== undefined) {
					mkdirSync(OUT, { recursive: true })
					writeFileSync(join(OUT, `${scheme}.doc`), bytes)
					writeFileSync(join(OUT, `${scheme}.docx`), right.bytes)
				}

				const plain = await plaintextOf(right.bytes, "Geheim")

				expect(strFromU8(unzipSync(plain)["word/document.xml"] ?? new Uint8Array())).toContain("Behind a password")
				await expect(plaintextOf(right.bytes, "geheim")).rejects.toThrow()
			} finally {
				harness.dispose()
			}
		}
	)

	it("refuses an .odt of a protected document even with its password", async () => {
		const harness = workerHarness()

		try {
			const { bytes } = await encryptedDoc({ pieces: [{ text: "Behind a password\r" }] }, { scheme: "rc4", password: "Geheim" })
			const result = await harness.pool.convert.convert(bytes, "doc", "odt", { password: "Geheim" })

			expect(result.kind === "refused" ? result.reason : result.kind).toBe("encrypted")
		} finally {
			harness.dispose()
		}
	})
})
