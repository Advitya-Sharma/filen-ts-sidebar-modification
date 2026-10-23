import { describe, expect, it } from "vitest"
import { bytesSource, ZipArchive } from "@filen/office/core/opc"
import type { DocxLegacyOpen } from "../dom/docx/legacy"
import { encryptedDoc } from "./docDecryptFixtures"
import { buildDoc } from "./docFixtures"
import { workerHarness } from "./docxEncryptedFixtures"

// An encrypted Word 97–2003 document through the docx worker, as the web opens it: refused as encrypted, then opened
// with the password (WebCrypto SHA-1 in the worker for CryptoAPI RC4); a wrong one is refused and can be retried.

function refusal(opened: DocxLegacyOpen): string {
	return opened.kind === "refused" ? `${opened.refusal.reason}:${opened.refusal.encryption ?? ""}` : opened.kind
}

describe("encrypted Word binary documents in the docx worker", () => {
	it.each(["xor", "rc4", "cryptoApiRc4"] as const)(
		"asks for the %s password, refuses a wrong one, opens with the right one",
		{ timeout: 30_000 },
		async scheme => {
			const harness = workerHarness()

			try {
				const { bytes } = await encryptedDoc({ pieces: [{ text: "Behind a password\r" }] }, { scheme, password: "Geheim" })
				const source = bytesSource(bytes)
				const first = await harness.pool.legacy.open(source)
				const wrong = await harness.pool.legacy.open(source, { password: "geheim" })
				const right = await harness.pool.legacy.open(source, { password: "Geheim" })

				expect([refusal(first), refusal(wrong), right.kind]).toEqual([`encrypted:${scheme}`, "wrongPassword:", "converted"])

				if (right.kind !== "converted") return

				const copy = await right.document.take()
				const zip = await ZipArchive.open(bytesSource(copy))
				const entry = zip.get("word/document.xml")
				const main = entry === undefined ? "" : new TextDecoder().decode(await zip.read(entry))

				expect([right.document.decrypted, main.includes("Behind a password")]).toEqual([true, true])
				right.document.release()
			} finally {
				harness.dispose()
			}
		}
	)

	it("says a plain document was not decrypted", { timeout: 30_000 }, async () => {
		const harness = workerHarness()

		try {
			const opened = await harness.pool.legacy.open(bytesSource(await buildDoc({ pieces: [{ text: "Open\r" }] })))

			expect(opened.kind === "converted" && !opened.document.decrypted).toBe(true)

			if (opened.kind === "converted") opened.document.release()
		} finally {
			harness.dispose()
		}
	})
})
