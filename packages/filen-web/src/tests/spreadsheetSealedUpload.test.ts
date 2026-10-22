import { describe, expect, it } from "vitest"
import { assertSealedUpload } from "@/features/spreadsheet/lib/xlsxEncryption"
import { encryptedWorkbook, plainWorkbook } from "@/tests/encryptedWorkbookFixtures"

describe("a protected workbook's upload guard", () => {
	it("passes a sealed workbook", async () => {
		const sealed = await encryptedWorkbook("agile")

		expect(() => {
			assertSealedUpload(sealed)
		}).not.toThrow()
	})

	it("stops the plain zip, and bytes that are no workbook at all", () => {
		for (const bytes of [plainWorkbook(), new Uint8Array(0)]) {
			expect(() => {
				assertSealedUpload(bytes)
			}).toThrow(/not sealed/)
		}
	})
})
