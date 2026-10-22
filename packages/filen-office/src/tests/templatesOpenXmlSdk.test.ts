import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { blankDocx, blankXlsx } from "@filen/office/core/templates"
import { OpenXmlSdk, openXmlSdkMissing, sdkErrorCounts } from "./oracles/openXmlSdk"

// Microsoft's Open XML SDK validator over every blank template, per paper: no error at all, allowlisted or not.
// Opt-in: FILEN_OFFICE_OOXML_VALIDATOR names the built validator (tools/openxml-validate), FILEN_OFFICE_DOTNET
// the dotnet host; skips without them.

const MISSING = openXmlSdkMissing()

const TEMPLATES: readonly (readonly [name: string, build: () => Promise<Uint8Array>])[] = [
	["letter.docx", () => blankDocx({ locale: "en-US" })],
	["a4.docx", () => blankDocx({ locale: "de-DE" })],
	["letter.xlsx", () => blankXlsx({ locale: "en-US", sheetName: "Sheet1" })],
	["a4.xlsx", () => blankXlsx({ locale: "fr-FR", sheetName: "Feuil1 & co" })]
]

describe.skipIf(MISSING !== undefined)("blank templates under the Open XML SDK validator", () => {
	let sdk: OpenXmlSdk
	let scratch: string

	beforeAll(() => {
		sdk = new OpenXmlSdk()
		scratch = mkdtempSync(join(tmpdir(), "filen-office-templates-"))
	})

	afterAll(async () => {
		await sdk.close()
		rmSync(scratch, { recursive: true, force: true })
	})

	it.each(TEMPLATES)("%s has no error", async (name, build) => {
		const path = join(scratch, name)

		writeFileSync(path, await build())

		expect([...sdkErrorCounts(await sdk.validate(path))]).toEqual([])
	})
})
