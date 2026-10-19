import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { CORPUS_DIR } from "../../corpusDir"
import { loadAllowlist, newSdkErrors, OpenXmlSdk, openXmlSdkMissing, sdkErrorCounts } from "../openXmlSdk"
import { buildXlsx } from "./synthetic"

// The Open XML SDK validator (tools/openxml-validate, FileFormatVersions.Microsoft365) on SpreadsheetML packages:
// it opens .xlsx/.xlsm/.xltx/.xltm as spreadsheets. Opt-in: FILEN_OFFICE_OOXML_VALIDATOR names the built
// validator, FILEN_OFFICE_DOTNET the dotnet host; skips without them.

const MISSING = openXmlSdkMissing()

describe.skipIf(MISSING !== undefined)("SpreadsheetML against the Open XML SDK validator", () => {
	let sdk: OpenXmlSdk
	let scratch: string

	beforeAll(() => {
		sdk = new OpenXmlSdk()
		scratch = mkdtempSync(join(tmpdir(), "filen-office-openxml-sheet-"))
	})

	afterAll(async () => {
		await sdk.close()
		rmSync(scratch, { recursive: true, force: true })
	})

	it("reports what it should", { timeout: 60_000 }, async () => {
		const valid = join(scratch, "valid.xlsx")
		const invalid = join(scratch, "invalid.xlsx")

		writeFileSync(valid, buildXlsx({ sheets: [{ name: "S", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }] }))
		// mergeCells before sheetData: out of schema order.
		writeFileSync(
			invalid,
			buildXlsx({ sheets: [{ name: "S", body: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><sheetData/>' }] })
		)

		const good = await sdk.validate(valid)
		const bad = sdkErrorCounts(await sdk.validate(invalid))

		expect(good.fatal).toBeUndefined()
		expect(good.errors).toEqual([])
		expect([...bad.keys()].join("\n")).toMatch(/\/xl\/worksheets\/sheet1\.xml Sch_/)
		expect(newSdkErrors(new Map(), bad, loadAllowlist()).length).toBeGreaterThan(0)
	})

	// A couple of corpus workbooks: each opens as a spreadsheet, and compared with itself adds nothing.
	it.skipIf(!existsSync(join(CORPUS_DIR, "xlsx")))("validates corpus workbooks", { timeout: 300_000 }, async () => {
		for (const relative of ["xlsx/govuk__25-26_CPID_list_.xlsx", "xlsx/lo__tdf105272.xlsx", "xlsm/docx4j__pivot.xlsm"]) {
			const path = join(CORPUS_DIR, relative)

			if (!existsSync(path)) continue

			const result = await sdk.validate(path)
			const counts = sdkErrorCounts(result)

			expect(result.fatal, relative).toBeUndefined()
			expect(newSdkErrors(counts, counts, loadAllowlist())).toEqual([])
			console.log(`sheet-openxml ${relative}: ${result.errors.length} errors in the source`)
		}
	})
})
