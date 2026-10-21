import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { CORPUS_DIR } from "./corpusDir"
import { openFixture, savedBytes } from "./sheetSessionFixtures"
import { parts } from "./xlsxWriteFixtures"

// A deleted sheet's parts: those only it reaches go, those something else reaches stay, and so does what they reach.

const PIVOTS = "xlsx/poi__LIBRE_OFFICE-128382-0.xlsx"

function corpus(name: string): Uint8Array | null {
	const path = join(CORPUS_DIR, name)

	return existsSync(path) ? new Uint8Array(readFileSync(path)) : null
}

describe("a sheet deleted", () => {
	it.skipIf(corpus(PIVOTS) === null)("keeps the records of a pivot cache the workbook still lists", async () => {
		const { session, info } = await openFixture(corpus(PIVOTS) as Uint8Array)
		const pivots = info.sheets.find(sheet => sheet.name === "Advanced features")

		expect((await session.command({ type: "sheet.delete", sheet: pivots?.id ?? -1 })).ok).toBe(true)

		const files = parts(await savedBytes(session))

		expect(Object.keys(files)).toContain("xl/pivotCache/pivotCacheDefinition1.xml")
		expect(Object.keys(files)).toContain("xl/pivotCache/pivotCache/pivotCacheRecords1.xml")
		expect(Object.keys(files)).not.toContain("xl/pivotTables/pivotTable1.xml")
		session.close()
	})
})
