import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { encodeUtf8 } from "@filen/office/core/xml"
import { SCHEMA_ARCHIVE } from "../../fuzz/harness"
import { renderOrder, smlSchemas } from "../../../scripts/smlOrder"
import { transitionalSchemas } from "../../../scripts/wmlOrder"
import { CORPUS_DIR } from "../../corpusDir"
import { newValidityErrors, NOT_WELL_FORMED, schemaFamily, XsdOracle, xsdOracleMissing } from "../xsd"
import { SML } from "./synthetic"

// The SpreadsheetML schema oracle (ECMA-376 Transitional sml.xsd and the drawing schemas it imports, through
// xsd.ts) and the generated SpreadsheetML child-order table. Opt-in: needs xmllint and the Part 4 archive
// (pnpm run generate:wml-order --fetch); skips without them.

const MISSING = xsdOracleMissing()

describe.skipIf(MISSING !== undefined)("SpreadsheetML against the ECMA-376 schemas", () => {
	let oracle: XsdOracle

	beforeAll(() => {
		oracle = new XsdOracle()
	})

	afterAll(() => {
		oracle.dispose()
	})

	it("reports what it should", { timeout: 60_000 }, () => {
		const sheet = (body: string): string => `<worksheet xmlns="${SML}">${body}</worksheet>`
		const documents = [
			sheet('<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>'),
			// mergeCells before sheetData: out of schema order.
			sheet('<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><sheetData/>'),
			sheet('<sheetData><row r="1"><c r="A1" t="q"><v>0</v></c></row></sheetData>'),
			sheet("<sheetData>"),
			`<sst xmlns="${SML}" count="1" uniqueCount="1"><si><t>a</t></si></sst>`
		]
		const results = oracle.validate(documents.map(encodeUtf8)).map(errors => [...errors.keys()].join(" | "))

		expect(results[0]).toBe("")
		expect(results[1]).toMatch(/\}mergeCells': This element is not expected/)
		expect(results[2]).toMatch(/attribute 't'/)
		expect(results[3]).toBe(NOT_WELL_FORMED)
		expect(results[4]).toBe("")
	})

	it("names the SpreadsheetML parts", () => {
		expect(schemaFamily("xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")).toBe(
			"sml"
		)
		expect(schemaFamily("xl/workbook.xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml")).toBe("sml")
		expect(schemaFamily("xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml")).toBe("dml")
		expect(schemaFamily("xl/threadedComments/t.xml", "application/vnd.ms-excel.threadedcomments+xml")).toBeUndefined()
		expect(
			schemaFamily("xl/printerSettings/p.bin", "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings")
		).toBeUndefined()
	})

	// A couple of corpus workbooks: every SpreadsheetML part validates, and a part compared with itself adds nothing.
	it.skipIf(!existsSync(join(CORPUS_DIR, "xlsx")))("validates corpus workbooks", { timeout: 300_000 }, async () => {
		for (const relative of ["xlsx/govuk__25-26_CPID_list_.xlsx", "xlsx/lo__tdf105272.xlsx", "xlsx/poi__FormatConditionTests.xlsx"]) {
			const path = join(CORPUS_DIR, relative)

			if (!existsSync(path)) continue

			const pkg = await OpcPackage.open(bytesSource(new Uint8Array(readFileSync(path))))
			const types = await pkg.contentTypes()
			const parts = pkg.zip.entries.map(entry => entry.name).filter(name => schemaFamily(name, types.typeOf(name)) === "sml")
			const results = oracle.validate(await Promise.all(parts.map(part => pkg.read(part))))
			const counts = results.map((errors, index) => `${parts[index]}: ${[...errors.values()].reduce((sum, count) => sum + count, 0)}`)

			expect(parts.length).toBeGreaterThan(0)
			for (const errors of results) expect(newValidityErrors(errors, errors)).toEqual([])

			console.log(`sml-schema ${relative}: ${counts.join(", ")}`)
		}
	})
})

describe.skipIf(!existsSync(SCHEMA_ARCHIVE))("the SpreadsheetML order table", () => {
	it("lists the worksheet's children in schema order", { timeout: 60_000 }, async () => {
		const schema = smlSchemas(transitionalSchemas(new Uint8Array(readFileSync(SCHEMA_ARCHIVE)))).get("sml.xsd") as Uint8Array
		const source = await renderOrder(schema)
		const worksheet = /export const WORKSHEET: readonly string\[\] = \[([^\]]*)\]/.exec(source)?.[1] ?? ""
		const order = [...worksheet.matchAll(/"(\w+)"/g)].map(match => match[1])

		expect(order.indexOf("sheetData")).toBeLessThan(order.indexOf("mergeCells"))
		expect(order.indexOf("mergeCells")).toBeLessThan(order.indexOf("conditionalFormatting"))
		expect(order.indexOf("conditionalFormatting")).toBeLessThan(order.indexOf("dataValidations"))
		expect(order.at(-1)).toBe("extLst")
		expect(source).toMatch(/export const CELL: readonly string\[\] = \["f", "v", "is", "extLst"\]/)
		expect(source).toMatch(/export const CELL_TYPE: ReadonlySet<string> = new Set\(\["b", "n", "e", "s", "str", "inlineStr"\]\)/)
	})
})
