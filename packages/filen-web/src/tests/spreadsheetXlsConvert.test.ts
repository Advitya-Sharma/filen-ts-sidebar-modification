import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { openXlsx } from "hucre/xlsx"
import { xlsChart, xlsFile, xlsLabel, xlsNumber, xlsWorkbookStream, XLS_CHART } from "@filen/office/testing/encryption"
import { xlsToXlsx, xlsxCopyName } from "@/features/spreadsheet/lib/xlsConvert"

const XLS = new Uint8Array(readFileSync(join(__dirname, "fixtures/spreadsheet/budget.xls")))

describe("xlsToXlsx", () => {
	it("converts every sheet's values into an .xlsx that opens", async () => {
		const workbook = await openXlsx(await xlsToXlsx(XLS))

		expect(workbook.sheets.map(sheet => sheet.name)).toEqual(["Budget", "Notes"])
		expect(workbook.sheets[0]?.rows).toEqual([
			["Item", "Cost"],
			["Rent", 1200],
			["Food", 300],
			["Total", 1500]
		])
		expect(workbook.sheets[1]?.rows).toEqual([["second sheet"]])
	})

	// A chart's series caches are NUMBER/LABEL records inside its own BOF…EOF substream ([MS-XLS] §2.4.21).
	it("keeps an embedded chart's cached series out of the sheet's cells, and reads the cells after it", async () => {
		const stream = xlsWorkbookStream([
			{
				name: "Data",
				records: [
					xlsLabel(0, 0, "Title"),
					xlsChart([xlsNumber(0, 0, 4074), xlsNumber(1, 0, 7), xlsLabel(2, 1, "series")]),
					xlsNumber(1, 0, 12.5)
				]
			}
		])
		const workbook = await openXlsx(await xlsToXlsx(await xlsFile(stream)))

		expect(workbook.sheets[0]?.rows).toEqual([["Title"], [12.5]])
	})

	it("gives a chart sheet no cells", async () => {
		const stream = xlsWorkbookStream([
			{ name: "Data", records: [xlsNumber(0, 0, 1)] },
			{ name: "Chart1", dt: XLS_CHART, records: [xlsNumber(0, 0, 2), xlsNumber(1, 0, 3), xlsNumber(2, 0, 4)] }
		])
		const workbook = await openXlsx(await xlsToXlsx(await xlsFile(stream)))

		expect(workbook.sheets.map(sheet => sheet.name)).toEqual(["Data", "Chart1"])
		expect(workbook.sheets[0]?.rows).toEqual([[1]])
		expect(workbook.sheets[1]?.rows ?? []).toEqual([])
	})
})

describe("xlsxCopyName", () => {
	it("swaps the extension for .xlsx", async () => {
		expect(await xlsxCopyName("budget.xls", () => Promise.resolve(false))).toBe("budget.xlsx")
		expect(await xlsxCopyName("Q3 Report.XLS", () => Promise.resolve(false))).toBe("Q3 Report.xlsx")
	})

	it("numbers the copy when the name is taken, the way a file manager does", async () => {
		const taken = new Set(["budget.xlsx", "budget (2).xlsx"])

		expect(await xlsxCopyName("budget.xls", name => Promise.resolve(taken.has(name)))).toBe("budget (3).xlsx")
	})
})
