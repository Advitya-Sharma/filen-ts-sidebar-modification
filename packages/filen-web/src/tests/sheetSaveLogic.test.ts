import { describe, expect, it } from "vitest"
import { csvDownloadName, fileStem, saveAsRefusalKey } from "@/features/spreadsheet/next/save/sheetSave.logic"

describe("what a spreadsheet's copies are called", () => {
	it("drops the extension, keeping a name that is all extension", () => {
		expect(fileStem("budget.csv")).toBe("budget")
		expect(fileStem("q1.report.xlsx")).toBe("q1.report")
		expect(fileStem(".csv")).toBe(".csv")
		expect(fileStem("README")).toBe("README")
	})

	it("downloads a one-sheet file's CSV under its own name, and a workbook's sheet with the sheet's name after it", () => {
		expect(csvDownloadName("data.tsv", "data", 1)).toBe("data.csv")
		expect(csvDownloadName("budget.xlsx", "Sheet1", 1)).toBe("budget.csv")
		expect(csvDownloadName("budget.xlsx", "Q1 / Q2: totals?", 3)).toBe("budget - Q1 _ Q2_ totals_.csv")
	})

	it("says why a Save as .xlsx was refused", () => {
		expect(saveAsRefusalKey({ reason: "tooLarge", detail: "grid" })).toBe("previewSheetSaveAsXlsxTooLarge")
		expect(saveAsRefusalKey({ reason: "unsupported" })).toBe("previewSpreadsheetSaveAsXlsxFailed")
		expect(saveAsRefusalKey(null)).toBe("previewSpreadsheetSaveAsXlsxFailed")
	})
})
