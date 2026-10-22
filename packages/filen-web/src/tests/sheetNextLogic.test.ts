import { describe, expect, it } from "vitest"
import { parseNameBox } from "@/features/spreadsheet/next/lib/nameBox.logic"
import { readOnlyNoteKey, refusalKey, viewOnlyKey } from "@/features/spreadsheet/next/lib/sheetMessages.logic"
import { SheetEventRelay } from "@/features/spreadsheet/next/lib/sheetEvents"
import type { SheetEvent } from "@filen/office/core/sheet"
import { preview } from "@/locales/en/preview"

describe("name box", () => {
	it("reads a cell, absolute or not, in either case", () => {
		expect(parseNameBox("B4")).toEqual({ startRow: 3, startCol: 1, endRow: 3, endCol: 1 })
		expect(parseNameBox(" $b$4 ")).toEqual({ startRow: 3, startCol: 1, endRow: 3, endCol: 1 })
		expect(parseNameBox("XFD1048576")).toEqual({ startRow: 1_048_575, startCol: 16_383, endRow: 1_048_575, endCol: 16_383 })
	})

	it("reads a range from either corner", () => {
		expect(parseNameBox("C3:A1")).toEqual({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 })
	})

	it("refuses anything past the grid or not a reference", () => {
		for (const input of ["", "A0", "XFE1", "A1048577", "A1:B2:C3", "Sheet1!A1", "1A", "=A1"]) {
			expect(parseNameBox(input), input).toBeNull()
		}
	})
})

describe("view-only and refusal messages", () => {
	const base = { editable: true, readOnlyReason: undefined, unnamed: false, renamed: false, kind: "xlsx" } as const

	it("names every view-only reason with a string the catalog has", () => {
		for (const reason of ["legacyFormat", "lossy", "tooLarge", "unsafe", "signed", "readOnlyRecommended", "modifyPassword"] as const) {
			expect(preview).toHaveProperty(viewOnlyKey(reason, "xlsx"))
		}

		expect(viewOnlyKey("unsafe", "csv")).toBe("previewSpreadsheetReadOnlyUnsafe")
		expect(viewOnlyKey("unsafe", "xlsx")).toBe("previewSpreadsheetReadOnlyLossy")
	})

	it("puts the app's own reasons before the session's verdict, and says a legacy workbook whoever looks", () => {
		const lossy = { state: "viewOnly", reason: "lossy" } as const

		expect(readOnlyNoteKey({ ...base, unnamed: true, writability: lossy })).toBe("previewSpreadsheetReadOnlyUnnamed")
		expect(readOnlyNoteKey({ ...base, renamed: true, writability: lossy })).toBe("previewSpreadsheetReadOnlyRenamed")
		expect(readOnlyNoteKey({ ...base, writability: lossy })).toBe("previewSpreadsheetReadOnlyLossy")
		expect(readOnlyNoteKey({ ...base, writability: { state: "checking" } })).toBe("previewSpreadsheetCheckingWritable")
		expect(readOnlyNoteKey({ ...base, writability: { state: "unproven" } })).toBeNull()
		expect(readOnlyNoteKey({ ...base, editable: false, writability: lossy })).toBeNull()
		expect(readOnlyNoteKey({ ...base, editable: false, readOnlyReason: "renamed", writability: lossy })).toBe(
			"previewSpreadsheetReadOnlyRenamed"
		)
		expect(readOnlyNoteKey({ ...base, editable: false, writability: { state: "viewOnly", reason: "legacyFormat" } })).toBe(
			"previewSpreadsheetReadOnlyXls"
		)
	})

	it("gives every refusal a string the catalog has, a view-only one by its verdict", () => {
		const reasons = [
			"readOnly",
			"tooLarge",
			"sheetName",
			"lastSheet",
			"formula",
			"arrayFormula",
			"tableHeader",
			"structureLocked",
			"pivot",
			"merged",
			"protectedSheet",
			"protectedWorkbook",
			"encoding",
			"styleLimit",
			"unsupported"
		] as const

		for (const reason of reasons) {
			expect(preview).toHaveProperty(refusalKey({ reason }, "xlsx"))
		}

		expect(refusalKey({ reason: "readOnly", detail: "signed" }, "xlsx")).toBe("previewSheetReadOnlySigned")
		expect(refusalKey({ reason: "readOnly", detail: "something else" }, "xlsx")).toBe("previewSpreadsheetReadOnlyLossy")
		// A sort over merged cells of different sizes says so, as Excel does.
		expect(refusalKey({ reason: "merged", detail: "mergeSize" }, "xlsx")).toBe("previewSheetSortMergeSize")
		expect(preview).toHaveProperty("previewSheetSortMergeSize")
	})
})

describe("event relay", () => {
	it("holds events until a target attaches, then passes them on in order", () => {
		const relay = new SheetEventRelay()
		const seen: SheetEvent["type"][] = []
		const engine = (state: "building" | "ready"): SheetEvent => ({ type: "engine", state })

		relay.emit(engine("building"))
		relay.emit({ type: "progress", task: "open", done: 1, total: 2 })

		const detach = relay.attach(event => seen.push(event.type))

		relay.emit(engine("ready"))
		expect(seen).toEqual(["engine", "progress", "engine"])

		detach()
		relay.emit(engine("ready"))
		expect(seen).toHaveLength(3)

		relay.attach(event => seen.push(event.type))
		expect(seen).toHaveLength(4)
	})
})
