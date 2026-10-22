import { readFileSync } from "node:fs"
import { join } from "node:path"
import { beforeAll, describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import type { SheetOpenResult } from "@filen/office/core/sheet"
import * as fileKind from "@/features/spreadsheet/lib/fileKind"
import * as preview from "@/features/drive/lib/preview.logic"
import * as sheetKind from "@/features/spreadsheet/lib/sheetKind"
import * as saveAs from "@/features/spreadsheet/save/sheetSaveAs"
import * as messages from "@/features/spreadsheet/lib/sheetMessages.logic"

// The spreadsheet formats beyond .xlsx, CSV and TSV that Google's "can open" list holds: a legacy template (.xlt)
// opens like an .xls (view-only, Save as .xlsx) and a .tab file as tab-separated text.

const XLS = new Uint8Array(readFileSync(join(__dirname, "fixtures/spreadsheet/budget.xls")))
const IMMEDIATE = { sliceMs: 1_000_000, now: () => 0, pause: () => Promise.resolve() }

describe("spreadsheet formats", () => {
	it("routes .xlt and .tab as spreadsheets: the template view-only, the text file editable", () => {
		for (const extension of ["xlt", "tab"]) expect(preview.previewCategoryForExtension(extension), extension).toBe("spreadsheet")

		expect(fileKind.isEditableSpreadsheetExtension("xlt")).toBe(false)
		expect(fileKind.isEditableSpreadsheetExtension("tab")).toBe(true)
		// A rename between .tab and .tsv keeps the format; one from .xlt to .xls too (both only save as .xlsx).
		expect(fileKind.spreadsheetSaveFormat("tab")).toBe(fileKind.spreadsheetSaveFormat("tsv"))
		expect(fileKind.spreadsheetSaveFormat("xlt")).toBe(fileKind.spreadsheetSaveFormat("xls"))
	})

	it("opens .xltx and .xltm as workbooks, telling them apart from .xlsx and .xlsm", () => {
		expect(fileKind.spreadsheetFileKind("xltx")).toBe("xlsx")
		expect(fileKind.spreadsheetFileKind("xltm")).toBe("xlsx")
		expect(new Set(["xlsx", "xlsm", "xltx", "xltm"].map(fileKind.spreadsheetSaveFormat)).size).toBe(4)
	})

	it("opens each as its own kind, the template as a legacy workbook offering Save as .xlsx", () => {
		expect(sheetKind.sheetKindOf("xlt", XLS, false)).toEqual({ kind: "xlt", unnamed: false })
		expect(sheetKind.sheetKindOf("tab", new TextEncoder().encode("a\tb"), false)).toEqual({ kind: "tab", unnamed: false })
		expect(sheetKind.isLegacySheetKind("xlt")).toBe(true)
		expect(sheetKind.isLegacySheetKind("xlsx")).toBe(false)
		expect(saveAs.savesAsXlsx("xlt")).toBe(true)
		expect(saveAs.savesAsXlsx("tab")).toBe(true)
		expect(saveAs.saveAsXlsxBodyKey("xlt")).toBe("previewSheetSaveAsXlsxXltBody")
		expect(saveAs.saveAsXlsxBodyKey("xls")).toBe("previewSpreadsheetSaveAsXlsxBody")
		expect(saveAs.saveAsXlsxBodyKey("tab")).toBe("previewSheetSaveAsXlsxTextBody")

		const legacy = { state: "viewOnly", reason: "legacyFormat" } as const
		const note = { editable: true, readOnlyReason: undefined, unnamed: false, renamed: false, writability: legacy }

		expect(messages.readOnlyNoteKey({ ...note, kind: "xlt" })).toBe("previewSheetReadOnlyXlt")
		expect(messages.readOnlyNoteKey({ ...note, kind: "xls" })).toBe("previewSpreadsheetReadOnlyXls")
		expect(messages.viewOnlyKey("legacyFormat", "xlt")).toBe("previewSheetReadOnlyXlt")
	})
})

describe("the workbook worker's opener", () => {
	let sheetWorkbookOpener: typeof import("@/features/spreadsheet/workers/sheetWorkbookOpener").sheetWorkbookOpener

	beforeAll(async () => {
		;({ sheetWorkbookOpener } = await import("@/features/spreadsheet/workers/sheetWorkbookOpener"))
	}, 120_000)

	it("opens a legacy template from its conversion, view-only, with the workbook's sheets", async () => {
		const sessions: unknown[] = []
		const result: SheetOpenResult = await sheetWorkbookOpener()(
			bytesSource(XLS.slice()),
			{ kind: "xlt", slicing: IMMEDIATE },
			() => undefined,
			session => sessions.push(session)
		)

		if (result.kind !== "opened") throw new Error(`not opened: ${result.kind}`)

		expect(result.info.kind).toBe("xlt")
		expect(result.info.writability).toEqual({ state: "viewOnly", reason: "legacyFormat" })
		expect(result.info.sheets.map(sheet => sheet.name)).toEqual(["Budget", "Notes"])
		expect(sessions).toHaveLength(1)
		result.session.close()
	})

	it("says a legacy file it cannot read is malformed", async () => {
		const result = await sheetWorkbookOpener()(
			bytesSource(new TextEncoder().encode("not a workbook")),
			{ kind: "xlt", slicing: IMMEDIATE },
			() => undefined,
			() => undefined
		)

		expect(result.kind).toBe("failed")
	})
})
