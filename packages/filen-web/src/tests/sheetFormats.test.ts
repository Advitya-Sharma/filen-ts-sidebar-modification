import { readFileSync } from "node:fs"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import type { SheetOpenResult } from "@filen/office/core/sheet"

// The spreadsheet formats beyond .xlsx, CSV and TSV that Google's "can open" list holds: a legacy template (.xlt)
// opens like an .xls (view-only, Save as .xlsx) and a .tab file as tab-separated text, in the office engine's viewer;
// builds without it keep routing as before. The modules are loaded fresh once under each flag.

const XLS = new Uint8Array(readFileSync(join(__dirname, "fixtures/spreadsheet/budget.xls")))
const IMMEDIATE = { sliceMs: 1_000_000, now: () => 0, pause: () => Promise.resolve() }

async function withFlag(flag: boolean) {
	vi.resetModules()
	vi.doMock("@/features/spreadsheet/next/sheetNextFlag", () => ({ SHEET_NEXT: flag }))

	return {
		fileKind: await import("@/features/spreadsheet/lib/fileKind"),
		preview: await import("@/features/drive/lib/preview.logic"),
		sheetKind: await import("@/features/spreadsheet/next/lib/sheetKind"),
		saveAs: await import("@/features/spreadsheet/next/save/sheetSaveAs"),
		messages: await import("@/features/spreadsheet/next/lib/sheetMessages.logic")
	}
}

type Loaded = Awaited<ReturnType<typeof withFlag>>

describe("spreadsheet formats in the office engine's viewer", () => {
	let on: Loaded
	let off: Loaded

	// Fresh module graphs are slow to load on a busy machine: once per flag, outside the tests' own time limit.
	beforeAll(async () => {
		on = await withFlag(true)
		off = await withFlag(false)
	}, 120_000)

	afterAll(() => {
		vi.doUnmock("@/features/spreadsheet/next/sheetNextFlag")
		vi.resetModules()
	})

	it("routes .xlt and .tab as spreadsheets: the template view-only, the text file editable", () => {
		const { fileKind, preview } = on

		for (const extension of ["xlt", "tab"]) expect(preview.previewCategoryForExtension(extension), extension).toBe("spreadsheet")

		expect(fileKind.isEditableSpreadsheetExtension("xlt")).toBe(false)
		expect(fileKind.isEditableSpreadsheetExtension("tab")).toBe(true)
		// A rename between .tab and .tsv keeps the format; one from .xlt to .xls too (both only save as .xlsx).
		expect(fileKind.spreadsheetSaveFormat("tab")).toBe(fileKind.spreadsheetSaveFormat("tsv"))
		expect(fileKind.spreadsheetSaveFormat("xlt")).toBe(fileKind.spreadsheetSaveFormat("xls"))
	})

	it("opens each as its own kind, the template as a legacy workbook offering Save as .xlsx", () => {
		const { sheetKind, saveAs, messages } = on

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

	it("keeps the routing of builds without it: neither extension is a spreadsheet there", () => {
		const { fileKind, preview } = off

		for (const extension of ["xlt", "tab"]) {
			expect(preview.previewCategoryForExtension(extension), extension).toBeNull()
			expect(fileKind.spreadsheetFileKind(extension), extension).toBeNull()
		}

		expect(fileKind.SPREADSHEET_EXTENSIONS).toEqual(["xlsx", "xlsm", "xltx", "xltm", "xls", "csv", "tsv"])
	})
})

describe("the workbook worker's opener", () => {
	let sheetWorkbookOpener: typeof import("@/features/spreadsheet/next/workers/sheetWorkbookOpener").sheetWorkbookOpener

	beforeAll(async () => {
		;({ sheetWorkbookOpener } = await import("@/features/spreadsheet/next/workers/sheetWorkbookOpener"))
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
