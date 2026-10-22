import type { SheetFileKind, SheetSummary } from "@filen/office/core/sheet"
import { DEFAULT_PAGE_SETUP } from "@filen/office/core/sheet/print"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { isTextSheetKind } from "@/features/spreadsheet/lib/sheetHosts"
import type { PrintSheetSettings } from "@/features/spreadsheet/print/print.logic"

// A sheet's print settings: the file's for a workbook, the defaults for a text file (which holds none); null for a
// sheet that does not print (a chart sheet).
export async function sheetSettings(model: SheetGridModel, kind: SheetFileKind, sheet: SheetSummary): Promise<PrintSheetSettings | null> {
	if (isTextSheetKind(kind)) {
		const layout = model.layout(sheet.id) ?? (await model.session.layout(sheet.id))
		const used =
			layout.rowCount > 0 && layout.colCount > 0
				? { startRow: 0, startCol: 0, endRow: layout.rowCount - 1, endCol: layout.colCount - 1 }
				: null

		return { sheet, setup: DEFAULT_PAGE_SETUP, used }
	}

	if (sheet.kind !== "worksheet") return null

	const answer = await model.session.query("print.settings", { sheet: sheet.id })

	return answer === null ? null : { sheet, setup: answer.setup, used: answer.used }
}
