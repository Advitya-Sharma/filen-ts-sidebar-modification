import type { SheetFileKind, SheetSummary } from "@filen/office/core/sheet"
import { DEFAULT_PAGE_SETUP } from "@filen/office/core/sheet/print"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { isTextSheetKind } from "@/features/spreadsheet/lib/sheetHosts"
import type { PrintSheetSettings } from "@/features/spreadsheet/print/print.logic"

// A sheet's print settings: the file's for a workbook (a chartsheet's with its chart), the defaults for a text file
// (which holds none); null for a sheet that does not print (a dialog or macro sheet). `model`: the grid's (its
// layouts already loaded are reused), or the session alone where no grid shows it.
export async function sheetSettings(
	model: Pick<SheetGridModel, "layout" | "session">,
	kind: SheetFileKind,
	sheet: SheetSummary
): Promise<PrintSheetSettings | null> {
	if (isTextSheetKind(kind)) {
		const layout = model.layout(sheet.id) ?? (await model.session.layout(sheet.id))
		const used =
			layout.rowCount > 0 && layout.colCount > 0
				? { startRow: 0, startCol: 0, endRow: layout.rowCount - 1, endCol: layout.colCount - 1 }
				: null

		return { sheet, setup: DEFAULT_PAGE_SETUP, used }
	}

	if (sheet.kind !== "worksheet" && sheet.kind !== "chartsheet") return null

	const answer = await model.session.query("print.settings", { sheet: sheet.id })

	if (answer === null) return null

	return answer.chart === undefined
		? { sheet, setup: answer.setup, used: answer.used }
		: { sheet, setup: answer.setup, used: answer.used, chart: answer.chart }
}
