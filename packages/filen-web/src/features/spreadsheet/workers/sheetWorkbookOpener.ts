import { MAX_SHEET_CELLS } from "@filen/office/core/sheet"
import { workbookOpener, type SheetOpener } from "@filen/office/dom/sheet"
import { xlsToXlsx } from "@/features/spreadsheet/lib/xlsConvert"

// What the workbook worker opens with. A legacy .xls or .xlt is converted to an .xlsx first and opened from that,
// view-only (the session's "legacyFormat"), decrypted first when it has a password to open; one that is really an
// encrypted OOXML package opens as one.
export function sheetWorkbookOpener(): SheetOpener {
	return workbookOpener({ legacy: file => xlsToXlsx(file, MAX_SHEET_CELLS) })
}
