import { SpreadsheetViewer } from "@/features/spreadsheet/components/spreadsheetViewer"
import { SHEET_NEXT } from "@/features/spreadsheet/next/sheetNextFlag"
import { SheetNextViewer } from "@/features/spreadsheet/next/sheetNextViewer"

// The spreadsheet viewer the build carries: the office engine's with VITE_SHEET_NEXT=1, the current one otherwise.
// A constant, so the build drops the other.
export const SpreadsheetView: typeof SpreadsheetViewer = SHEET_NEXT ? SheetNextViewer : SpreadsheetViewer
