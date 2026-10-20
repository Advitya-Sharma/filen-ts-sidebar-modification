import type { SheetFileKind } from "@filen/office/core/sheet"
import { spreadsheetFileKind } from "@/features/spreadsheet/lib/fileKind"
import { sniffSpreadsheetKind } from "@/features/spreadsheet/lib/spreadsheetClient"

// What a file opens as: its extension when that names a spreadsheet (the templates and .xlsm open as themselves,
// so a save keeps their content type), else its first bytes; only a stored type can say tab-separated.
export function sheetKindOf(extension: string, bytes: Uint8Array, typeSaysTsv: boolean): { kind: SheetFileKind; unnamed: boolean } {
	if (spreadsheetFileKind(extension) !== null) {
		return { kind: extension as SheetFileKind, unnamed: false }
	}

	const sniffed = sniffSpreadsheetKind(bytes)

	return { kind: sniffed === "csv" && typeSaysTsv ? "tsv" : sniffed, unnamed: true }
}
