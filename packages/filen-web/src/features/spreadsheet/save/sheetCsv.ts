import type { SheetId, SheetRefusal } from "@filen/office/core/sheet"
import type { SheetHandle } from "@/features/spreadsheet/hooks/useSheetSession"
import { csvDownloadName } from "@/features/spreadsheet/save/sheetSave.logic"
import { downloadBlob } from "@/lib/downloadBlob"

// Downloads one sheet as Excel's "CSV UTF-8 (Comma delimited)": each cell as it is shown, with a byte order mark,
// made by the session. A password-protected workbook's cells never leave it so: the session refuses. `sheets` is
// how many the file now holds (one CSV holds one).
export async function downloadSheetCsv(
	handle: SheetHandle,
	sheet: { readonly id: SheetId; readonly name: string },
	sheets: number,
	fileName: string
): Promise<{ readonly ok: true } | { readonly ok: false; readonly refusal: SheetRefusal }> {
	const result = await handle.session.export("csv", { sheet: sheet.id })

	if (!result.ok) {
		return result
	}

	downloadBlob(
		csvDownloadName(fileName, sheet.name, sheets),
		new Blob([result.bytes as Uint8Array<ArrayBuffer>], { type: "text/csv;charset=utf-8" })
	)

	return { ok: true }
}
