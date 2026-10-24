import type { SheetFileKind, SheetRefusal } from "@filen/office/core/sheet"
import type { PreviewKey } from "@/lib/i18n"
import type { SheetHandle } from "@/features/spreadsheet/hooks/useSheetSession"
import { isTextSheetKind, textToXlsx } from "@/features/spreadsheet/lib/sheetHosts"
import { isLegacySheetKind } from "@/features/spreadsheet/lib/sheetKind"
import { fileStem } from "@/features/spreadsheet/save/sheetSave.logic"

export type SheetXlsxCopy =
	{ readonly ok: true; readonly bytes: Uint8Array } | { readonly ok: false; readonly refusal: SheetRefusal | null }

// Which open files Save as .xlsx: a text file (its cells as they now are, edits included) and a legacy workbook
// (.xls or .xlt, opened from its conversion, view-only).
export function savesAsXlsx(kind: SheetFileKind): boolean {
	return isLegacySheetKind(kind) || isTextSheetKind(kind)
}

// What the confirm before a Save as .xlsx says the copy keeps: a text file's cells as typed, or a legacy workbook's
// values, naming the format it is in.
export function saveAsXlsxBodyKey(kind: SheetFileKind): PreviewKey {
	if (isTextSheetKind(kind)) {
		return "previewSheetSaveAsXlsxTextBody"
	}

	return kind === "xlt" ? "previewSheetSaveAsXlsxXltBody" : "previewSpreadsheetSaveAsXlsxBody"
}

// The .xlsx a Save as uploads. A text file's cells go from its worker to the workbook worker, which makes the
// workbook, its sheet named after the file. A legacy workbook's session holds the .xlsx it was converted to, saved
// as it stands. An open entry is the caller's to commit first.
export async function sheetXlsxCopy(handle: SheetHandle, kind: SheetFileKind, fileName: string): Promise<SheetXlsxCopy> {
	if (isTextSheetKind(kind)) {
		const cells = await handle.session.query("saveAs.cells", {})

		// A file only partly read (more rows than a sheet holds) is refused: the copy would leave the rest out.
		if (!cells.ok) {
			return cells
		}

		const converted = await textToXlsx(cells.bytes, fileStem(fileName))

		return converted.ok ? { ok: true, bytes: converted.bytes } : { ok: false, refusal: converted.refusal }
	}

	const saved = await handle.session.save()

	return saved.ok ? { ok: true, bytes: saved.bytes } : { ok: false, refusal: null }
}
