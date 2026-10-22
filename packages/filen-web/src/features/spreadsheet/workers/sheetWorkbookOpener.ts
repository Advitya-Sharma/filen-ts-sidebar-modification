import { bytesSource } from "@filen/office/core/opc"
import { MAX_SHEET_CELLS } from "@filen/office/core/sheet"
import { encryptedPackageVerdict } from "@filen/office/dom/encryption"
import { workbookOpener, type SheetOpener } from "@filen/office/dom/sheet"
import { xlsToXlsx } from "@/features/spreadsheet/lib/xlsConvert"
import { isLegacySheetKind } from "@/features/spreadsheet/lib/sheetKind"

// What the workbook worker opens with. A legacy .xls or .xlt is converted to an .xlsx first and opened from that,
// view-only (the session's "legacyFormat"); one that is really an encrypted OOXML package opens as one.
export function sheetWorkbookOpener(): SheetOpener {
	const open = workbookOpener()

	return async (source, options, onEvent, track) => {
		if (!isLegacySheetKind(options.kind)) {
			return open(source, options, onEvent, track)
		}

		if ((await encryptedPackageVerdict(source)) !== undefined) {
			return open(source, { ...options, kind: "xlsx" }, onEvent, track)
		}

		let converted: Uint8Array

		try {
			converted = await xlsToXlsx(await source.read(0, source.size), MAX_SHEET_CELLS)
		} catch (error) {
			return { kind: "failed", reason: "malformed", message: error instanceof Error ? error.message : String(error) }
		}

		return open(bytesSource(converted), options, onEvent, track)
	}
}
