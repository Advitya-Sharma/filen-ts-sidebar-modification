import { bytesSource } from "@filen/office/core/opc"
import { MAX_SHEET_CELLS } from "@filen/office/core/sheet"
import { encryptedPackageVerdict } from "@filen/office/dom/encryption"
import { serveSheetWorker, workbookOpener } from "@filen/office/dom/sheet"
import { xlsToXlsx } from "@/features/spreadsheet/lib/xlsConvert"
import { workerScope } from "@/features/spreadsheet/next/workers/workerScope"

// Every workbook the page opens (lib/sheetHosts.ts): the package, its sessions, the formula engine and a
// password-protected file's key all stay here. A legacy .xls is converted to an .xlsx first and opened from that,
// view-only (the session's "legacyFormat"); an .xls that is really an encrypted OOXML package opens as one.
const open = workbookOpener()

serveSheetWorker(workerScope(), async (source, options, onEvent, track) => {
	if (options.kind !== "xls") {
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
})
