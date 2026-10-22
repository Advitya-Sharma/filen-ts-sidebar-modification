import { serveSheetWorker } from "@filen/office/dom/sheet"
import { sheetWorkbookOpener } from "@/features/spreadsheet/next/workers/sheetWorkbookOpener"
import { workerScope } from "@/features/spreadsheet/next/workers/workerScope"

// Every workbook the page opens (lib/sheetHosts.ts): the package, its sessions, the formula engine and a
// password-protected file's key all stay here.
serveSheetWorker(workerScope(), sheetWorkbookOpener())
