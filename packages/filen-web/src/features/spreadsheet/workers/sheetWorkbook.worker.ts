import { serveSheetWorker } from "@filen/office/dom/sheet"
import { officePdfFontAssets } from "@/features/office/lib/pdfFontAssets"
import { sheetWorkbookJobs } from "@/features/spreadsheet/workers/sheetWorkbookJobs"
import { sheetWorkbookOpener } from "@/features/spreadsheet/workers/sheetWorkbookOpener"
import { workerScope } from "@/features/spreadsheet/workers/workerScope"

// Every workbook the page opens (lib/sheetHosts.ts): the package, its sessions, the formula engine and a
// password-protected file's key all stay here, and a print's PDF is made here beside them.
serveSheetWorker(workerScope(), sheetWorkbookOpener(), sheetWorkbookJobs(officePdfFontAssets))
