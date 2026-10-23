import { serveSheetWorker, textOpener } from "@filen/office/dom/sheet/text"
import { sheetPdfJob } from "@filen/office/react/sheet/print"
import { officePdfFontAssets } from "@/features/office/lib/pdfFontAssets"
import { workerScope } from "@/features/spreadsheet/workers/workerScope"

// CSV and TSV files (lib/sheetHosts.ts): neither the workbook reader nor the formula engine is loaded here. A print's
// PDF is made here beside the session.
serveSheetWorker(workerScope(), textOpener(), sheetPdfJob(officePdfFontAssets))
