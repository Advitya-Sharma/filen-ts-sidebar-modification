import { serveTextWorker } from "@filen/office/dom/sheet/text"
import { workerScope } from "@/features/spreadsheet/workers/workerScope"

// CSV and TSV files (lib/sheetHosts.ts): neither the workbook reader nor the formula engine is loaded here.
serveTextWorker(workerScope())
