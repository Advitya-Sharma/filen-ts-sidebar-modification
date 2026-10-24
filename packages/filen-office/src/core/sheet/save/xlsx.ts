import { textWorkbookXlsx, type TextWorkbookResult } from "../../xlsx/convert/text"
import { defineSheetFeature } from "../registry"
import type { TextSheetDocument } from "../text"

// Save as .xlsx of a text file, its second half, in the worker that holds the workbook writer: the cells
// saveAs.cells handed out, opened there as a TSV, become a workbook of one sheet (xlsx/convert/text.ts).

declare module "../protocol" {
	interface SheetQueryMap {
		// The text file's cells as a new .xlsx whose sheet is named after `name` (the file's name without its
		// extension); refused when a workbook cannot hold them.
		"saveAs.xlsx": {
			request: { readonly name: string }
			result: { readonly ok: true; readonly bytes: Uint8Array } | { readonly ok: false; readonly refusal: SheetRefusal }
		}
	}
}

export const saveAsXlsxFeature = defineSheetFeature<TextSheetDocument>({
	id: "saveAs.xlsx",
	queries: {
		"saveAs.xlsx": {
			// A file only partly read is refused: the copy would leave its other rows out.
			run: (context, request): Promise<TextWorkbookResult> =>
				!context.doc.complete
					? Promise.resolve({ ok: false, refusal: { reason: "tooLarge", detail: "incomplete" } })
					: textWorkbookXlsx(context.doc.rows, request.name, {
							...(context.slicing !== undefined && { slicing: context.slicing }),
							...(context.signal !== undefined && { signal: context.signal })
						})
		}
	}
})
