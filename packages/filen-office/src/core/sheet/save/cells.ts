import { writeCsv } from "../csv/csvParse"
import { defineSheetFeature } from "../registry"
import type { TextSheetDocument } from "../text"

// Save as .xlsx of a text file, its first half, where the file is edited: its cells as they now are, in a form
// another worker reads back exactly (saveAs.xlsx converts them there, so the text worker never loads the
// workbook writer).

declare module "../protocol" {
	interface SheetQueryMap {
		// The cells as UTF-8 tab-separated values with a BOM, a value quoted (RFC 4180) when it holds a tab, a quote
		// or a line end, or when it starts the file with a BOM of its own (a reader takes that for the hand-off's):
		// opened as a TSV, it reads back as the same cells. Refused (tooLarge) for a file only partly read: a copy would
		// leave its other rows out.
		"saveAs.cells": {
			request: Record<string, never>
			result: { readonly ok: true; readonly bytes: Uint8Array } | { readonly ok: false; readonly refusal: SheetRefusal }
		}
	}
}

const UTF8_BOM = [0xef, 0xbb, 0xbf] as const

export function textCellsTsv(rows: readonly (readonly string[])[]): Uint8Array {
	let text = writeCsv(rows, "\t", "\n")

	// Unquoted, the first cell holds no quote, tab or line end: quoting it needs no escapes.
	if (text.startsWith("\uFEFF")) {
		const first = rows[0]?.[0] ?? ""

		text = `"${first}"${text.slice(first.length)}`
	}

	const body = new TextEncoder().encode(text)
	const bytes = new Uint8Array(body.length + UTF8_BOM.length)

	bytes.set(UTF8_BOM, 0)
	bytes.set(body, UTF8_BOM.length)

	return bytes
}

export const saveAsCellsFeature = defineSheetFeature<TextSheetDocument>({
	id: "saveAs.cells",
	queries: {
		"saveAs.cells": {
			run: context =>
				context.doc.complete
					? { ok: true, bytes: textCellsTsv(context.doc.rows) }
					: { ok: false, refusal: { reason: "tooLarge", detail: "incomplete" } }
		}
	}
})
