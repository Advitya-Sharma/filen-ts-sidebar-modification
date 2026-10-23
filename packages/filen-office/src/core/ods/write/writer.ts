import type { InflateHook } from "../../opc/zipReader"
import type { SheetViewOnlyReason } from "../../sheet/protocol"
import type { XlsxFormatWriter } from "../../xlsx/edit/document"
import type { XlsxWorkbookModel } from "../../xlsx/model/types"
import { odsSource, type OdsSource } from "../source"
import { ContentRefused, missingDeclarations } from "./content"
import { saveOds } from "./save"
import { odsEditBegins, odsEditRefusal } from "./check"

// Why an OpenDocument workbook can never be saved here, known at open: content past what is materialised (a repeat
// beyond the caps), tables inside cells, or a prefix the writer's markup needs bound to another namespace.
function viewOnly(source: OdsSource | undefined): SheetViewOnlyReason | null {
	if (source === undefined) return "legacyFormat"
	if (source.tables.some(table => table.truncated)) return "tooLarge"
	if (source.issues.some(issue => issue.kind === "subTable")) return "lossy"

	try {
		missingDeclarations(source.prefixes)
	} catch (error) {
		if (error instanceof ContentRefused) return "lossy"

		throw error
	}

	return null
}

export function odsWriter(workbook: XlsxWorkbookModel, inflate: InflateHook | undefined): XlsxFormatWriter {
	return {
		viewOnly: viewOnly(odsSource(workbook)),
		begin: document => {
			document.workbook.changes.beginJournal?.()
			odsEditBegins(document)
		},
		refused: document => document.workbook.changes.forgetJournal?.(),
		check: (document, touched, command) => odsEditRefusal(document, touched, command),
		save: (document, options) =>
			saveOds(document, {
				untouched: options.untouched,
				...(options.slicing !== undefined && { slicing: options.slicing }),
				...(options.signal !== undefined && { signal: options.signal }),
				...(inflate !== undefined && { inflate })
			})
	}
}
