import type { SheetFileKind } from "@filen/office/core/sheet"
import type { LiveConversion } from "@/features/drive/lib/downloadAs"
import type { SheetHandle } from "@/features/spreadsheet/hooks/useSheetSession"
import { isTextSheetKind, openSheet, textToXlsx } from "@/features/spreadsheet/lib/sheetHosts"
import { fileStem } from "@/features/spreadsheet/save/sheetSave.logic"

type LiveAnswer = Awaited<ReturnType<LiveConversion>>

// Download as ▸ .xlsx / .ods from the open spreadsheet, unsaved edits included: the cell being typed committed first
// (as Download as CSV and PDF commit it), then a workbook converted by its own session in the workbook worker, or a
// text file's cells (the saveAs.cells query) made a workbook as Save as .xlsx makes it, converted on to .ods in a
// workbook session of its own. None for a password-protected workbook: its copy could not keep the protection (the
// saved-file path refuses it with why).
export function liveConversion(
	handle: SheetHandle,
	kind: SheetFileKind,
	name: string,
	commit: () => Promise<boolean>
): LiveConversion | undefined {
	if (handle.info.encrypted !== undefined) {
		return undefined
	}

	return async (format, signal): Promise<LiveAnswer> => {
		if (!(await commit())) {
			return { ok: false, reason: "uncommitted", message: "The cell being edited holds an entry that cannot be committed" }
		}

		if (!isTextSheetKind(kind)) {
			return unlessAborted(await handle.session.query("saveAs.format", { format }, { signal }), signal)
		}

		const { bytes } = await handle.session.query("saveAs.cells", {}, { signal })
		// The workbook worker takes no signal for this step: a Cancel meanwhile is read once it answers.
		const xlsx = await textToXlsx(bytes, fileStem(name))

		if (signal.aborted) {
			return ABORTED
		}

		if (!xlsx.ok) {
			return { ok: false, reason: "unsupported", message: "The cells could not be made a workbook" }
		}

		return format === "xlsx" ? { ok: true, bytes: xlsx.bytes, lossy: [] } : unlessAborted(await odsOf(xlsx.bytes, signal), signal)
	}
}

const ABORTED: LiveAnswer = { ok: false, reason: "aborted", message: "The conversion was cancelled" }

// Nothing goes out once cancelled, whatever the worker answered.
function unlessAborted(answer: LiveAnswer, signal: AbortSignal): LiveAnswer {
	return signal.aborted ? ABORTED : answer
}

interface EntryModel {
	commitEditing(next: null): Promise<boolean>
	getSnapshot(): { readonly editing: { readonly invalid?: boolean } | null }
}

// Commits the cell being typed before an export reads the cells (PDF, CSV, another format). When it cannot be
// committed the export stops, and the sheet tells why: its refusal toast, the data validation alert, or, for a
// formula it cannot read (the entry reopened marked, its note only beside the cell contents), `sayInvalid` with the
// sheet's own entry message.
export async function commitForExport(model: EntryModel, sayInvalid: () => void): Promise<boolean> {
	if (await model.commitEditing(null)) {
		return true
	}

	if (model.getSnapshot().editing?.invalid === true) {
		sayInvalid()
	}

	return false
}

// An .xlsx converted to .ods in a workbook session of its own (bytes transferred), closed once done.
async function odsOf(xlsx: Uint8Array, signal: AbortSignal): Promise<LiveAnswer> {
	const opened = await openSheet(xlsx, { kind: "xlsx", signal }, () => undefined)

	if (opened.kind !== "opened") {
		if (opened.kind === "needsPassword") {
			opened.locked.dispose()
		}

		return {
			ok: false,
			reason: opened.kind === "aborted" ? "aborted" : "unsupported",
			message: `The workbook did not open (${opened.kind})`
		}
	}

	try {
		return await opened.session.query("saveAs.format", { format: "ods" }, { signal })
	} finally {
		opened.session.close()
	}
}
