import { toast } from "sonner"
import { driveItemName, isAbortError } from "@filen/shared"
import { isCompoundFile } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import type { SheetFileKind, SheetOpened } from "@filen/office/core/sheet"
import type { DocxConvertedDocument, DocxEncryptionRefusal, DocxLegacyRefusal, DocxUnlockedDocument } from "@filen/office/dom/docx"
import type { DocxPdfProgress } from "@filen/office/dom/docx/pdfClient"
import { recordContentVerdict } from "@/features/drive/lib/contentVerdicts"
import { driveItemMime, type DriveItem } from "@/features/drive/lib/item"
import { narrowToAnyFile } from "@/features/drive/lib/download"
import { extensionOf, itemTypeExtension } from "@/features/drive/lib/preview.logic"
import { DOWNLOAD_AS_TARGETS, downloadAsSource, type DownloadAsTarget } from "@/features/drive/lib/downloadAs.logic"
import { askPassword } from "@/features/drive/lib/downloadAsPassword"
import { writeZipOf, zipDelivery, zipEntryNamer, type ZipDelivery, type ZipFile } from "@/features/drive/lib/downloadAsZip"
import { isPickerCancelled } from "@/features/drive/lib/saveDownload"
import { pdfPasswordAllowed } from "@/features/office/lib/pdfPassword"
import {
	docxRefusalKey,
	docxFormatOf,
	isTextTemplate,
	legacyNeedsPassword,
	legacyRefusalKey
} from "@/features/preview/components/docxViewer.logic"
import { docxLegacy, docxProtection, docxRtf } from "@/features/preview/lib/docxEngine"
import { docxPdfFailure, makeDocxPdf, reportDocxPdf, type DocxPdfSource } from "@/features/preview/lib/docxPdf"
import { pdfProgressText } from "@/features/preview/lib/docxPdf.logic"
import type { PreviewAccessMode } from "@/features/preview/lib/accessMode"
import { joinPreviewBytes } from "@/features/preview/lib/previewCache"
import { openSheet } from "@/features/spreadsheet/lib/sheetHosts"
import { SheetEventRelay } from "@/features/spreadsheet/lib/sheetEvents"
import { withWorkbookDigit } from "@/features/spreadsheet/lib/sheetDigitWidth"
import { sheetKindOf } from "@/features/spreadsheet/lib/sheetKind"
import { browserLanguage, sheetLocaleOf } from "@/features/spreadsheet/formulaBar/sheetLocale"
import { formOf, pdfRequest, printRequests } from "@/features/spreadsheet/print/print.logic"
import { sheetSettings } from "@/features/spreadsheet/print/printSettings"
import { makeSheetPdf, sheetPdfMessages } from "@/features/spreadsheet/save/sheetPdf"
import { fileStem } from "@/features/spreadsheet/save/sheetSave.logic"
import { runOp } from "@/lib/actions/outcome"
import { downloadBlob } from "@/lib/downloadBlob"
import { i18n, type DriveKey, type PreviewKey } from "@/lib/i18n"
import { errorLabel } from "@/lib/i18n/errorLabel"
import { log } from "@/lib/log"
import { asErrorDTO } from "@/lib/sdk/errors"
import { sdkApi } from "@/lib/sdk/client"

// Download as from the item menus: each file exported from its saved bytes without opening its viewer, one after
// another under one toast whose Cancel stops the download and the worker job in flight (and skips the rest). Several
// files come as one zip, as the drive's Download delivers them, and what went wrong is told once, after the run. The
// bytes come as the viewer's do (an open preview's buffer is reused, else one streamed read of the file); the work
// runs where the viewers' does: the document engine's worker converts Word 97–2003 and RTF files and unlocks
// protected ones, the PDF worker lays the document out (ended with the export), the sheet worker opens the workbook
// and prints it (released once idle). A protected file asks its password first, which then protects the PDF. What
// comes out equals the editors' File menu export of the same saved file.

// Where the menu was opened: a public link's files are read as the viewer reads them there.
export interface DownloadAsContext {
	readonly accessMode: PreviewAccessMode
	readonly cacheScope: string | null
}

type Message =
	| { readonly key: `preview:${PreviewKey}` | `drive:${DriveKey}`; readonly values: Readonly<Record<string, number>> }
	// Already in words (an error's own label).
	| { readonly text: string }

type Exported =
	| { readonly kind: "pdf"; readonly blob: Blob; readonly warnings: readonly Message[] }
	| { readonly kind: "failed"; readonly error: Message }
	// The user let the password prompt go: nothing to report.
	| { readonly kind: "skipped" }

type Phase =
	| { readonly kind: "downloading" }
	| { readonly kind: "document"; readonly progress: DocxPdfProgress | null }
	| { readonly kind: "sheet"; readonly done: number | null; readonly total: number }

interface Run {
	readonly signal: AbortSignal
	readonly show: (phase: Phase) => void
}

class Aborted extends Error {
	constructor() {
		super("Download as was cancelled")
		this.name = "AbortError"
	}
}

function failed(key: PreviewKey, values: Readonly<Record<string, number>> = {}): Exported {
	return { kind: "failed", error: { key: `preview:${key}`, values } }
}

function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
	if (signal.aborted) {
		return Promise.reject(new Aborted())
	}

	return new Promise((resolve, reject) => {
		const onAbort = (): void => {
			reject(new Aborted())
		}

		signal.addEventListener("abort", onAbort, { once: true })
		promise.then(
			value => {
				signal.removeEventListener("abort", onAbort)
				resolve(value)
			},
			(error: unknown) => {
				signal.removeEventListener("abort", onAbort)
				reject(error instanceof Error ? error : new Error(String(error)))
			}
		)
	})
}

// The file's bytes: a buffer the preview already holds or is loading (shared: `owned` false, copied before it is
// handed to a worker), else a read of its own, cancelled with the export.
async function fileBytes(item: DriveItem, context: DownloadAsContext, signal: AbortSignal): Promise<{ bytes: Uint8Array; owned: boolean }> {
	const joined = await abortable(joinPreviewBytes(context.cacheScope, item.data.uuid), signal)

	if (joined !== undefined) {
		return { bytes: joined, owned: false }
	}

	const file = narrowToAnyFile(item)
	const token = crypto.randomUUID()
	const cancel = (): void => {
		void sdkApi.cancelPreviewDownload(token)
	}

	signal.addEventListener("abort", cancel, { once: true })

	try {
		const bytes = await runOp(
			context.accessMode === "anon" ? sdkApi.downloadLinkedFileBytesAnon(file, token) : sdkApi.downloadFileBytes(file, token)
		)

		if (signal.aborted) {
			throw new Aborted()
		}

		return { bytes, owned: true }
	} finally {
		signal.removeEventListener("abort", cancel)
	}
}

async function documentPdf(source: DocxPdfSource, name: string, run: Run, password?: string): Promise<Exported> {
	run.show({ kind: "document", progress: null })

	try {
		const result = await makeDocxPdf(source, name, {
			signal: run.signal,
			password,
			onProgress: progress => {
				run.show({ kind: "document", progress })
			}
		})
		const warning = reportDocxPdf(result)

		return { kind: "pdf", blob: result.blob as Blob, warnings: warning === null ? [] : [{ key: `preview:${warning}`, values: {} }] }
	} catch (error) {
		const key = docxPdfFailure(error, run.signal.aborted)

		if (key === null) {
			throw new Aborted()
		}

		return failed(key)
	}
}

// A Word 97–2003 document's package, converted in the document engine's worker (taken, then dropped there).
async function convertedPackage(document: DocxConvertedDocument, signal: AbortSignal): Promise<Uint8Array> {
	try {
		return await document.take({ signal })
	} finally {
		document.release()
	}
}

// What a password check in the worker found: opened (`value`, which the caller releases if it goes unused), a wrong
// password, a cancelled check, or a file no password opens (`refusal`).
type Checked<T, R> =
	| { readonly kind: "opened"; readonly value: T }
	| { readonly kind: "wrong" }
	| { readonly kind: "aborted" }
	| { readonly kind: "refused"; readonly refusal: R }

type Prompted<T, R> =
	| { readonly kind: "accepted"; readonly value: T; readonly password: string }
	// "password": the password verified as typed, but a PDF cannot be protected with it (SASLprep), checked first.
	| { readonly kind: "refused"; readonly refusal: R | "password" }
	| { readonly kind: "skipped" }

// Asks for the file's password (the shared prompt) until `check` opens the file with one; the password then protects
// the PDF. A password the PDF cannot take is refused before any work starts, as in the editors' prompts.
async function promptFor<T, R>(
	name: string,
	run: Run,
	check: (password: string, signal: AbortSignal) => Promise<Checked<T, R>>,
	release: (value: T) => void,
	// What is held for the locked file, let go unless it opened (then it goes with what opened).
	dispose: () => void
): Promise<Prompted<T, R>> {
	const state: { found: { value: T; password: string } | null; refusal: R | "password" | null } = { found: null, refusal: null }
	const accepted = await askPassword(
		name,
		async (typed, signal) => {
			if (!pdfPasswordAllowed(typed)) {
				state.refusal = "password"

				return "refused"
			}

			const answer = await check(typed, signal)

			switch (answer.kind) {
				case "wrong":
				case "aborted":
					return answer.kind
				case "refused":
					state.refusal = answer.refusal

					return "refused"
				case "opened":
					// The prompt or the export went while the check finished: nothing is exported.
					if (signal.aborted || run.signal.aborted) {
						release(answer.value)

						return "aborted"
					}

					state.found = { value: answer.value, password: typed }

					return "accepted"
			}
		},
		run.signal
	)

	if (accepted && state.found !== null) {
		return { kind: "accepted", ...state.found }
	}

	dispose()

	if (state.refusal !== null) {
		return { kind: "refused", refusal: state.refusal }
	}

	if (run.signal.aborted) {
		throw new Aborted()
	}

	return { kind: "skipped" }
}

// A Word 97–2003 document, converted in the document engine's worker as the viewer converts it (an encrypted one
// converted again with its password, each try).
async function legacyPdf(bytes: Uint8Array, name: string, run: Run): Promise<Exported> {
	const opened = await docxLegacy.open(bytesSource(bytes), { signal: run.signal })

	if (opened.kind === "aborted") {
		throw new Aborted()
	}

	if (opened.kind === "converted") {
		return documentPdf({ kind: "file", bytes: await convertedPackage(opened.document, run.signal), owned: true }, name, run)
	}

	if (!legacyNeedsPassword(opened.refusal)) {
		return failed(legacyRefusalKey(opened.refusal))
	}

	const prompted = await promptFor<DocxConvertedDocument, DocxLegacyRefusal>(
		name,
		run,
		async (password, signal) => {
			const answer = await docxLegacy.open(bytesSource(bytes), { signal, password })

			if (answer.kind === "converted") {
				return { kind: "opened", value: answer.document }
			}

			if (answer.kind === "aborted") {
				return { kind: "aborted" }
			}

			return answer.refusal.reason === "wrongPassword" ? { kind: "wrong" } : { kind: "refused", refusal: answer.refusal }
		},
		document => {
			document.release()
		},
		() => undefined
	)

	switch (prompted.kind) {
		case "accepted":
			return documentPdf(
				{ kind: "file", bytes: await convertedPackage(prompted.value, run.signal), owned: true },
				name,
				run,
				prompted.password
			)
		case "refused":
			return failed(prompted.refusal === "password" ? "previewDocxPdfPasswordUnsupported" : legacyRefusalKey(prompted.refusal))
		case "skipped":
			return prompted
	}
}

// A document laid out as its viewer lays it out (docxViewer.tsx): an RTF file read into its package, a compound
// file probed for a password (then unlocked) or converted as a Word 97–2003 document, anything else as it is.
async function exportDocument(item: DriveItem, bytes: Uint8Array, owned: boolean, context: DownloadAsContext, run: Run): Promise<Exported> {
	const name = driveItemName(item)
	const format = docxFormatOf(name)

	// A `.dot` the drive could not yet tell apart, read here as its viewer reads it: a text file is no document. What
	// it is is recorded as the viewer records it, so the menus stop offering it.
	if (isTextTemplate(item, bytes)) {
		recordContentVerdict({ uuid: item.data.uuid, name, mime: driveItemMime(item), size: item.data.size }, "notWordBinary", {
			memoryOnly: context.accessMode === "anon"
		})

		return { kind: "failed", error: { key: "drive:driveDownloadAsNotDocument", values: {} } }
	}

	run.show({ kind: "document", progress: null })

	if (format === "rtf") {
		const opened = await docxRtf.open(bytes, { signal: run.signal })

		if (opened.kind === "aborted") {
			throw new Aborted()
		}

		if (opened.kind === "notRtf") {
			return failed("previewDocxPdfFailed")
		}

		let docx: Uint8Array

		try {
			docx = await opened.document.take()
		} finally {
			opened.document.release()
		}

		return documentPdf({ kind: "file", bytes: docx, owned: true }, name, run)
	}

	if (!isCompoundFile(bytes)) {
		return documentPdf({ kind: "file", bytes, format, owned }, name, run)
	}

	const probed = await docxProtection.open(bytesSource(bytes), { signal: run.signal })

	switch (probed.kind) {
		case "aborted":
			throw new Aborted()
		case "refused":
			return failed(docxRefusalKey(probed.reason))
		case "plain":
			return legacyPdf(bytes, name, run)
		case "needsPassword":
			break
	}

	const locked = probed.locked
	const prompted = await promptFor<DocxUnlockedDocument, DocxEncryptionRefusal>(
		name,
		run,
		async (password, signal) => {
			const result = await locked.unlock(password, { signal })

			if (result.ok) {
				return { kind: "opened", value: result.document }
			}

			return result.reason === "wrongPassword" || result.reason === "aborted"
				? { kind: result.reason === "wrongPassword" ? "wrong" : "aborted" }
				: { kind: "refused", refusal: result.reason }
		},
		document => {
			document.release()
		},
		() => {
			locked.dispose()
		}
	)

	switch (prompted.kind) {
		case "accepted":
			return documentPdf({ kind: "unlocked", document: prompted.value }, name, run, prompted.password)
		case "refused":
			return failed(prompted.refusal === "password" ? "previewDocxPdfPasswordUnsupported" : docxRefusalKey(prompted.refusal))
		case "skipped":
			return prompted
	}
}

// The workbook opened in its sheet worker as its viewer opens it (useSheetSession.ts), its active sheet printed as
// the File menu's Download as ▸ PDF prints it (sheetBody.tsx) once the viewer opened it.
async function exportSpreadsheet(item: DriveItem, bytes: Uint8Array, owned: boolean, name: string, run: Run): Promise<Exported> {
	const { kind } = sheetKindOf(extensionOf(name), bytes, itemTypeExtension(item) === "tsv")
	const events = new SheetEventRelay()

	run.show({ kind: "sheet", done: null, total: 0 })

	const opened = await openSheet(owned ? bytes : bytes.slice(), { kind, signal: run.signal }, events.emit)

	switch (opened.kind) {
		case "aborted":
			throw new Aborted()
		case "failed":
			return failed("previewSpreadsheetUnreadable")
		case "refused":
			return failed(docxRefusalKey(opened.reason))
		case "opened":
			return printWorkbook(opened, events, kind, name, run)
		case "needsPassword":
			break
	}

	const locked = opened.locked
	const prompted = await promptFor<SheetOpened, DocxEncryptionRefusal>(
		name,
		run,
		async (password, signal) => {
			const answer = await locked.unlock(password, { signal }).catch(() => null)

			if (answer === null) {
				return { kind: "aborted" }
			}

			if (answer.ok) {
				return { kind: "opened", value: answer }
			}

			return answer.reason === "wrongPassword" || answer.reason === "aborted"
				? { kind: answer.reason === "wrongPassword" ? "wrong" : "aborted" }
				: { kind: "refused", refusal: answer.reason }
		},
		workbook => {
			workbook.session.close()
		},
		() => {
			locked.dispose()
		}
	)

	switch (prompted.kind) {
		case "accepted":
			return printWorkbook(prompted.value, events, kind, name, run, prompted.password)
		case "refused":
			return failed(prompted.refusal === "password" ? "previewSheetPdfPasswordUnsupported" : docxRefusalKey(prompted.refusal))
		case "skipped":
			return prompted
	}
}

async function printWorkbook(
	opened: SheetOpened,
	events: SheetEventRelay,
	kind: SheetFileKind,
	name: string,
	run: Run,
	password?: string
): Promise<Exported> {
	const { session } = opened

	try {
		const { info } = await withWorkbookDigit(opened, events)

		// The locale the viewer's grid has its session show numbers and dates in (useSheetFormula.tsx).
		session.configure({ locale: sheetLocaleOf(browserLanguage()) })
		const sheet = info.sheets.find(entry => entry.id === info.activeSheet)
		const active = sheet === undefined ? null : await sheetSettings({ session, layout: () => null }, kind, sheet)

		if (active === null) {
			return failed("previewSheetPdfNothing")
		}

		const requests = printRequests(formOf(active.setup), active, [active], []) ?? []
		const request = pdfRequest(
			requests,
			{ name, title: fileStem(name), lang: browserLanguage() },
			{ kind, sheets: info.sheets, date1904: info.date1904, showFormulas: false }
		)
		const result = await makeSheetPdf(session, request, {
			signal: run.signal,
			password,
			onProgress: (done, total) => {
				run.show({ kind: "sheet", done, total })
			}
		})

		if (run.signal.aborted) {
			throw new Aborted()
		}

		const { warnings, error } = sheetPdfMessages(result)

		if (result.ok) {
			return {
				kind: "pdf",
				blob: result.blob,
				warnings: warnings.map(warning => ({ key: `preview:${warning.key}` as const, values: warning.values }))
			}
		}

		return failed(error?.key ?? "previewSheetPdfFailed", error?.values ?? {})
	} catch (error) {
		if (run.signal.aborted) {
			throw new Aborted()
		}

		log.error("preview", "spreadsheet pdf failed", error)

		return failed("previewSheetPdfFailed")
	} finally {
		session.close()
	}
}

function phaseText(phase: Phase): string {
	switch (phase.kind) {
		case "downloading":
			return i18n.t("drive:driveDownloadAsDownloading")
		case "document": {
			const text = pdfProgressText(phase.progress)

			return text.key === "previewDocxPdfProgress"
				? i18n.t("preview:previewDocxPdfProgress", { done: text.done, total: text.total })
				: i18n.t("preview:previewDocxPdfPreparing")
		}
		case "sheet":
			return phase.done === null
				? i18n.t("preview:previewSheetPdfPreparing")
				: i18n.t("preview:previewSheetPdfProgress", { done: phase.done, total: phase.total })
	}
}

async function exportOne(item: DriveItem, context: DownloadAsContext, run: Run): Promise<Exported> {
	const source = downloadAsSource(item)

	if (source === null) {
		return failed("previewDocxPdfFailed")
	}

	const name = driveItemName(item)

	run.show({ kind: "downloading" })

	const { bytes, owned } = await fileBytes(item, context, run.signal)

	return source === "document" ? exportDocument(item, bytes, owned, context, run) : exportSpreadsheet(item, bytes, owned, name, run)
}

let runs = 0

// What a run tells of one file once it ends: why it gave no PDF, or what its PDF lacks (never the PDF itself, which
// a bulk run lets go once written).
interface Outcome {
	readonly name: string
	readonly failed: Message | null
	readonly warnings: readonly Message[]
}

// Files a bulk run's report names, the rest counted.
const REPORT_LINES = 5

function messageText(message: Message): string {
	return "text" in message ? message.text : i18n.t(message.key, message.values)
}

// One report for a bulk run: the files that failed (else those whose PDF lacks something), each with why.
function reportBulk(outcomes: readonly Outcome[]): void {
	const failed = outcomes.filter(outcome => outcome.failed !== null)
	const warned = outcomes.filter(outcome => outcome.warnings.length > 0)
	const listed = [...failed, ...warned]

	if (listed.length === 0) {
		return
	}

	const lines = listed
		.slice(0, REPORT_LINES)
		.map(({ name, failed: error, warnings }) => `${name}: ${(error === null ? warnings : [error]).map(messageText).join(" ")}`)

	if (listed.length > REPORT_LINES) {
		lines.push(i18n.t("drive:driveDownloadAsReportMore", { count: listed.length - REPORT_LINES }))
	}

	const options = { description: lines.join("\n"), classNames: { description: "whitespace-pre-line" } }

	if (failed.length > 0) {
		toast.error(i18n.t("drive:driveDownloadAsReportFailed", { count: failed.length, total: outcomes.length }), options)
	} else {
		toast.warning(i18n.t("drive:driveDownloadAsReportWarned", { count: warned.length }), options)
	}
}

interface Exports {
	readonly signal: AbortSignal
	readonly id: string
	readonly action: { readonly label: string; readonly onClick: () => void }
	readonly outcomes: Outcome[]
}

// Read through a call: an earlier check would narrow the flag for the rest of the turn.
function cancelled(signal: AbortSignal): boolean {
	return signal.aborted
}

// Exports `items` one after another under the run's toast, each file's PDF handed out as it is made (the next file
// is started once the caller asks for it); what became of each goes to `outcomes`. Ends early once cancelled.
async function* exportPdfs(
	items: readonly DriveItem[],
	context: DownloadAsContext,
	run: Exports
): AsyncGenerator<{ name: string; blob: Blob }> {
	for (const [index, item] of items.entries()) {
		if (run.signal.aborted) {
			return
		}

		const name = driveItemName(item)
		const description =
			items.length > 1 ? i18n.t("drive:driveDownloadAsBulkItem", { name, index: index + 1, count: items.length }) : name
		let exported: Exported

		try {
			exported = await exportOne(item, context, {
				signal: run.signal,
				show: phase => {
					toast.loading(phaseText(phase), { id: run.id, action: run.action, description, duration: Number.POSITIVE_INFINITY })
				}
			})
		} catch (error) {
			if (cancelled(run.signal)) {
				return
			}

			log.error("drive", "download as failed", error)
			exported = {
				kind: "failed",
				error: error instanceof Error ? { key: "preview:previewDocxPdfFailed", values: {} } : { text: errorLabel(error) }
			}
		}

		switch (exported.kind) {
			case "pdf":
				run.outcomes.push({ name, failed: null, warnings: exported.warnings })
				yield { name, blob: exported.blob }
				break
			case "failed":
				run.outcomes.push({ name, failed: exported.error, warnings: [] })
				break
			case "skipped":
				break
		}
	}
}

// Exports `items` to `target` one after another; one file downloads as "<name>.<extension>", several as one zip of
// them, saved as the drive's Download saves several files. Call it synchronously off the click (the save picker
// needs the gesture). Resolves once done, failed or cancelled; never rejects.
export async function startDownloadAs(items: readonly DriveItem[], target: DownloadAsTarget, context: DownloadAsContext): Promise<void> {
	const controller = new AbortController()
	const cancel = (): void => {
		controller.abort()
	}
	const run: Exports = {
		signal: controller.signal,
		id: `download-as-${String(++runs)}`,
		action: { label: i18n.t("preview:previewDocxPdfCancel"), onClick: cancel },
		outcomes: []
	}
	const extension = DOWNLOAD_AS_TARGETS[target].extension
	// Asked first, while the click still counts as one.
	// The archive's download ended before the archive did: the browser's own Cancel ends the run as the toast's
	// does; anything else (a failed or stalled download) ends it with why.
	const onEnded = (error: unknown): void => {
		if (!controller.signal.aborted) {
			controller.abort(isAbortError(error) || asErrorDTO(error).kind === "Cancelled" ? undefined : new DeliveryFailed(error))
		}
	}
	const delivery =
		items.length > 1 ? zipDelivery(i18n.t("transfers:transfersZipDownloadDefaultName"), controller.signal, onEnded, downloadBlob) : null

	if (delivery === null) {
		for await (const pdf of exportPdfs(items, context, run)) {
			downloadBlob(`${fileStem(pdf.name)}.${extension}`, pdf.blob)
		}

		toast.dismiss(run.id)

		const [only] = run.outcomes

		if (!controller.signal.aborted && only !== undefined) {
			if (only.failed !== null) {
				toast.error(messageText(only.failed), { description: only.name })
			}

			for (const warning of only.warnings) {
				toast.warning(messageText(warning), { description: only.name })
			}
		}

		return
	}

	let zip: ZipDelivery

	try {
		zip = await delivery
	} catch (error) {
		if (!isPickerCancelled(error)) {
			log.error("drive", "download as could not start its zip", error)
			toast.error(errorLabel(error))
		}

		return
	}

	const pdfs = exportPdfs(items, context, run)
	const name = zipEntryNamer(extension)

	try {
		const first = await pdfs.next()

		if (first.done === true) {
			// No PDF came of it: nothing is saved.
			await zip.fail(new Error("nothing to save"))
		} else {
			async function* files(): AsyncGenerator<ZipFile> {
				let next: IteratorResult<{ name: string; blob: Blob }> = first

				while (next.done !== true) {
					yield { name: await name(next.value.name), blob: next.value.blob, modified: new Date() }
					next = await pdfs.next()
				}
			}

			// A cancelled run's archive is never finished (its end records never written).
			await writeZipOf(files(), {
				write: chunk => {
					controller.signal.throwIfAborted()

					return zip.sink.write(chunk)
				},
				close: () => {
					controller.signal.throwIfAborted()

					return zip.sink.close()
				}
			})
			await zip.saved()
		}
	} catch (error) {
		await zip.fail(error)

		if (!controller.signal.aborted) {
			controller.abort(new DeliveryFailed(error))
		}
	} finally {
		await pdfs.return(undefined)
	}

	toast.dismiss(run.id)

	const reason: unknown = controller.signal.reason

	if (reason instanceof DeliveryFailed) {
		log.error("drive", "download as zip failed", reason.error)
		toast.error(errorLabel(reason.error))
	} else if (!controller.signal.aborted) {
		reportBulk(run.outcomes)
	}
}

// Why a bulk run ended that was not the user's Cancel: its archive could not be saved.
class DeliveryFailed {
	readonly error: unknown

	constructor(error: unknown) {
		this.error = error
	}
}
