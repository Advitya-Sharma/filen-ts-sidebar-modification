import { toast } from "sonner"
import { bytesSource } from "@filen/office/core/opc"
import type { Slicing } from "@filen/office/core/opc"
import { macrotaskYield, type DocxUnlockedDocument } from "@filen/office/dom/docx"
import {
	DocxPdfExportError,
	docxPdfSnapshot,
	exportDocxPdfInWorker,
	type DocxPdfInput,
	type DocxPdfLive,
	type DocxPdfProgress
} from "@filen/office/dom/docx/pdfClient"
import { pdfDownloadName, pdfErrorKey, pdfProgressText, pdfReportKey, pdfTitle } from "@/features/preview/lib/docxPdf.logic"
import { DOCX_FONT_URLS, docxProtection } from "@/features/preview/lib/docxEngine"
import DocxPdfWorker from "@/features/preview/workers/docxPdf.worker.ts?worker"
import { recoverIfNewerBuild } from "@/lib/appUpdate"
import { downloadBlob } from "@/lib/downloadBlob"
import { i18n } from "@/lib/i18n"
import { log } from "@/lib/log"

// Download as PDF for a Word document: made in a worker of its own (ended with the export), with a toast that
// counts the pages and can cancel; the result is a local download, the file in the drive is unchanged. What
// is exported is what the pages show: the file laid out again in the worker (Final view), or the editor's
// layout with its unsaved edits.

export type DocxPdfSource =
	// The file as downloaded (not password-protected); an OpenDocument text file is laid out as the viewer does.
	| { readonly kind: "file"; readonly bytes: Uint8Array; readonly format?: "docx" | "odt" }
	// A password-protected file's decrypted package, unlocked for this export.
	| { readonly kind: "unlocked"; readonly document: DocxUnlockedDocument }
	// The editor's document, unsaved edits included, in the Final view; `markup`: the editor shows tracked changes or
	// comments, which the PDF leaves out (the toast says so).
	| {
			readonly kind: "editor"
			readonly live: (signal: AbortSignal) => Promise<DocxPdfLive | undefined>
			readonly markup: boolean
	  }

function worker() {
	const instance = new DocxPdfWorker()

	return {
		post(message: unknown, transfer: readonly ArrayBuffer[]) {
			instance.postMessage(message, [...transfer])
		},
		listen(onMessage: (data: unknown) => void, onError: () => void) {
			instance.addEventListener("message", event => {
				onMessage(event.data)
			})
			instance.addEventListener("error", () => {
				// A deploy that changed this worker removes its old file (appUpdate.ts).
				recoverIfNewerBuild()
				onError()
			})
			instance.addEventListener("messageerror", onError)
		},
		terminate() {
			instance.terminate()
		}
	}
}

// The editor's snapshot is built on this thread: a macrotask turn (a message, which browsers do not clamp as
// they clamp nested timers) once each slice is up.
const SNAPSHOT_SLICE_MS = 8

function snapshotSlicing(): Slicing {
	return { pause: macrotaskYield(), sliceMs: SNAPSHOT_SLICE_MS, now: () => performance.now() }
}

async function inputOf(source: DocxPdfSource, fileName: string, signal: AbortSignal): Promise<DocxPdfInput> {
	switch (source.kind) {
		case "file":
			// A copy of its own: the preview cache keeps the file's buffer.
			return { kind: "bytes", bytes: source.bytes.slice(), fontUrls: DOCX_FONT_URLS }
		case "unlocked":
			return { kind: "bytes", bytes: (await source.document.take({ signal })).bytes, fontUrls: DOCX_FONT_URLS }
		case "editor": {
			const live = await source.live(signal)

			if (live === undefined) {
				throw new DocxPdfExportError("openFailed", "The document is not open")
			}

			return {
				kind: "snapshot",
				document: await docxPdfSnapshot(live, { copy: true, title: pdfTitle(fileName), signal, slicing: snapshotSlicing() })
			}
		}
	}
}

let exports = 0

// Makes the PDF and downloads it as "<name>.pdf". `password` encrypts it (a protected document's own).
export async function downloadDocxPdf(source: DocxPdfSource, fileName: string, password?: string): Promise<void> {
	const id = `docx-pdf-${String(++exports)}`
	const controller = new AbortController()
	const action = {
		label: i18n.t("preview:previewDocxPdfCancel"),
		onClick: () => {
			controller.abort()
		}
	}
	const show = (progress: DocxPdfProgress | null): void => {
		const text = pdfProgressText(progress)

		toast.loading(
			text.key === "previewDocxPdfProgress"
				? i18n.t("preview:previewDocxPdfProgress", { done: text.done, total: text.total })
				: i18n.t("preview:previewDocxPdfPreparing"),
			{
				id,
				action,
				duration: Number.POSITIVE_INFINITY,
				...(source.kind === "editor" && source.markup ? { description: i18n.t("preview:previewDocxPdfFinalOnly") } : {})
			}
		)
	}

	show(null)

	try {
		const input = await inputOf(source, fileName, controller.signal)
		const result = await exportDocxPdfInWorker(worker, {
			input,
			options: {
				title: pdfTitle(fileName),
				lang: i18n.language,
				...(source.kind === "file" && source.format !== undefined ? { format: source.format } : {}),
				...(password === undefined ? {} : { password })
			},
			signal: controller.signal,
			onProgress: show
		})

		downloadBlob(pdfDownloadName(fileName), result.blob as Blob)
		toast.dismiss(id)

		const reportKey = pdfReportKey(result.report)

		if (reportKey !== null) {
			log.warn("docx", "pdf export simplified", { missing: result.report.missing, refused: result.report.refused })
			toast.warning(i18n.t(`preview:${reportKey}`))
		}
	} catch (error) {
		const reason = error instanceof DocxPdfExportError ? error.reason : controller.signal.aborted ? "aborted" : "failed"
		const key = pdfErrorKey(reason)

		if (key === null) {
			toast.dismiss(id)

			return
		}

		if (reason === "failed" || reason === "workerFailed" || reason === "openFailed") log.error("docx", "pdf export failed", error)

		toast.dismiss(id)
		toast.error(i18n.t(`preview:${key}`))
	} finally {
		if (source.kind === "unlocked") source.document.release()
	}
}

// A password-protected file's password checked for its PDF: the file is unlocked again with it (where the
// docx engine runs; nothing decrypted reaches this thread). Accepted: the unlocked document, to export from or
// release. The password itself is never kept.
export async function unlockForPdf(
	fileBytes: Uint8Array,
	password: string,
	signal: AbortSignal
): Promise<
	{ readonly verdict: "accepted"; readonly document: DocxUnlockedDocument } | { readonly verdict: "wrong" | "aborted" | "failed" }
> {
	const opened = await docxProtection.open(bytesSource(fileBytes), { signal })

	if (opened.kind === "aborted") {
		return { verdict: "aborted" }
	}

	if (opened.kind !== "needsPassword") {
		return { verdict: "failed" }
	}

	const unlocked = await opened.locked.unlock(password, { signal })

	if (unlocked.ok) {
		return { verdict: "accepted", document: unlocked.document }
	}

	opened.locked.dispose()

	return { verdict: unlocked.reason === "wrongPassword" ? "wrong" : unlocked.reason === "aborted" ? "aborted" : "failed" }
}
