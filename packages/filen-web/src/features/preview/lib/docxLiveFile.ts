import { prepareDocxSave, type DocxPreparedSave } from "@filen/office/core/docx/write"
import { macrotaskYield } from "@filen/office/dom/docx"
import type { DocxRtfDocument } from "@filen/office/dom/docx/rtf"
import type { DocxEditingHost } from "@filen/office/dom/docx/editing"
import type { Slicing } from "@filen/office/core/opc"

// The document editor's document as it holds it now, unsaved edits included, for Download as. Nothing is stored or
// marked saved (nor written at all without an unsaved edit or for a view-only document: the saved file is converted
// then); the sessions are written from stand-ins, so typing meanwhile changes nothing written; the run's signal
// ends the write at its next slice or stage (an AbortError). An OpenDocument text or RTF document: the docx writer's
// parts converted to .docx where the document is held (`convertEdited`), never written in its own format and read
// again (an edit its writer refuses still converts); an .odt target converts on from that .docx. A .docx: the file
// as a save writes it, for the worker to convert.

type DocxConversion = Awaited<ReturnType<DocxRtfDocument["convertEdited"]>>

export type EditedFile =
	// Nothing to write: no unsaved edit, or a document that cannot hold edits (view-only): its saved file is the one.
	| { readonly kind: "saved" }
	// `format`: "docx" when the bytes are the document converted where it is held, not the file's own format.
	// `lossy`: what that .docx lacks, for the next step's report.
	| {
			readonly kind: "file"
			readonly bytes: Uint8Array
			readonly format?: "docx"
			readonly lossy?: Extract<DocxConversion, { kind: "converted" }>["lossy"]
	  }
	| { readonly kind: "converted"; readonly conversion: DocxConversion }

// The document's edits converted to .docx where it is held (rtfEditing.ts, ODT documents).
export type ConvertEdited = (prepared: DocxPreparedSave, signal: AbortSignal) => Promise<DocxConversion>

const SLICE_MS = 12

function abortError(): Error {
	const error = new Error("The write was aborted")

	error.name = "AbortError"

	return error
}

function slicingOf(signal: AbortSignal): Slicing {
	const pause = macrotaskYield()

	return {
		sliceMs: SLICE_MS,
		now: () => performance.now(),
		pause: async () => {
			await pause()

			if (signal.aborted) {
				throw abortError()
			}
		}
	}
}

export async function editedFile(
	host: DocxEditingHost,
	target: "docx" | "odt",
	signal: AbortSignal,
	convertEdited?: ConvertEdited
): Promise<EditedFile> {
	const input = host.saveInput()

	if (input === undefined) {
		throw new Error("The document is not open for editing")
	}

	if (input.context.gate.length > 0 || !input.sessions.some(session => session.snapshot().dirty)) {
		return { kind: "saved" }
	}

	const sessions = input.sessions.map(session => session.withTransaction(session.state.tr))

	if (convertEdited !== undefined) {
		const prepared = await prepareDocxSave({ ...input, sessions }, { slicing: slicingOf(signal) })

		if (signal.aborted) {
			throw abortError()
		}

		const conversion = await convertEdited(prepared, signal)

		// An .odt target converts on from that .docx.
		return target === "odt" && conversion.kind === "converted"
			? { kind: "file", bytes: conversion.bytes, format: "docx", lossy: conversion.lossy }
			: { kind: "converted", conversion }
	}

	return { kind: "file", bytes: (await host.save(sessions, { signal })).bytes }
}
