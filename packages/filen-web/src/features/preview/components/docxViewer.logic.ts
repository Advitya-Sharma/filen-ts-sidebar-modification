import type { DocxErrorReason } from "@filen/office/dom/docx"
import type { DocxDocument } from "@filen/office/react/docx"

// Which engine failures fall back to docx-preview for this release: the engine throwing (a defect or a
// limit it hit) or rejecting markup docx-preview may still read. Files that are not a readable Word
// document at all (encrypted, legacy .doc, not a zip) fail the same way in both and get a message instead.
const FALLBACK_REASONS: ReadonlySet<DocxErrorReason> = new Set<DocxErrorReason>(["failed", "malformed", "invalidPackage"])

// docx-preview renders the whole document on the main thread at once: past these sizes it freezes the tab
// for many seconds and takes gigabytes (a 131 MB document.xml: 37 s, 6.6 GB), so the error shows instead.
// The main part's uncompressed size when the engine read the zip directory, else the file's own size.
export const FALLBACK_MAX_MAIN_PART_BYTES = 24 * 1024 * 1024
export const FALLBACK_MAX_FILE_BYTES = 8 * 1024 * 1024

export type DocxFailure = { readonly fallback: boolean; readonly detail: string } | null

export function docxFailure(document: DocxDocument, fileBytes: number): DocxFailure {
	const error = document.error

	if (document.status !== "error" || error === undefined || !FALLBACK_REASONS.has(error.reason)) {
		return null
	}

	const size = error.mainPartBytes
	const fallback = size === undefined ? fileBytes <= FALLBACK_MAX_FILE_BYTES : size <= FALLBACK_MAX_MAIN_PART_BYTES

	return { fallback, detail: `${error.reason}: ${error.message} (main part ${size === undefined ? "unknown" : String(size)} bytes)` }
}

export type DocxErrorKey = "previewDocxEncrypted" | "previewDocxLegacyFormat" | "previewDocxNotDocument" | "previewDocxLoadFailed"

export function docxErrorKey(reason: DocxErrorReason | undefined): DocxErrorKey {
	switch (reason) {
		case "encrypted":
			return "previewDocxEncrypted"
		case "legacyBinary":
			return "previewDocxLegacyFormat"
		case "notWordprocessing":
		case "compoundFile":
		case "notZip":
			return "previewDocxNotDocument"
		default:
			return "previewDocxLoadFailed"
	}
}
