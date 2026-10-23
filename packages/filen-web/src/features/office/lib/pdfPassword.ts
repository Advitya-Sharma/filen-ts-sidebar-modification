import { PdfWriteError, preparePassword } from "@filen/office/core/pdfwrite"

// Whether a PDF can be encrypted with `password`: the standard security handler takes it only through SASLprep
// (ISO 32000-2 §7.6.4.3.3), so a password it prohibits is refused in the prompt, before any export work starts.
// A platform that cannot check (no Unicode normalization) answers yes; the export then reports it.
export function pdfPasswordAllowed(password: string): boolean {
	try {
		preparePassword(password)

		return true
	} catch (error) {
		return !(error instanceof PdfWriteError && error.reason === "password")
	}
}
